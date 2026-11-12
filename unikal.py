#!/usr/bin/env python3
"""UniKalender – Logik-Teil.

Liest den Excel-Stundenplan der Leibniz-FH (nur Standardbibliothek, kein openpyxl nötig),
wertet Uni-Mails nach Ausfällen aus und erzeugt die Soll-Liste der Kalendertermine.
Der Swift-Helfer `unikal` kümmert sich um Kalender- und Mail-Zugriff.

Befehle:
  python3 unikal.py parse <datei.xlsx> [--group A1/2]      -> Termine als Tabelle (zum Prüfen)
  python3 unikal.py plan --config c.json --mails m.json --state s.json
                                                          -> JSON {events, notes} auf stdout
"""
import argparse
import datetime as dt
import hashlib
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

COLOR_RED = "CC0000"      # Änderung / gestrichen
COLOR_PURPLE = "A02B93"   # Online
COLOR_GREEN = "3C7D22"    # Prüfung

TIME_RE = re.compile(r"\(?\s*(\d{1,2})[:.](\d{2})\s*[-–]\s*(\d{1,2})[:.](\d{2})\s*\)?")
GROUP_RE = re.compile(r"\(\s*(A\d+/\d+)\s*\)")
DATE_HDR_RE = re.compile(r"(\d{2})\.(\d{2})\.(\d{4})")
SLOT_NAMES = {"vormittags": "vormittag", "nachmittags": "nachmittag", "abends": "abend", "abgaben": "abgabe"}


# --------------------------------------------------------------------------- Excel lesen

def _col_row(ref):
    m = re.match(r"([A-Z]+)(\d+)", ref)
    return m.group(1), int(m.group(2))


def _attrs_from_props(el):
    """rPr- oder font-Element -> dict(strike, bold, color)."""
    a = {"strike": False, "bold": False, "color": ""}
    if el is None:
        return a
    for ch in el:
        tag = ch.tag.replace(M, "")
        val = ch.get("val")
        on = val not in ("0", "false")
        if tag == "strike":
            a["strike"] = on
        elif tag == "b":
            a["bold"] = on
        elif tag == "color" and ch.get("rgb"):
            a["color"] = ch.get("rgb")[-6:].upper()
    return a


def read_sheet(path):
    """Gibt {(col,row): [(char, attrs), ...]} und die Rohtexte zurück."""
    z = zipfile.ZipFile(path)
    names = z.namelist()
    styles = ET.fromstring(z.read("xl/styles.xml"))
    fonts = [_attrs_from_props(f) for f in styles.find(M + "fonts")]
    xfs = [int(x.get("fontId", 0)) for x in styles.find(M + "cellXfs")]

    shared = []
    if "xl/sharedStrings.xml" in names:
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(M + "si"):
            runs = si.findall(M + "r")
            if runs:
                shared.append([((r.find(M + "t").text or ""), _attrs_from_props(r.find(M + "rPr"))) for r in runs])
            else:
                t = si.find(M + "t")
                shared.append([((t.text if t is not None else "") or "", None)])

    # erstes Arbeitsblatt
    sheet_name = sorted(n for n in names if n.startswith("xl/worksheets/sheet"))[0]
    sheet = ET.fromstring(z.read(sheet_name))
    cells = {}
    for c in sheet.iter(M + "c"):
        col, row = _col_row(c.get("r"))
        cell_font = fonts[xfs[int(c.get("s", 0))]] if xfs else {"strike": False, "bold": False, "color": ""}
        t = c.get("t")
        runs = None
        if t == "s":
            v = c.find(M + "v")
            if v is None:
                continue
            runs = shared[int(v.text)]
        elif t == "inlineStr":
            is_ = c.find(M + "is")
            rr = is_.findall(M + "r") if is_ is not None else []
            if rr:
                runs = [((r.find(M + "t").text or ""), _attrs_from_props(r.find(M + "rPr"))) for r in rr]
            elif is_ is not None:
                runs = [("".join(x.text or "" for x in is_.iter(M + "t")), None)]
        else:
            v = c.find(M + "v")
            if v is not None:
                runs = [(v.text or "", None)]
        if not runs:
            continue
        chars = []
        for text, attrs in runs:
            a = attrs or cell_font
            chars.extend((ch, a) for ch in text.replace("\r", ""))
        if "".join(ch for ch, _ in chars).strip():
            cells[(col, row)] = chars
    return cells


def _text(chars):
    return "".join(ch for ch, _ in chars)


def _line_attrs(chars):
    vis = [a for ch, a in chars if not ch.isspace()]
    if not vis:
        return {"strike": False, "bold": False, "color": ""}
    colors = {}
    for a in vis:
        colors[a["color"]] = colors.get(a["color"], 0) + 1
    return {
        "strike": sum(a["strike"] for a in vis) > len(vis) / 2,
        "bold": sum(a["bold"] for a in vis) > len(vis) / 2,
        "color": max(colors, key=colors.get),
        "colors": set(colors),
    }


def _split_lines(chars):
    lines, cur = [], []
    for ch, a in chars:
        if ch == "\n":
            lines.append(cur)
            cur = []
        else:
            cur.append((ch, a))
    lines.append(cur)
    return [l for l in lines if _text(l).strip()]


# --------------------------------------------------------------------------- Stundenplan-Struktur

def parse_legend(cells):
    courses, exams = {}, {}
    rows = sorted({r for (_, r) in cells})
    start = None
    for r in rows:
        if _text(cells.get(("A", r), [])).strip().upper() == "VORLESUNGEN":
            start = r
            break
    if start is None:
        return courses, exams
    for r in rows:
        if r <= start + 1:
            continue
        a, c = _text(cells.get(("A", r), [])).strip(), _text(cells.get(("C", r), [])).strip()
        f, g = _text(cells.get(("F", r), [])).strip(), _text(cells.get(("G", r), [])).strip()
        if a and c:
            courses[a] = c
        if f and g:
            exams[f] = g
    return courses, exams


def _hm(h, m):
    return "%02d:%02d" % (int(h), int(m))


def parse_entries(chars):
    """Zelle -> Liste von Einträgen (Titelzeile + angehängte Klammerzeilen)."""
    entries = []
    for line in _split_lines(chars):
        txt = _text(line).strip()
        if txt.startswith("(") and entries:
            entries[-1]["extra"].append(txt)
            entries[-1]["attr_lines"].append(_line_attrs(line))
        else:
            entries.append({"title": txt, "extra": [], "attrs": _line_attrs(line), "attr_lines": []})
    return entries


def parse_schedule(path, group=None):
    cells = read_sheet(path)
    courses, exams = parse_legend(cells)
    stand = ""
    for (col, row), ch in cells.items():
        t = _text(ch)
        if t.startswith("Stand:"):
            stand = t.replace("Stand:", "").strip()

    rows = sorted({r for (_, r) in cells})
    day_cols = ["C", "D", "E", "F", "G", "H"]
    current_dates = None
    events, seen = [], set()

    for r in rows:
        a = _text(cells.get(("A", r), [])).strip()
        if re.match(r"KW\s*\d+", a):
            current_dates = {}
            for col in day_cols:
                m = DATE_HDR_RE.search(_text(cells.get((col, r), [])))
                if m:
                    current_dates[col] = dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            continue
        if a.upper() == "VORLESUNGEN":
            break
        if current_dates is None:
            continue
        label = _text(cells.get(("B", r), [])).strip()
        if not label:
            continue
        slot = SLOT_NAMES.get(label.split("\n")[0].strip().lower())
        if not slot:
            continue
        tm = TIME_RE.search(label)
        default = (_hm(tm.group(1), tm.group(2)), _hm(tm.group(3), tm.group(4))) if tm else None

        for col in day_cols:
            if col not in current_dates or (col, r) not in cells:
                continue
            day = current_dates[col]
            for e in parse_entries(cells[(col, r)]):
                ev = build_event(e, day, slot, default, courses, exams)
                if ev is None:
                    continue
                if group and ev["group"] and ev["group"] != group:
                    continue
                key = (ev["date"], ev["title"], ev["start"], ev["end"])
                if key in seen:
                    continue
                seen.add(key)
                ev["cell"] = "%s%d" % (col, r)
                events.append(ev)

    # Ganztägige Einträge (Feiertage) stehen meist in mehreren Slots -> nur einmal
    out, seen_allday = [], set()
    for ev in events:
        if ev["all_day"]:
            k = (ev["date"], ev["title"])
            if k in seen_allday:
                continue
            seen_allday.add(k)
        out.append(ev)
    return out, stand, courses


def build_event(e, day, slot, default, courses, exams):
    title = e["title"]
    full = " ".join([title] + e["extra"])
    attrs = e["attrs"]
    new = False
    if title.upper().startswith("NEU "):
        new = True
        title = title[4:].strip()

    tm = TIME_RE.search(full)
    gm = GROUP_RE.search(full)
    group = gm.group(1) if gm else ""
    clean = GROUP_RE.sub("", TIME_RE.sub("", title)).strip(" -–")
    clean = re.sub(r"\s+", " ", clean)
    if not clean:
        return None

    if tm:
        start, end = _hm(tm.group(1), tm.group(2)), _hm(tm.group(3), tm.group(4))
    elif default:
        start, end = default
    else:
        start = end = None

    colors = set(attrs.get("colors", set()))
    for la in e["attr_lines"]:
        colors |= la.get("colors", set())
    online = COLOR_PURPLE in colors
    cancelled = attrs["strike"]
    exam_m = re.match(r"^K\d+\s", clean)
    is_exam = bool(exam_m) or (attrs["color"] == COLOR_GREEN and attrs["bold"])

    abbr, lecturer = clean, ""
    if " - " in clean:
        abbr, lecturer = [x.strip() for x in clean.split(" - ", 1)]

    kind = "lecture"
    all_day = False
    if is_exam:
        kind = "exam"
    elif slot == "abgabe":
        kind, all_day = "deadline", True
    elif not lecturer and not tm and courses.get(abbr) == abbr:
        kind, all_day = "holiday", True   # z. B. "Tag der Dt. Einheit"
    elif not lecturer:
        kind = "other"

    if kind == "exam":
        name = exams.get(clean, "Prüfung " + clean)
        display = "🎓 %s (%s)" % (name, clean.split()[0])
    elif kind == "deadline":
        display = "📌 Abgabe: " + clean
    elif lecturer:
        display = "%s – %s" % (abbr, lecturer)
    else:
        display = clean

    uid_src = "|".join([day.isoformat(), slot, abbr, lecturer, group])
    return {
        "uid": hashlib.sha1(uid_src.encode("utf-8")).hexdigest()[:20],
        "date": day.isoformat(),
        "slot": slot,
        "start": None if all_day else start,
        "end": None if all_day else end,
        "all_day": all_day or start is None,
        "title": display,
        "abbr": abbr,
        "lecturer": lecturer,
        "course": courses.get(abbr, ""),
        "group": group,
        "kind": kind,
        "online": online,
        "new": new,
        "cancelled_excel": cancelled,
        "raw": full,
    }


# --------------------------------------------------------------------------- Mails auswerten

CANCEL_RE = re.compile(r"entf[aä]ll|ausfall|f[aä]ll\w*\s+(?:\w+\s+){0,6}aus\b|abgesagt|findet\s+nicht\s+statt", re.I)
WEEKDAYS = {"montag": 0, "dienstag": 1, "mittwoch": 2, "donnerstag": 3, "freitag": 4, "samstag": 5}


def _mail_dates(text, received):
    """Alle Termine, auf die sich eine Mail bezieht."""
    dates = set()
    for m in re.finditer(r"\b(\d{1,2})\.\s?(\d{1,2})\.(\d{2,4})?", text):
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        year = int(y) if y else received.year
        if year < 100:
            year += 2000
        try:
            cand = dt.date(year, mo, d)
        except ValueError:
            continue
        if not y and cand < received - dt.timedelta(days=60):
            cand = cand.replace(year=year + 1)
        dates.add(cand)
    low = text.lower()
    if re.search(r"\bheute\b", low):
        dates.add(received)
    if re.search(r"(?<!guten )\bmorgen\b", low) and "übermorgen" not in low:
        dates.add(received + dt.timedelta(days=1))
    if "übermorgen" in low:
        dates.add(received + dt.timedelta(days=2))
    for name, wd in WEEKDAYS.items():
        if re.search(r"\b(?:am|diesen|kommenden|nächsten)\s+%s" % name, low):
            delta = (wd - received.weekday()) % 7
            dates.add(received + dt.timedelta(days=delta))
    return dates


def _mail_slots(low):
    slots = set()
    if re.search(r"vormittag|vor-\s*(und|&)", low):
        slots.add("vormittag")
    if "nachmittag" in low:
        slots.add("nachmittag")
    if re.search(r"\babend", low):
        slots.add("abend")
    return slots


def _cohort_ok(text, cohort):
    """Prüft, ob die Mail unseren Jahrgang/Kohorte (z. B. A1) betrifft."""
    years = set(re.findall(r"d?WI\s?(20\d\d)", text))
    if years and cohort["year"] not in years:
        return False
    found = set(re.findall(r"\bA(\d)\b", text))
    if found and cohort["cohort"] not in found:
        return False
    return True


def match_mail(mail, events, cohort):
    """-> (liste_uids, grund) ; leere Liste + grund 'unklar' wenn Ausfall-Mail, aber nicht zuordenbar."""
    text = (mail.get("subject", "") + "\n" + mail.get("content", ""))
    if not CANCEL_RE.search(text):
        return None, "kein Ausfall"
    if not _cohort_ok(text, cohort):
        return None, "andere Gruppe"
    received = dt.datetime.fromisoformat(mail["date"]).date()
    dates = _mail_dates(text, received)
    if not dates:
        return [], "kein Datum erkannt"
    low = text.lower()
    hits = []
    for ev in events:
        if dt.date.fromisoformat(ev["date"]) not in dates or ev["kind"] in ("holiday",):
            continue
        names = [n for n in (ev["abbr"], ev["lecturer"].split("-")[0].rstrip(".")) if n and len(n) >= 2]
        if any(re.search(r"\b%s\b" % re.escape(n), text, re.I) for n in names):
            hits.append(ev)
    slots = _mail_slots(low)
    if slots:
        narrowed = [e for e in hits if e["slot"] in slots]
        hits = narrowed or hits
    if not hits:
        return [], "keine passende Veranstaltung gefunden"
    return [e["uid"] for e in hits], "ok"


# --------------------------------------------------------------------------- Räume (SharePoint-Terminliste)
#
# Die Hochschule pflegt eine SharePoint-Liste mit allen Terminen inkl. Raum und Status.
# Sie ist nur mit Uni-Login erreichbar. Externe Apps sperrt die Uni, deshalb nutzt das
# Tool einen eigenen Chrome-Browser mit gespeicherter Anmeldung (einmalig: `login`).

PROFILE_DIR = os.path.expanduser("~/Library/Application Support/UniKalender/browser-profile")


class LoginRequired(Exception):
    pass


def _rooms_cfg(config):
    r = config.get("rooms") or {}
    return {
        "enabled": r.get("enabled", False),
        "site": r.get("site", "https://leibnizfh.sharepoint.com/sites/display"),
        "list": r.get("list", "Terminliste"),
    }


ROOMS_JS = """async ([site, list]) => {
  const h = {Accept: 'application/json;odata=nometadata'};
  let url = site + "/_api/web/lists/GetByTitle('" + list + "')/items?$select=Title,Gruppe,Lehrende_x002f_Orga,Hinweis,Status,Raum,Datum,Zeit&$top=1000";
  let all = [];
  while (url) {
    const r = await fetch(url, {headers: h});
    if (!r.ok) return {error: r.status};
    const j = await r.json();
    all = all.concat(j.value);
    url = j['odata.nextLink'] || j['@odata.nextLink'];
  }
  return {items: all};
}"""


def fetch_room_items(config, headless=True, wait_login=0):
    """Lädt die Terminliste über einen Chrome mit gespeicherter Uni-Anmeldung."""
    from playwright.sync_api import sync_playwright
    rc = _rooms_cfg(config)
    with sync_playwright() as p:
        ctx = p.chromium.launch_persistent_context(PROFILE_DIR, channel="chrome", headless=headless)
        try:
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            page.goto(rc["site"] + "/Lists/" + rc["list"] + "/AllItems.aspx", wait_until="domcontentloaded", timeout=60000)
            deadline = dt.datetime.now() + dt.timedelta(seconds=wait_login)
            while True:
                try:
                    page.wait_for_load_state("networkidle", timeout=15000)
                except Exception:
                    pass
                if "sharepoint.com" in page.url and "login" not in page.url:
                    res = page.evaluate(ROOMS_JS, [rc["site"], rc["list"]])
                    if "items" in res:
                        return res["items"]
                if dt.datetime.now() >= deadline:
                    raise LoginRequired(page.url)
                page.wait_for_timeout(2000)
        finally:
            ctx.close()


def _room_label(raw):
    raw = (raw or "").strip()
    if not raw or raw.lower() in ("tba", "tbd", "-", "n.n."):
        return None
    m = re.match(r"^R\s*(\S+)$", raw)
    return "Raum " + m.group(1) if m else raw


def _minutes(hm):
    h, m = hm.split(":")
    return int(h) * 60 + int(m)


def match_rooms(events, items, wanted_groups):
    """Ordnet jedem Excel-Termin den passenden Listeneintrag zu -> {uid: item}."""
    mine = [i for i in items if set((i.get("Gruppe") or "").split("#")) & wanted_groups]
    by_date = {}
    for it in mine:
        tm = TIME_RE.search(it.get("Zeit") or "")
        if not tm or not it.get("Datum"):
            continue
        it["_s"], it["_e"] = _minutes(_hm(tm.group(1), tm.group(2))), _minutes(_hm(tm.group(3), tm.group(4)))
        by_date.setdefault(it["Datum"][:10], []).append(it)
    result = {}
    for ev in events:
        if ev["all_day"] or not ev["start"]:
            continue
        s, e = _minutes(ev["start"]), _minutes(ev["end"])
        cands = [it for it in by_date.get(ev["date"], []) if it["_s"] < e and s < it["_e"]]
        best, best_score = None, 0
        for it in cands:
            title = (it.get("Title") or "").strip()
            lect = (it.get("Lehrende_x002f_Orga") or "").lower()
            score = 0
            if ev["lecturer"] and lect.startswith(ev["lecturer"].lower().rstrip(".").split("-")[0]):
                score += 2
            if title and title in (ev["course"], ev["abbr"]):
                score += 2
            if ev["kind"] == "exam" and title.lower().startswith("prüfung"):
                score += 1
            if it["_s"] == s:
                score += 1
            if score > best_score:
                best, best_score = it, score
        if best is not None and best_score >= 2:
            result[ev["uid"]] = best
    return result


# --------------------------------------------------------------------------- Plan erzeugen

def fmt_ev(ev):
    d = dt.date.fromisoformat(ev["date"])
    wd = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"][d.weekday()]
    t = "ganztägig" if ev["all_day"] else "%s–%s" % (ev["start"], ev["end"])
    return "%s %s %s %s" % (wd, d.strftime("%d.%m."), t, ev["title"])


def plan(config, mails, state):
    events, stand, courses = parse_schedule(os.path.expanduser(config["xlsxPath"]), config.get("group"))
    if not events:
        raise SystemExit("Keine Termine in der Excel-Datei gefunden – Abbruch (Datei kaputt/unbekanntes Format?)")
    notes = []
    group = config.get("group", "A1/2")
    cohort = {"cohort": re.match(r"A(\d)", group).group(1) if re.match(r"A(\d)", group) else "1",
              "year": str(config.get("jahrgang", "2025"))}

    # --- Mails
    mail_cancel = state.setdefault("mail_cancellations", {})
    processed = set(state.get("processed_mails", []))
    for mail in mails:
        mid = mail.get("id") or (mail.get("subject", "") + mail.get("date", ""))
        uids, reason = match_mail(mail, events, cohort)
        if uids is None:
            continue
        if mid in processed:
            continue
        processed.add(mid)
        if uids:
            for uid in uids:
                mail_cancel[uid] = {"subject": mail.get("subject", ""), "date": mail.get("date", "")}
            for ev in events:
                if ev["uid"] in uids and ev["date"] >= dt.date.today().isoformat():
                    notes.append("Mail: entfällt – " + fmt_ev(ev))
        else:
            notes.append("Mail „%s“ klingt nach Ausfall, konnte aber nicht zugeordnet werden (%s) – bitte selbst prüfen."
                         % (mail.get("subject", "?"), reason))
    state["processed_mails"] = sorted(processed)[-500:]

    # --- Änderungen gegenüber letztem Lauf (für die Benachrichtigung)
    prev = {e["uid"]: e for e in state.get("last_events", [])}
    cur = {e["uid"]: e for e in events}
    if prev:
        today = dt.date.today().isoformat()
        for uid, ev in cur.items():
            if ev["date"] < today:
                continue
            p = prev.get(uid)
            if p is None:
                notes.append("Neu im Plan: " + fmt_ev(ev))
            elif (p["start"], p["end"]) != (ev["start"], ev["end"]):
                notes.append("Zeit geändert: %s (vorher %s–%s)" % (fmt_ev(ev), p["start"], p["end"]))
            elif ev["cancelled_excel"] and not p["cancelled_excel"]:
                notes.append("Plan: entfällt – " + fmt_ev(ev))
            elif ev["online"] != p["online"]:
                notes.append(("Jetzt online: " if ev["online"] else "Jetzt Präsenz: ") + fmt_ev(ev))
        for uid, p in prev.items():
            if uid not in cur and p["date"] >= today:
                notes.append("Aus dem Plan entfernt: " + fmt_ev(p))
    state["last_events"] = events
    state["stand"] = stand

    # --- Räume aus der SharePoint-Terminliste
    rooms = state.setdefault("rooms", {})            # uid -> {"room", "status", "hint"}
    rc = _rooms_cfg(config)
    if rc["enabled"]:
        items = None
        try:
            items = fetch_room_items(config)
            state["room_items_fetched"] = dt.datetime.now().isoformat(timespec="minutes")
        except LoginRequired:
            notes.append("Räume: Uni-Anmeldung abgelaufen – bitte „UniKalender Login.command“ doppelklicken.")
        except Exception as exc:  # Netzwerk, Chrome fehlt, …
            notes.append("Räume konnten nicht geladen werden (%s) – nutze letzten Stand." % type(exc).__name__)
        if items is not None:
            yy = str(config.get("jahrgang", "2025"))[2:]
            wanted = {"dWI%s-A%s" % (yy, cohort["cohort"]), "dWI%s-%s" % (yy, group)}
            today = dt.date.today().isoformat()
            tomorrow = (dt.date.today() + dt.timedelta(days=1)).isoformat()
            matched = match_rooms(events, items, wanted)
            for ev in events:
                it = matched.get(ev["uid"])
                if it is None:
                    continue
                new = {"room": _room_label(it.get("Raum")),
                       "status": (it.get("Status") or "").strip(),
                       "hint": (it.get("Hinweis") or "").strip()}
                old = rooms.get(ev["uid"], {})
                if new["room"] and new["room"] != old.get("room") and today <= ev["date"] <= tomorrow:
                    notes.append("Raum: %s → %s" % (fmt_ev(ev), new["room"]))
                if new["status"].lower() == "storniert" and old.get("status", "").lower() != "storniert" \
                        and ev["date"] >= today and not ev["cancelled_excel"]:
                    notes.append("Terminliste: storniert – " + fmt_ev(ev))
                if not new["room"] and old.get("room"):
                    new["room"] = old["room"]           # bekannten Raum nicht durch "tba" überschreiben
                rooms[ev["uid"]] = new

    # --- Kalendertermine bauen
    out = []
    for ev in events:
        rinfo = rooms.get(ev["uid"], {})
        list_cancel = rinfo.get("status", "").lower() == "storniert"
        cancelled = ev["cancelled_excel"] or ev["uid"] in mail_cancel or list_cancel
        title = ev["title"]
        if cancelled:
            title = "❌ ENTFÄLLT: " + title
        info = []
        if ev["course"] and ev["course"] != ev["abbr"]:
            info.append(ev["course"])
        if ev["lecturer"]:
            info.append("Dozent/in: " + ev["lecturer"])
        if ev["group"]:
            info.append("Gruppe: " + ev["group"])
        if ev["online"]:
            info.append("Online-Veranstaltung")
        if ev["new"]:
            info.append("Im Plan als NEU markiert")
        if ev["cancelled_excel"]:
            info.append("Im Stundenplan durchgestrichen")
        if ev["uid"] in mail_cancel:
            info.append("Ausfall per Mail: " + mail_cancel[ev["uid"]]["subject"])
        if list_cancel:
            info.append("In der Terminliste storniert")
        if rinfo.get("hint"):
            info.append("Hinweis: " + rinfo["hint"])
        room = rinfo.get("room")
        if rc["enabled"] and not ev["all_day"] and not room and not ev["online"] \
                and ev["date"] >= dt.date.today().isoformat():
            info.append("Raum: noch nicht festgelegt (tba)")
        info.append("Stundenplan-Stand: " + (stand or "?"))
        location = room or ("Online" if ev["online"] else "")
        out.append({
            "uid": ev["uid"],
            "title": title,
            "date": ev["date"],
            "start": ev["start"],
            "end": ev["end"],
            "allDay": ev["all_day"],
            "location": location,
            "notes": "\n".join(info),
            "cancelled": cancelled,
        })
    notes = list(dict.fromkeys(notes))  # gleiche Meldung nur einmal (z. B. Mail doppelt im Posteingang)
    return {"events": out, "notes": notes, "stand": stand}


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    p1 = sub.add_parser("parse")
    p1.add_argument("xlsx")
    p1.add_argument("--group", default=None)
    p2 = sub.add_parser("plan")
    p2.add_argument("--config", required=True)
    p2.add_argument("--mails", required=True)
    p2.add_argument("--state", required=True)
    p2.add_argument("--dry-run", action="store_true", help="state nicht speichern")
    p3 = sub.add_parser("login", help="Chrome-Fenster öffnen und einmalig mit der Uni-Mail anmelden (für die Räume)")
    p3.add_argument("--config", required=True)
    p4 = sub.add_parser("rooms", help="Räume abrufen und Zuordnung anzeigen (zum Testen)")
    p4.add_argument("--config", required=True)
    a = ap.parse_args()

    if a.cmd in ("login", "rooms"):
        with open(a.config, encoding="utf-8") as f:
            config = json.load(f)
        if a.cmd == "login":
            print("Ein Chrome-Fenster öffnet sich. Bitte mit der Uni-Mail anmelden (\"Angemeldet bleiben\": Ja).")
            items = fetch_room_items(config, headless=False, wait_login=600)
            print("✓ Anmeldung gespeichert – %d Einträge in der Terminliste gefunden. Das Fenster schließt sich jetzt." % len(items))
            return
        group = config.get("group", "A1/2")
        yy = str(config.get("jahrgang", "2025"))[2:]
        wanted = {"dWI%s-A%s" % (yy, group[1]), "dWI%s-%s" % (yy, group)}
        events, _, _ = parse_schedule(os.path.expanduser(config["xlsxPath"]), group)
        matched = match_rooms(events, fetch_room_items(config), wanted)
        for ev in events:
            if ev["date"] >= dt.date.today().isoformat():
                it = matched.get(ev["uid"])
                print("%-60s %s" % (fmt_ev(ev), "%s [%s]" % (it.get("Raum"), it.get("Status")) if it else "—"))
        return

    if a.cmd == "parse":
        events, stand, _ = parse_schedule(a.xlsx, a.group)
        print("Stand:", stand, "|", len(events), "Termine")
        for ev in events:
            flags = "".join([
                " [ONLINE]" if ev["online"] else "",
                " [NEU]" if ev["new"] else "",
                " [GESTRICHEN]" if ev["cancelled_excel"] else "",
                " (%s)" % ev["group"] if ev["group"] else "",
            ])
            print("%-5s %s%s" % (ev["cell"], fmt_ev(ev), flags))
        return

    with open(a.config, encoding="utf-8") as f:
        config = json.load(f)
    with open(a.mails, encoding="utf-8") as f:
        mails = json.load(f)
    try:
        with open(a.state, encoding="utf-8") as f:
            state = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        state = {}
    result = plan(config, mails, state)
    if not a.dry_run:
        with open(a.state, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=1)
    json.dump(result, sys.stdout, ensure_ascii=False)


if __name__ == "__main__":
    main()
