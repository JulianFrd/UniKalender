# UniKalender 📅

Bringt den Excel-Stundenplan der **Leibniz-FH** (Wirtschaftsinformatik, dWI) automatisch in den **iPhone-/iCloud-Kalender** und hält ihn aktuell.

Der Stundenplan wird von der Hochschule als Excel-Datei in OneDrive gepflegt und ändert sich laufend. Ausfälle kommen zusätzlich oft morgens per Mail. UniKalender gleicht beides jeden Morgen automatisch mit einem eigenen Kalender ab. So zeigt das iPhone immer den aktuellen Stand.

## Was das Tool kann

- **Liest den Excel-Stundenplan** direkt aus der von OneDrive synchronisierten Datei. Es braucht keine zusätzliche Python-Bibliothek, nur die Standardbibliothek.
- **Versteht die Formatierung des Plans:**
  - Standard-Zeitslots (Vormittag, Nachmittag, Abend) und abweichende Zeiten wie `(18:00-21:15)`
  - Lila Schrift → Ort „Online“
  - Durchgestrichen und rot → `❌ ENTFÄLLT: …`
  - Grün und fett (`K60 …`) → `🎓 Prüfung …`
  - `NEU`-Markierungen und Feiertage (als ganztägige Termine)
- **Filtert nach Übungsgruppe** (z. B. nur `A1/2`).
- **Wertet Uni-Mails aus Apple Mail aus.** Eine Mail wie „leider entfallen heute alle Veranstaltungen von Herrn X am Vormittag“ wird dem passenden Termin zugeordnet. Das Tool erkennt dabei das Datum („heute“, „morgen“, „11.11.“, „am Mittwoch“), den Kurs bzw. die Dozentin oder den Dozenten, den Zeitslot und die Kohorte.
- **Synchronisiert, statt neu anzulegen:** Termine werden aktualisiert oder entfernt, es entstehen keine Duplikate. Vergangene Termine und selbst eingetragene Termine bleiben unangetastet.
- **Sicherheitsnetz:** Liefert die Datei plötzlich viel weniger Termine, bricht das Tool ab, statt den Kalender zu leeren.
- **Läuft automatisch** täglich um 8:00. Ist der Mac zugeklappt, läuft es direkt nach dem Aufklappen.
- **Schickt eine Mac-Benachrichtigung**, wenn sich etwas geändert hat.

## Aufbau

```
OneDrive (Excel)  ─┐
                   ├─►  unikal.py  (Python: Excel parsen, Mails auswerten, Soll-Liste erzeugen)
Apple Mail ────────┘          │
                              ▼
                     unikal  (Swift: EventKit-Sync)  ─►  iCloud-Kalender „Uni“  ─►  iPhone
                              ▲
                  launchd (täglich 8:00 / nach dem Aufwachen)
```

| Datei | Aufgabe |
|---|---|
| `unikal.py` | Logik: Excel lesen (ZIP und XML inkl. Rich-Text-Formatierung), Mails auswerten, Änderungen erkennen |
| `Sources/main.swift` | Systemzugriff: Kalender über EventKit, Mails über AppleScript, Benachrichtigungen |
| `Info.plist` | Begründungstexte für die macOS-Berechtigungen (Kalender, Mail) |
| `install.sh` | Baut den Swift-Helfer, installiert alles und richtet den launchd-Job ein |

**Warum Python und Swift?** macOS vergibt Kalender- und Mail-Berechtigungen pro Programm. Ein kleines, signiertes Swift-Programm mit eigener `Info.plist` bekommt diese Berechtigungen zuverlässig, auch wenn es im Hintergrund über launchd läuft. Python übernimmt die Auswertung, weil sich das dort einfacher testen lässt.

## Installation

Voraussetzungen: macOS 14+, Xcode Command Line Tools (`xcode-select --install`), Python 3 (z. B. über Homebrew), die OneDrive-App mit Uni-Konto und Apple Mail mit Uni-Konto (optional).

```bash
git clone https://github.com/<dein-name>/UniKalender.git
cd UniKalender
./install.sh
```

Danach in `~/Library/Application Support/UniKalender/config.json` eintragen:

| Feld | Bedeutung |
|---|---|
| `xlsxPath` | Pfad zur synchronisierten Stundenplan-Datei (siehe [ANLEITUNG.md](ANLEITUNG.md)) |
| `group` | deine Übungsgruppe, z. B. `A1/2` |
| `jahrgang` | z. B. `2025`. Mails für andere Jahrgänge werden ignoriert. |
| `calendarName` | Name des Kalenders, der angelegt wird (Standard: `Uni`) |
| `mailSenderFilter` | Nur Mails von diesem Absender bzw. dieser Domain werden gelesen. |
| `settleSeconds` | Wartezeit nach dem Aufwachen, damit OneDrive die neueste Datei holt (Standard: 90) |

Ersten Lauf starten. macOS fragt dabei nach den Berechtigungen für Kalender, Mail und OneDrive:

```bash
launchctl kickstart gui/$(id -u)/local.unikalender
```

Die ausführliche Schritt-für-Schritt-Anleitung (OneDrive-Verknüpfung, Bedienung, Deinstallation) steht in **[ANLEITUNG.md](ANLEITUNG.md)**.

## Kalender mit anderen teilen

Den iCloud-Kalender „Uni“ kannst du auf dem iPhone unter Kalender → ⓘ → **Öffentlicher Kalender** freigeben. Den Link können Kommilitonen auf iPhone, Android oder in Google Kalender abonnieren. Sie bekommen alle Updates automatisch, ohne selbst etwas zu installieren.

## Grenzen

- Das Tool läuft auf dem **Mac**. Ist der Mac zugeklappt, läuft der Abgleich erst beim Aufklappen. Eine reine Cloud-Lösung scheitert daran, dass die Hochschule externen Apps keinen Zugriff auf OneDrive erlaubt (Administratorgenehmigung erforderlich).
- Das Tool ist auf das Format der Leibniz-FH-Stundenpläne zugeschnitten. Andere Hochschulen bräuchten einen angepassten Parser.
- Die Mail-Auswertung arbeitet mit Regeln, nicht mit KI. Ungewöhnlich formulierte Mails werden nicht zugeordnet. Das Tool meldet sie dann aber als „bitte selbst prüfen“.

## Lizenz

MIT – siehe [LICENSE](LICENSE).
