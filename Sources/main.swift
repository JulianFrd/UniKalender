// UniKalender – Systemzugriff (Kalender via EventKit, Mails via Apple Mail).
// Die eigentliche Logik (Excel lesen, Mails auswerten) steckt in unikal.py.
//
// Aufruf:  unikal run [--dry-run] [--no-mail] [--force]
//          unikal mails            (zeigt die gelesenen Uni-Mails, zum Testen)
//          unikal calendars        (listet Kalender/Accounts, zum Testen)

import EventKit
import Foundation
import Network

let supportDir = FileManager.default.homeDirectoryForCurrentUser
    .appendingPathComponent("Library/Application Support/UniKalender")
let configURL = supportDir.appendingPathComponent("config.json")
let stateURL = supportDir.appendingPathComponent("state.json")
let mailsURL = supportDir.appendingPathComponent("mails.json")
let scriptURL = supportDir.appendingPathComponent("unikal.py")
let logURL = supportDir.appendingPathComponent("unikal.log")
let marker = "unikal://event/"

struct Config: Decodable {
    var xlsxPath: String
    var group: String?
    var calendarName: String?
    var python: String?
    var mailDays: Int?
    var mailSenderFilter: String?
    var notify: Bool?
    var settleSeconds: Int?
}

struct PlannedEvent: Decodable {
    var uid: String
    var title: String
    var date: String
    var start: String?
    var end: String?
    var allDay: Bool
    var location: String
    var notes: String
    var cancelled: Bool
}

struct Plan: Decodable {
    var events: [PlannedEvent]
    var notes: [String]
    var stand: String
}

// MARK: - Logging

func log(_ msg: String) {
    let df = DateFormatter()
    df.dateFormat = "yyyy-MM-dd HH:mm:ss"
    let line = "[\(df.string(from: Date()))] \(msg)\n"
    FileHandle.standardError.write(line.data(using: .utf8)!)
    if let h = try? FileHandle(forWritingTo: logURL) {
        h.seekToEndOfFile()
        h.write(line.data(using: .utf8)!)
        try? h.close()
    } else {
        try? line.data(using: .utf8)!.write(to: logURL)
    }
}

func fail(_ msg: String) -> Never {
    log("FEHLER: \(msg)")
    notify(title: "UniKalender – Fehler", body: msg)
    exit(1)
}

// MARK: - AppleScript (Mail + Benachrichtigung)

func runAppleScript(_ source: String) -> NSAppleEventDescriptor? {
    var err: NSDictionary?
    let script = NSAppleScript(source: source)
    let res = script?.executeAndReturnError(&err)
    if let err = err {
        log("AppleScript-Fehler: \(err)")
        return nil
    }
    return res
}

func asQuote(_ s: String) -> String {
    "\"" + s.replacingOccurrences(of: "\\", with: "\\\\").replacingOccurrences(of: "\"", with: "\\\"") + "\""
}

func notify(title: String, body: String) {
    _ = runAppleScript("display notification \(asQuote(body)) with title \(asQuote(title))")
}

func fetchMails(days: Int, senderFilter: String) -> [[String: String]] {
    // Holt Mails der letzten `days` Tage aus allen Posteingängen von Apple Mail.
    let src = """
    set RS to ASCII character 30
    set FS to ASCII character 31
    set cutoff to (current date) - (\(days) * days)
    set out to ""
    tell application "Mail"
        try
            check for new mail
        end try
    end tell
    delay 20
    tell application "Mail"
        set msgs to (messages of inbox whose date received > cutoff and sender contains \(asQuote(senderFilter)))
        repeat with m in msgs
            try
                set d to date received of m
                set iso to (d as «class isot» as string)
                set out to out & (message id of m) & FS & (subject of m) & FS & (sender of m) & FS & iso & FS & (content of m) & RS
            end try
        end repeat
    end tell
    return out
    """
    guard let res = runAppleScript(src), let str = res.stringValue else {
        log("Mails konnten nicht gelesen werden (Berechtigung für Mail erteilt?)")
        return []
    }
    var mails: [[String: String]] = []
    for rec in str.components(separatedBy: "\u{1E}") where !rec.isEmpty {
        let f = rec.components(separatedBy: "\u{1F}")
        if f.count < 5 { continue }
        mails.append(["id": f[0], "subject": f[1], "sender": f[2], "date": f[3],
                      "content": f[4...].joined(separator: " ")])
    }
    return mails
}

// MARK: - Python-Teil

func runPlan(config: Config, dryRun: Bool) -> Plan {
    let py = ((config.python ?? "/opt/homebrew/bin/python3") as NSString).expandingTildeInPath
    let p = Process()
    p.executableURL = URL(fileURLWithPath: py)
    var args = ["-I", scriptURL.path, "plan", "--config", configURL.path,
                "--mails", mailsURL.path, "--state", stateURL.path]
    if dryRun { args.append("--dry-run") }
    p.arguments = args
    let out = Pipe(), err = Pipe()
    p.standardOutput = out
    p.standardError = err
    do { try p.run() } catch { fail("Python konnte nicht gestartet werden: \(error)") }
    let data = out.fileHandleForReading.readDataToEndOfFile()
    let errData = err.fileHandleForReading.readDataToEndOfFile()
    p.waitUntilExit()
    if p.terminationStatus != 0 {
        fail("Auswertung fehlgeschlagen: \(String(data: errData, encoding: .utf8) ?? "")")
    }
    do { return try JSONDecoder().decode(Plan.self, from: data) } catch {
        fail("Ungültige Ausgabe der Auswertung: \(error)")
    }
}

// MARK: - Kalender

let store = EKEventStore()

func requestCalendarAccess() {
    let sem = DispatchSemaphore(value: 0)
    var granted = false
    store.requestFullAccessToEvents { ok, err in
        granted = ok
        if let err = err { log("Kalenderzugriff: \(err)") }
        sem.signal()
    }
    sem.wait()
    if !granted {
        fail("Kein Kalenderzugriff. Systemeinstellungen → Datenschutz & Sicherheit → Kalender → „unikal“ erlauben.")
    }
}

func findOrCreateCalendar(named name: String, create: Bool) -> EKCalendar? {
    let cals = store.calendars(for: .event)
    if let c = cals.first(where: { $0.title == name && $0.source.sourceType == .calDAV && $0.source.title == "iCloud" })
        ?? cals.first(where: { $0.title == name }) {
        return c
    }
    guard create else { return nil }
    let source = store.sources.first(where: { $0.sourceType == .calDAV && $0.title == "iCloud" })
        ?? store.defaultCalendarForNewEvents?.source
    guard let src = source else { fail("Keine iCloud-Kalenderquelle gefunden.") }
    let cal = EKCalendar(for: .event, eventStore: store)
    cal.title = name
    cal.source = src
    cal.cgColor = CGColor(red: 0.16, green: 0.45, blue: 0.85, alpha: 1)
    do { try store.saveCalendar(cal, commit: true) } catch { fail("Kalender „\(name)“ konnte nicht angelegt werden: \(error)") }
    log("Kalender „\(name)“ in \(src.title) angelegt.")
    return cal
}

func uid(of ev: EKEvent) -> String? {
    if let u = ev.url?.absoluteString, u.hasPrefix(marker) { return String(u.dropFirst(marker.count)) }
    if let n = ev.notes, let r = n.range(of: "[unikal:") {
        let rest = n[r.upperBound...]
        if let end = rest.firstIndex(of: "]") { return String(rest[..<end]) }
    }
    return nil
}

func makeDate(_ day: String, _ time: String?) -> Date {
    var c = DateComponents()
    let d = day.split(separator: "-").map { Int($0)! }
    c.year = d[0]; c.month = d[1]; c.day = d[2]
    if let t = time {
        let hm = t.split(separator: ":").map { Int($0)! }
        c.hour = hm[0]; c.minute = hm[1]
    }
    return Calendar.current.date(from: c)!
}

func apply(_ p: PlannedEvent, to ev: EKEvent) -> Bool {
    let start = makeDate(p.date, p.allDay ? nil : p.start)
    let end = p.allDay ? Calendar.current.date(byAdding: .day, value: 1, to: start)! : makeDate(p.date, p.end)
    let notes = p.notes + "\n[unikal:\(p.uid)]"
    let url = URL(string: marker + p.uid)
    let avail: EKEventAvailability = p.cancelled || p.allDay ? .free : .busy
    var changed = false
    if ev.title != p.title { ev.title = p.title; changed = true }
    if ev.isAllDay != p.allDay { ev.isAllDay = p.allDay; changed = true }
    if p.allDay {
        // EventKit speichert ganztägige Enden intern anders – nur den Tag vergleichen
        if ev.startDate == nil || !Calendar.current.isDate(ev.startDate, inSameDayAs: start) {
            ev.startDate = start; ev.endDate = end; changed = true
        }
    } else {
        if ev.startDate != start { ev.startDate = start; changed = true }
        if ev.endDate != end { ev.endDate = end; changed = true }
    }
    if (ev.location ?? "") != p.location { ev.location = p.location.isEmpty ? nil : p.location; changed = true }
    if (ev.notes ?? "") != notes { ev.notes = notes; changed = true }
    if ev.url != url { ev.url = url; changed = true }
    if ev.availability != avail && ev.calendar?.source.sourceType != .birthdays { ev.availability = avail }
    return changed
}

func sync(plan: Plan, calendarName: String, dryRun: Bool, force: Bool) -> (Int, Int, Int) {
    guard let cal = findOrCreateCalendar(named: calendarName, create: !dryRun) else {
        log("Kalender „\(calendarName)“ existiert noch nicht – Dry-Run legt ihn nicht an.")
        return (plan.events.count, 0, 0)
    }
    log("Kalender: „\(cal.title)“ im Konto „\(cal.source.title)“ (Typ \(cal.source.sourceType == .calDAV ? "iCloud/CalDAV" : "lokal/anderes"))")
    let dates = plan.events.map { makeDate($0.date, nil) }
    let from = min(dates.min()!, Date()).addingTimeInterval(-86400 * 30)
    let to = max(dates.max()!, Date()).addingTimeInterval(86400 * 60)
    let existing = store.events(matching: store.predicateForEvents(withStart: from, end: to, calendars: [cal]))
    var mine: [String: EKEvent] = [:]
    for ev in existing { if let u = uid(of: ev) { mine[u] = ev } }

    // Sicherheitsnetz: wenn plötzlich viel weniger Termine kommen, lieber nichts löschen.
    let now = Date()
    let futureMine = mine.values.filter { $0.endDate > now }.count
    let futurePlanned = plan.events.filter { makeDate($0.date, nil) > now.addingTimeInterval(-86400) }.count
    if !force && futureMine > 10 && futurePlanned < futureMine / 2 {
        fail("Plan enthält nur \(futurePlanned) statt \(futureMine) künftige Termine – vermutlich ist die Datei unvollständig. Abbruch ohne Änderungen (mit --force erzwingen).")
    }

    var created = 0, updated = 0, deleted = 0
    var wanted = Set<String>()
    for p in plan.events {
        wanted.insert(p.uid)
        if let ev = mine[p.uid] {
            if apply(p, to: ev) {
                updated += 1
                log("Aktualisiert: \(p.date) \(p.title)")
                if !dryRun { try? store.save(ev, span: .thisEvent, commit: false) }
            }
        } else {
            let ev = EKEvent(eventStore: store)
            ev.calendar = cal
            _ = apply(p, to: ev)
            created += 1
            if !dryRun {
                do { try store.save(ev, span: .thisEvent, commit: false) } catch { log("Speichern fehlgeschlagen \(p.title): \(error)") }
            }
        }
    }
    for (u, ev) in mine where !wanted.contains(u) {
        if ev.endDate < now { continue }   // Vergangenes nie löschen
        deleted += 1
        log("Gelöscht: \(ev.startDate!) \(ev.title ?? "")")
        if !dryRun { try? store.remove(ev, span: .thisEvent, commit: false) }
    }
    if !dryRun {
        do { try store.commit() } catch { fail("Kalender konnte nicht gespeichert werden: \(error)") }
    }
    return (created, updated, deleted)
}

// MARK: - Netzwerk

func waitForNetwork(timeout: TimeInterval) {
    let monitor = NWPathMonitor()
    let sem = DispatchSemaphore(value: 0)
    var online = false
    monitor.pathUpdateHandler = { path in
        if path.status == .satisfied && !online { online = true; sem.signal() }
    }
    monitor.start(queue: DispatchQueue(label: "net"))
    if sem.wait(timeout: .now() + timeout) == .timedOut {
        log("Kein Internet nach \(Int(timeout)) s – mache trotzdem mit der lokalen Datei weiter.")
    } else {
        log("Internet verfügbar.")
    }
    monitor.cancel()
}

// MARK: - main

try? FileManager.default.createDirectory(at: supportDir, withIntermediateDirectories: true)
let args = Array(CommandLine.arguments.dropFirst())
let cmd = args.first ?? "run"
let dryRun = args.contains("--dry-run")

guard let cdata = try? Data(contentsOf: configURL), let config = try? JSONDecoder().decode(Config.self, from: cdata) else {
    fail("Konfiguration fehlt oder ist ungültig: \(configURL.path)")
}

switch cmd {
case "calendars":
    requestCalendarAccess()
    for c in store.calendars(for: .event) {
        print("\(c.title)  [\(c.source.title), typ \(c.source.sourceType.rawValue)]  schreibbar: \(c.allowsContentModifications)")
    }
case "mails":
    let m = fetchMails(days: config.mailDays ?? 3, senderFilter: config.mailSenderFilter ?? "leibniz-fh.de")
    for x in m { print("\(x["date"]!)  \(x["sender"]!)  |  \(x["subject"]!)") }
    print("\(m.count) Mails")
case "run":
    log("Start\(dryRun ? " (Dry-Run)" : "")")
    if !args.contains("--now") {
        // Nach dem Aufklappen: erst auf WLAN warten, dann OneDrive und Mail kurz synchronisieren lassen.
        waitForNetwork(timeout: 300)
        let settle = config.settleSeconds ?? 90
        log("Warte \(settle) s, damit OneDrive die neueste Datei holt …")
        Thread.sleep(forTimeInterval: TimeInterval(settle))
    }
    var mails: [[String: String]] = []
    if !args.contains("--no-mail") {
        mails = fetchMails(days: config.mailDays ?? 3, senderFilter: config.mailSenderFilter ?? "leibniz-fh.de")
        log("\(mails.count) Uni-Mails der letzten \(config.mailDays ?? 3) Tage gelesen")
    }
    let mdata = try! JSONSerialization.data(withJSONObject: mails)
    try! mdata.write(to: mailsURL)
    let plan = runPlan(config: config, dryRun: dryRun)
    requestCalendarAccess()
    let (c, u, d) = sync(plan: plan, calendarName: config.calendarName ?? "Uni", dryRun: dryRun, force: args.contains("--force"))
    log("Fertig: \(c) neu, \(u) geändert, \(d) gelöscht (Stand Plan: \(plan.stand))")
    for n in plan.notes { log("  • \(n)") }
    if (config.notify ?? true) && !dryRun && !plan.notes.isEmpty {
        let body = plan.notes.prefix(4).joined(separator: "\n") + (plan.notes.count > 4 ? "\n… +\(plan.notes.count - 4) weitere" : "")
        notify(title: "Stundenplan: \(plan.notes.count) Änderung(en)", body: body)
    }
default:
    print("Unbekannter Befehl. Nutzung: unikal run [--dry-run] [--no-mail] [--force] | mails | calendars")
}
