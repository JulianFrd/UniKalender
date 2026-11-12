# UniKalender – Anleitung

## Teil 1: Stundenplan-Datei per OneDrive auf den Mac holen

Ziel: Die Excel-Datei der Uni liegt als **immer aktuelle Kopie** auf deinem Mac. So kann das Tool sie jeden Morgen lesen, ohne sich selbst einzuloggen.

### Schritt 1 – Ordner im Browser öffnen
1. Öffne den Stundenplan-Link aus Teams/Moodle im Browser (Safari oder Chrome) und melde dich mit deiner Uni-Mail an.
2. Die Datei öffnet sich in **Excel im Browser**. Klicke oben links auf das **App-Raster-Symbol (⋮⋮⋮)** und wähle **OneDrive**.
3. Klicke in OneDrive links auf **„Geteilt“** (bzw. „Freigegeben“) und dann oben auf **„Mit Ihnen“**.
4. Suche den Ordner **„dWI 2025-28“** (oder „Wirtschaftsinformatik“ bzw. „Stundenpläne“). Wenn du nur die Datei siehst und keinen Ordner: weiter mit Schritt 2, Variante B.

### Schritt 2 – Verknüpfung anlegen
**Variante A (Ordner ist sichtbar, empfohlen):**
1. Fahre mit der Maus über den Ordner **„dWI 2025-28“** und klicke auf die **drei Punkte (…)**.
2. Wähle **„Verknüpfung zu ‚Meine Dateien‘ hinzufügen“** (engl. „Add shortcut to My files“).
3. Prüfen: Links auf **„Meine Dateien“** klicken. Der Ordner erscheint dort mit einem kleinen Verknüpfungs-Pfeil.

**Variante B (nur die Datei ist sichtbar):**
1. Drei Punkte (…) neben der Datei → **„Verknüpfung zu ‚Meine Dateien‘ hinzufügen“**.
2. Fehlt diese Option bei der Datei, öffne die Datei. Klicke oben in der Pfadzeile (z. B. „… › dWI 2025-28“) auf den Ordnernamen und mach dort Variante A.
3. Geht beides nicht, sag mir Bescheid und schick mir einen Screenshot. Dann bauen wir einen anderen Weg.

### Schritt 3 – OneDrive-App auf dem Mac einrichten
1. Öffne die App **OneDrive** (ist schon installiert, liegt in „Programme“).
2. Melde dich mit deiner **Uni-Mail** an (nicht mit einem privaten Microsoft-Konto).
3. Den vorgeschlagenen Speicherort übernehmen und auf **Weiter** klicken, bis die Einrichtung fertig ist.
4. Im Finder erscheint links in der Seitenleiste **„OneDrive – Leibniz-FH …“**. Darin liegt deine Verknüpfung, z. B. `dWI 2025-28/Stundenpläne 3. Semester dWI2025 A1.xlsx`.

### Schritt 4 – Datei dauerhaft lokal halten
1. Im Finder **Rechtsklick** auf die Stundenplan-Datei bzw. den Ordner.
2. **„Immer auf diesem Gerät behalten“** wählen. Das Wolken-Symbol wird dann zu einem ausgefüllten grünen Haken.

So ist die Datei um 8 Uhr garantiert aktuell und lokal verfügbar.

### Schritt 5 – Bescheid geben
Sag mir, dass du fertig bist. Ich suche dann die Datei unter `~/Library/CloudStorage/` und stelle das Tool auf diese Live-Datei um. Danach brauchst du die Kopie auf dem Schreibtisch nicht mehr.

---

## Teil 2: Wann läuft der Abgleich?

- **Mac um 8:00 offen:** Der Abgleich läuft um 8:00.
- **Mac um 8:00 zugeklappt:** Der Abgleich läuft direkt nach dem Aufklappen.
- **Mac war ganz ausgeschaltet:** Der Abgleich läuft direkt nach der Anmeldung.

So läuft ein Abgleich ab:
1. Das Tool wartet auf WLAN.
2. Es wartet dann 90 Sekunden, damit OneDrive die neueste Datei holen kann.
3. Es liest die Mails und gleicht den Kalender ab.

Ca. 2–3 Minuten nach dem Aufklappen ist der Kalender aktuell, kurz danach auch auf dem iPhone.

Die Wartezeit lässt sich in `config.json` über `"settleSeconds"` ändern.

---

## Teil 3: Bedienung

| Was | Befehl |
|---|---|
| Sofort aktualisieren | `launchctl kickstart gui/$(id -u)/local.unikalender` |
| Log ansehen | `tail -n 40 ~/Library/Application\ Support/UniKalender/unikal.log` |
| Einstellungen | `~/Library/Application Support/UniKalender/config.json` |
| Neu installieren nach Änderungen | `./install.sh` im Ordner `UniKalender` |
| Komplett entfernen | `launchctl bootout gui/$(id -u)/local.unikalender`, dann den Kalender „Uni“ in der Kalender-App löschen |

**So sieht es im Kalender aus:**
- Vorlesung: `WuD – Christmann`. In der Notiz stehen der volle Kursname und die Dozentin bzw. der Dozent.
- Online-Veranstaltung: Ort = „Online“
- Prüfung: `🎓 Prüfung Statistik 1 (K60)`
- Ausfall (durchgestrichen im Plan **oder** per Mail): `❌ ENTFÄLLT: …`. Der Termin bleibt sichtbar und wird als „frei“ markiert.
- Deine Termine, die du selbst in den Kalender „Uni“ einträgst, werden nie angefasst.
- Vergangene Termine werden nie gelöscht.
