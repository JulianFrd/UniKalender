#!/bin/zsh
# Baut UniKalender, installiert es nach ~/Library/Application Support/UniKalender
# und richtet den täglichen Lauf (Uhrzeit: HOUR unten) per launchd ein.
set -euo pipefail

SRC="${0:A:h}"
DEST="$HOME/Library/Application Support/UniKalender"
LABEL="local.unikalender"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
HOUR=8   # Uhrzeit des täglichen Abgleichs

mkdir -p "$DEST" "$HOME/Library/LaunchAgents"

# Nur neu bauen, wenn sich der Swift-Code geändert hat – sonst fragt macOS die
# Berechtigungen (Kalender, Mail, OneDrive) erneut ab.
if [[ ! -f "$DEST/unikal" || "$SRC/Sources/main.swift" -nt "$DEST/unikal" || "$SRC/Info.plist" -nt "$DEST/unikal" ]]; then
  echo "→ Baue Swift-Helfer …"
  swiftc -O -o "$DEST/unikal" "$SRC/Sources/main.swift" \
    -Xlinker -sectcreate -Xlinker __TEXT -Xlinker __info_plist -Xlinker "$SRC/Info.plist"
  codesign --force --sign - --identifier "$LABEL" "$DEST/unikal"
else
  echo "→ Swift-Helfer unverändert, wird nicht neu gebaut."
fi

cp "$SRC/unikal.py" "$DEST/unikal.py"

if [[ ! -f "$DEST/config.json" ]]; then
  cp "$SRC/config.example.json" "$DEST/config.json"
  echo "→ Konfiguration angelegt: $DEST/config.json"
fi

cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <!-- caffeinate hält den Mac wach, solange der Abgleich läuft -->
  <array><string>/usr/bin/caffeinate</string><string>-s</string><string>-i</string><string>$DEST/unikal</string><string>run</string></array>
  <key>StartCalendarInterval</key>
  <dict><key>Hour</key><integer>$HOUR</integer><key>Minute</key><integer>0</integer></dict>
  <!-- zusätzlich bei jeder Anmeldung (falls der Mac ganz aus war statt nur zugeklappt) -->
  <key>RunAtLoad</key><true/>
  <key>StandardOutPath</key><string>$DEST/launchd.log</string>
  <key>StandardErrorPath</key><string>$DEST/launchd.log</string>
</dict>
</plist>
EOF

launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$PLIST"
echo "✓ Installiert. Läuft täglich um $HOUR:00 (oder direkt nach dem Aufwachen, falls der Mac schlief)."
echo "  Sofort ausführen:  launchctl kickstart gui/$(id -u)/$LABEL"
echo "  Log ansehen:       tail -n 50 \"$DEST/unikal.log\""
