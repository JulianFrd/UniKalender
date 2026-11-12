#!/bin/zsh
# Doppelklicken, um den Raum-Abruf (SharePoint-Terminliste) mit der Uni-Mail anzumelden.
# Nötig beim ersten Mal und wenn das Tool meldet, dass die Anmeldung abgelaufen ist.
D="$HOME/Library/Application Support/UniKalender"
"$D/venv/bin/python" "$D/unikal.py" login --config "$D/config.json"
echo
echo "Du kannst dieses Fenster jetzt schließen."
