#!/bin/bash
# Installeert de wekelijkse Fleetmate-export (donderdag 07:00) als launchd-agent.
set -euo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
LABEL="nl.rwmconsultancy.fleetmate-export"
PLIST_DST="$HOME/Library/LaunchAgents/$LABEL.plist"

cd "$DIR"
mkdir -p logs exports backups

echo "→ Python venv + openpyxl"
python3 -m venv .venv
.venv/bin/pip install --quiet --upgrade pip openpyxl

chmod 600 config.env
chmod +x fleetmate_export.py

echo "→ launchd-agent installeren"
mkdir -p "$HOME/Library/LaunchAgents"
sed "s|__DIR__|$DIR|g" "$LABEL.plist" > "$PLIST_DST"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
launchctl bootstrap "gui/$(id -u)" "$PLIST_DST"
launchctl enable "gui/$(id -u)/$LABEL"

echo
echo "Klaar. De export draait elke donderdag om 07:00."
echo "Nu direct testen:   launchctl kickstart -k gui/$(id -u)/$LABEL"
echo "Log bekijken:       tail -f \"$DIR/logs/export.log\""
echo "Verwijderen:        launchctl bootout gui/$(id -u)/$LABEL && rm \"$PLIST_DST\""
