#!/usr/bin/env bash
# One-time setup: creates a private virtualenv next to this script and
# installs the Python dependencies (msal, requests, icalendar,
# recurring-ical-events) into it. Never touches system Python packages.
set -euo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [[ ! -d "$DIR/venv" ]]; then
  echo "Creating virtualenv at $DIR/venv ..."
  python3 -m venv "$DIR/venv"
fi

echo "Installing dependencies ..."
"$DIR/venv/bin/pip" install --quiet --upgrade pip
"$DIR/venv/bin/pip" install --quiet -r "$DIR/requirements.txt"

echo
echo "Done. Next steps:"
echo
echo "  Option A (recommended, no sign-in): publish an Outlook ICS calendar"
echo "  link and save it to ~/.config/omarchy/next-meeting/config.json - see"
echo "  README.md for the exact steps."
echo
echo "  Option B (Microsoft Graph): run"
echo "    $DIR/venv/bin/python $DIR/bin/test_auth.py"
echo "  and follow the device-code sign-in prompt."
echo
echo "Then reload the plugin:"
echo "  omarchy-shell shell rescanPlugins"
