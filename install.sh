#!/usr/bin/env bash
# MANDATORY post-install setup for the craig.next-meeting bar widget.
#
# The widget shells out to bin/get_next_meeting.py on every poll, and that
# script imports msal / requests / icalendar / recurring-ical-events. None
# of those are guaranteed to exist on an Omarchy system, and installing
# them system-wide is not something a bar widget gets to do. So this script
# builds a private virtualenv at ./venv, and BarWidget.qml runs
# ./venv/bin/python by absolute path - nothing else. Skip this step and the
# widget just reports that the backend failed to start.
#
# Run it once after cloning, and again after pulling changes that touch
# requirements.lock. It is idempotent.
#
# Dependencies come from requirements.lock, not requirements.txt:
#   - requirements.txt holds the human-edited lower bounds (msal>=1.28 ...).
#     It is the *input* to the lock, never installed from directly.
#   - requirements.lock pins every direct and transitive package to an
#     exact version plus the SHA-256 of each permitted artifact.
# Installing with --require-hashes means pip refuses anything whose bytes
# don't match, so a hijacked release or a tampered index fails the install
# rather than landing in a script that reads your calendar.
#
# Note there is deliberately no `pip install --upgrade pip` here. That
# fetched an unpinned pip from the network on every run - an unverified
# package installed ahead of, and with the same privileges as, the very
# step meant to be hash-verified. The pip that ships with the interpreter's
# venv module is what we use; if it is too old to understand the lock file,
# the install fails loudly instead of silently self-updating.
set -euo pipefail

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV="$DIR/venv"
PY="$VENV/bin/python"

if [[ ! -x "$PY" ]]; then
  echo "Creating virtualenv at $VENV ..."
  python3 -m venv "$VENV"
fi

if [[ ! -f "$DIR/requirements.lock" ]]; then
  echo "error: requirements.lock is missing; cannot do a verified install." >&2
  echo "       See the header of requirements.lock for how to regenerate it." >&2
  exit 1
fi

echo "Installing pinned, hash-verified dependencies ..."
# --require-hashes  : refuse any requirement lacking a matching hash.
# --no-deps         : the lock is already the complete, closed dependency
#                     set; letting pip resolve again could pull something
#                     unpinned. (pip implies this under --require-hashes,
#                     but stating it keeps the intent obvious.)
# --only-binary     : never build an sdist, which would execute arbitrary
#                     setup.py code from the index at install time.
"$PY" -m pip install \
  --quiet \
  --disable-pip-version-check \
  --require-hashes \
  --no-deps \
  --only-binary :all: \
  -r "$DIR/requirements.lock"

# Sanity-check that the interpreter the widget will actually invoke can
# import everything it needs. Catches a half-finished install now rather
# than as a mystery error on the bar later.
"$PY" - <<'PYEOF'
import importlib.util, sys
missing = [m for m in ("msal", "requests", "icalendar", "recurring_ical_events")
           if importlib.util.find_spec(m) is None]
if missing:
    sys.exit("error: missing after install: " + ", ".join(missing))
PYEOF

echo
echo "Done. Next steps:"
echo
echo "  Option A (recommended, no sign-in): publish an Outlook ICS calendar"
echo "  link and save it to ~/.config/omarchy/next-meeting/config.json - see"
echo "  README.md for the exact steps."
echo
echo "  Option B (Microsoft Graph): run"
echo "    $PY $DIR/bin/test_auth.py"
echo "  and follow the device-code sign-in prompt."
echo
echo "Then reload the plugin:"
echo "  omarchy-shell shell rescanPlugins"
