# Omarchy Next Meeting

A bar widget for [Omarchy](https://omarchy.org/) that shows a live countdown
to your next Outlook / Microsoft 365 calendar meeting, e.g. `Standup in 23m`
or `▶ Standup` while it's happening.

## Features

- Live countdown on the bar, ticking every 15s, re-checking your calendar
  every 3 minutes (or instantly whenever `config.json` changes).
- Two independent, auto-selected backends:
  - **Published ICS calendar link** (recommended) — a plain HTTPS fetch, no
    sign-in at all. Works even when your organization's Conditional Access
    policies block OAuth apps on unmanaged devices.
  - **Microsoft Graph** — device-code sign-in using Microsoft's own public
    "Graph Command Line Tools" client ID, so no Azure app registration is
    required for most people.
- Skips cancelled events, all-day events, and events marked "Free".
- Click the widget to force an immediate refresh (or to sign in, if not
  authenticated yet).

## Requirements

- Omarchy with Quickshell plugin support.
- Python 3.10+ (used only for the polling backend; installed into a private
  virtualenv, no system packages touched).

## Installation

### Via the Omarchy plugin system

```bash
omarchy plugin add https://github.com/craigloewen-msft/my-next-meeting.git --enable
```

### Manual installation

```bash
git clone https://github.com/craigloewen-msft/my-next-meeting.git \
  ~/.config/omarchy/plugins/craig.next-meeting

~/.config/omarchy/plugins/craig.next-meeting/install.sh

omarchy plugin enable craig.next-meeting --section right
```

`install.sh` creates a private virtualenv at `venv/` inside the plugin
folder and installs the Python dependencies (`msal`, `requests`,
`icalendar`, `recurring-ical-events`) from `requirements.txt`. Nothing is
installed system-wide.

## Configuration

The widget starts in a "not signed in" state until you set up one of the two
backends below. Config lives at `~/.config/omarchy/next-meeting/config.json`
(created by you, not by the plugin).

### Option A — published ICS calendar link (recommended, no sign-in)

This is a plain HTTPS GET of a secret calendar URL — no OAuth, no
device-code flow, so it isn't affected by Conditional Access policies that
block unmanaged/non-compliant devices (common on locked-down corporate
tenants, including Microsoft's own).

1. Go to https://outlook.office.com → gear icon (Settings) → "View all
   Outlook settings" → **Calendar** → **Shared calendars**.
2. Under "Publish a calendar", pick your calendar, set permissions to
   **"Can view all details"**, click **Publish**.
3. Copy the **ICS** link it gives you (ends in `.ics`).
4. Save it to `~/.config/omarchy/next-meeting/config.json`:
   ```json
   { "ics_url": "https://outlook.office365.com/owa/calendar/....ics" }
   ```
5. Done — the widget picks this up automatically within ~3 minutes (or
   click it to force a refresh immediately). No sign-in step needed.

If "Publish a calendar" is greyed out or missing, your org has disabled
external calendar publishing — use option B instead.

> **Security note:** the published link is a bearer secret — anyone with
> the URL can read your calendar. Treat `config.json` like a credential
> file (it's created with your regular user permissions; nothing in this
> repo transmits it anywhere except directly to Microsoft's servers).

### Option B — Microsoft Graph (used automatically if no `ics_url` is set)

```bash
~/.config/omarchy/plugins/craig.next-meeting/venv/bin/python \
  ~/.config/omarchy/plugins/craig.next-meeting/bin/test_auth.py
```

It prints a `https://microsoft.com/devicelogin` URL + code — open it, sign
in, and approve the `Calendars.Read` / `User.Read` permissions.

By default this uses Microsoft's own public "Microsoft Graph Command Line
Tools" client ID (the same one `Connect-MgGraph`/PowerShell uses), which
Microsoft pre-registers in every commercial tenant — so no Azure app
registration of your own is needed.

**This is commonly blocked on locked-down corporate tenants** via
Conditional Access requiring a compliant/managed device. If sign-in fails
with something like *"Your sign-in was successful but does not meet the
criteria to access this resource"*, that's this policy — no client ID swap
will fix it, since Conditional Access policies normally apply to all cloud
apps, not just this one. Use option A instead.

If you have your own Azure app registration (or your admin gives you one),
override the default client in `config.json`:
```json
{ "client_id": "<Application (client) ID>", "tenant_id": "common" }
```

## Uninstalling

```bash
omarchy plugin remove craig.next-meeting
rm -rf ~/.config/omarchy/next-meeting ~/.cache/omarchy/next-meeting
```

(The second command removes your saved config and cached sign-in token; skip
it if you plan to reinstall later.)

## How it works / repo layout

- `manifest.json` — Omarchy plugin manifest.
- `BarWidget.qml` — the bar widget itself: polls the backend script every 3
  minutes via `Quickshell.Io.Process`, recomputes the on-screen countdown
  every 15s, and watches `config.json` for instant refresh on change.
- `bin/get_next_meeting.py` — non-interactive poll script. Always exits 0
  and prints exactly one JSON line so the widget never has to handle a
  crash. Auto-selects the ICS or Graph backend based on `config.json`.
- `bin/test_auth.py` — interactive, one-time device-code sign-in for the
  Graph backend. Never run automatically by the widget.
- `install.sh` / `requirements.txt` — creates the private virtualenv.

## License

MIT — see [LICENSE](LICENSE).
