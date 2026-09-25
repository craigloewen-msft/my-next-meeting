# Omarchy Next Meeting

A bar widget for [Omarchy](https://omarchy.org/) that shows a live countdown
to your next Outlook / Microsoft 365 calendar meeting, e.g. `Standup in 23m`
or `▶ Standup` while it's happening.

## Features

- Live countdown on the bar, ticking every 15s, re-checking your calendar
  every 20 minutes (or instantly whenever `config.json` changes).
- **Click the widget for a timeline of your whole day** — a popup that lays
  every meeting out against an hourly scale, so a block's height *is* its
  duration and the empty space between blocks *is* your free time.
  Overlapping meetings sit side by side in their own columns, a red rule
  marks where "now" falls, and anything already running is highlighted as
  in-progress. Meetings that have already finished stay on the timeline,
  dimmed, so you can see the shape of the day you've had as well as the one
  ahead. It opens centred on "now" and scrolls to either end; hover any
  block for its full title and times. Middle- or right-click forces an
  immediate refresh.
- Two independent, auto-selected backends:
  - **Published ICS calendar link** (recommended) — a plain HTTPS fetch, no
    sign-in at all. Works even when your organization's Conditional Access
    policies block OAuth apps on unmanaged devices.
  - **Microsoft Graph** — device-code sign-in using Microsoft's own public
    "Graph Command Line Tools" client ID, so no Azure app registration is
    required for most people.
- Skips cancelled events, all-day events, and events marked "Free".
- Bounded and defensive by construction: every fetch, parse and subprocess
  runs under an explicit size and time limit, credentials on disk are
  handled without following symlinks, and error text is redacted before it
  can reach your screen. See [Security posture](#security-posture).

## Requirements

- Omarchy with Quickshell plugin support.
- Python 3.10+ (used only for the polling backend; installed into a private
  virtualenv, no system packages touched).
- `coreutils` (for `timeout`, used to bound the backend process) — present
  on any normal Arch/Omarchy system.

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

### Required post-install setup (do not skip)

**Running `install.sh` is mandatory**, however you installed the plugin.
The widget shells out to `venv/bin/python` and does nothing else — until
that virtualenv exists, the bar shows an error instead of a countdown. The
plugin deliberately never falls back to a system Python, so that what it
runs is always the exact, verified set of packages installed here.

```bash
~/.config/omarchy/plugins/craig.next-meeting/install.sh
```

The script:

- creates a private virtualenv at `venv/` inside the plugin folder (nothing
  is installed system-wide, and no system packages are touched);
- installs **only** from `requirements.lock`, which pins every package —
  direct *and* transitive — to an exact version and a set of SHA-256
  hashes, using `pip install --require-hashes --no-deps --only-binary :all:`.
  If any downloaded artifact doesn't match its recorded hash, pip aborts and
  installs nothing;
- does **not** upgrade `pip` (an unpinned `pip install --upgrade pip` would
  pull an unverified package from the live index, which is exactly what the
  lockfile exists to prevent);
- finishes with an import check, so a partial install fails loudly at
  install time rather than silently on the bar later.

Re-run it any time to repair or rebuild the environment; it is idempotent.

#### Updating dependencies

`requirements.txt` is the human-edited *input* (loose lower bounds);
`requirements.lock` is the generated, hash-pinned artifact that is actually
installed. After changing `requirements.txt`, regenerate the lock:

```bash
pip install pip-tools
pip-compile --generate-hashes --output-file=requirements.lock requirements.txt
```

Commit both files together, and never hand-edit `requirements.lock`.

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
5. Done — the widget picks this up automatically within ~20 minutes, but it
   also watches `config.json`, so saving the file refreshes the bar right
   away.

If "Publish a calendar" is greyed out or missing, your org has disabled
external calendar publishing — use option B instead.

> **Security note:** the published link is a bearer secret — anyone with
> the URL can read your calendar. Treat `config.json` like a credential
> file. The plugin stores it in a `0700` directory, reads it as `0600`
> without following symlinks, sends it only to the host it names over
> HTTPS, and strips it out of any error message before that message can be
> shown or logged. See [Security posture](#security-posture).

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

## Using the widget

| Action | Result |
| --- | --- |
| Hover (bar) | Tooltip with the next meeting's full title and time. |
| Left click | Opens **your whole day as a timeline** — an hourly scale with every meeting drawn as a block whose height matches its real length, so gaps and back-to-backs are obvious at a glance. Overlapping meetings share the track in side-by-side columns, meetings that have already ended are dimmed, and a red rule shows where you are right now. Click again to close. |
| Hover (timeline block) | Shows that meeting's full title, clock range, duration and lead time in the detail line under the timeline — a 30-minute block is too short to hold all of it. |
| Scroll (timeline) | Moves through the day. It opens centred on "now"; scroll up for the morning, down for the evening. |
| Left click (signed out) | Opens a terminal for device-code sign-in instead. |
| Middle / right click | Forces an immediate refresh. |

The timeline covers the full calendar day, midnight to midnight, at a fixed
scale — so 9AM is always the same distance from 10AM, and a block's height
is always its real duration. That scale is set so a 30-minute meeting — the
most common kind — still gets a full-height block, and shorter meetings than
that never disappear either.

A whole day at a readable scale is much taller than the popup, so the
timeline scrolls. It opens with "now" about 40% down the viewport, so what's
next is front and centre with the morning a scroll up and the evening a
scroll down; a scrollbar on the right and a hint line underneath say which
way there's more to see. Reopening the popup re-centres on "now", but the
15s tick and background refreshes will not yank the view out from under you
while you're reading.

The agenda comes down in the same poll as the countdown, so opening it costs
no extra work and no extra network request. Both the countdown and the
timeline are recomputed every 15s, so meetings move from upcoming to ended
as the day goes on and the "now" marker keeps sliding, without waiting for
the next poll.

## Security posture

The calendar URL is a bearer credential and the feed itself is
attacker-influenced data (anyone who can put a meeting on your calendar
controls part of it), so the plugin is written to bound and distrust both.

**Credentials on disk** (`bin/secure_io.py`) — `config.json` and the Graph
token cache are opened relative to a verified `0700` directory descriptor
with `O_NOFOLLOW`, then checked *on the descriptor* for being a regular
file, owned by you, with no extra hard links, and within a byte budget.
Symlinks, FIFOs, foreign owners and oversized files are refused rather than
followed. The token cache is written exclusively at `0600` to a temporary
file, fsynced, and atomically renamed into place, so its contents are never
briefly world-readable and never half-written.

**Network** (`bin/net.py`) — HTTPS is required and redirects are followed
manually, bounded in number, refused on downgrade to HTTP, and stripped of
the `Authorization` header when the host changes. Redirect bodies are
discarded rather than read: `requests` will otherwise buffer a redirect
response in full inside `Session.send`, before any cap can see it. Responses
are streamed into a hard byte ceiling (checked against `Content-Length` *and*
enforced while reading), the content type is validated, and one wall-clock
deadline covers the whole exchange including retries.

**Untrusted calendar data** (`bin/get_next_meeting.py`) — the raw feed is
capped in bytes and in `VEVENT` count before parsing. Recurrence rules that
expand faster than any real meeting ever could — sub-hourly frequencies, and
hourly ones with no end date — are refused on the raw bytes, before the
parser sees them, because the expansion itself is the cost. What remains is
expanded only inside a closed look-ahead window, never open-endedly, and in
bounded chunks, so the occurrence cap applies to a running total instead of
to an allocation that already happened. Parsing additionally runs under a
`SIGALRM` deadline, since a single rule can otherwise spend unbounded time
inside one iteration and never yield for a clock check. Subjects and error
details are length-limited before display.

**Error text** — every error path funnels through one redaction step that
strips URLs with secret paths, `Authorization` headers, and URL userinfo,
including the bare-path form that `urllib3` uses in connection errors. The
QML side redacts again before drawing, and the helper's stderr is counted
but never displayed, since a stray traceback could quote a local variable
holding the URL.

**The backend process** (`BarWidget.qml`) — the poll is wrapped in
`timeout`, which signals the whole process group so grandchildren can't
outlive it, with a QML watchdog behind it escalating TERM → KILL and only
then releasing the handle. Output is read as a stream, not buffered whole,
with separate byte ceilings on stdout and stderr, so a runaway helper can't
grow the shell's memory.

## Uninstalling


```bash
omarchy plugin remove craig.next-meeting
rm -rf ~/.config/omarchy/next-meeting ~/.cache/omarchy/next-meeting
```

(The second command removes your saved config and cached sign-in token; skip
it if you plan to reinstall later.)

## How it works / repo layout

- `manifest.json` — Omarchy plugin manifest.
- `BarWidget.qml` — the bar widget itself: polls the backend script every 20
  minutes via `Quickshell.Io.Process` (under a `timeout` wrapper and a QML
  watchdog), recomputes the on-screen countdown and today's agenda every
  15s, lays the whole day out as a proportional timeline (hour grid,
  duration-scaled blocks, column packing for overlaps, dimmed past
  meetings, a live "now" rule, scrolling and a hover detail line) in the
  click-through popup, and watches `config.json` for instant refresh on
  change.
- `bin/get_next_meeting.py` — non-interactive poll script. Always exits 0
  and prints exactly one JSON line so the widget never has to handle a
  crash. Auto-selects the ICS or Graph backend based on `config.json`, and
  emits the next meeting, the rest of today's agenda, and today's already
  finished meetings in that one line.
- `bin/secure_io.py` — the only path to credential files on disk: verified
  private directory, no-follow bounded reads, atomic `0600` writes.
- `bin/net.py` — the only path to the network: HTTPS-only bounded fetches,
  plus the redaction used on every error string.
- `bin/test_auth.py` — interactive, one-time device-code sign-in for the
  Graph backend. Never run automatically by the widget.
- `install.sh` — creates the private virtualenv and installs from the lock.
- `requirements.txt` — human-edited dependency input.
- `requirements.lock` — generated, hash-pinned, and what is actually
  installed.

## License

MIT — see [LICENSE](LICENSE).
