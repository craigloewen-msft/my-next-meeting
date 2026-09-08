#!/usr/bin/env python3
"""
Stage 1: Standalone auth + Graph API smoke test for the "next meeting" bar widget.

Run this directly to verify device-code login and calendar access work
before wiring anything into the Quickshell plugin:

    ~/.config/omarchy/plugins/craig.next-meeting/venv/bin/python \
        ~/.config/omarchy/plugins/craig.next-meeting/bin/test_auth.py

On first run it will print a Microsoft device-login URL + code. After you
sign in once, the token is cached on disk and later runs are silent.

By default this signs in using Microsoft's own public "Graph Command Line
Tools" client ID (see DEFAULT_CLIENT_ID below) — the same one used by
Connect-MgGraph/PowerShell — so most people need NO Azure app registration
of their own. This matters because many organizations disable "users can
register applications" in Entra ID. If your org has instead blocked THAT
app (via Conditional Access or an assignment requirement), see BLOCKED_HINT
below for what to do, including a no-OAuth ICS-calendar fallback.
"""
import json
import sys
from pathlib import Path

import msal

import net
import secure_io

CONFIG_DIR = Path.home() / ".config" / "omarchy" / "next-meeting"
CONFIG_NAME = "config.json"
CONFIG_FILE = CONFIG_DIR / CONFIG_NAME
CACHE_DIR = Path.home() / ".cache" / "omarchy" / "next-meeting"
CACHE_NAME = "token_cache.bin"
CACHE_FILE = CACHE_DIR / CACHE_NAME

SCOPES = ["Calendars.Read", "User.Read"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

# Bounds on the smoke-test request, matching get_next_meeting.py. Even a
# hand-run diagnostic reads a remote response into this process, so it gets
# the same byte cap, content-type check and redirect policy.
MAX_GRAPH_BYTES = 2 * 1024 * 1024
GRAPH_FETCH_SECONDS = 30
GRAPH_CONTENT_TYPES = frozenset({"application/json"})

# "Microsoft Graph Command Line Tools" — a first-party public client app that
# Microsoft pre-registers in every commercial tenant (same one used by
# Connect-MgGraph / the Graph CLI). Using it means most people need ZERO
# Azure app registration of their own, which matters because many orgs
# disable "users can register applications". No client secret is involved —
# it's a public client using device-code flow, same trust level as any
# Microsoft first-party app.
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
DEFAULT_TENANT_ID = "common"

BLOCKED_HINT = """
Sign-in was blocked by your organization's policy (Conditional Access, an
"assignment required" restriction on the "Microsoft Graph Command Line
Tools" enterprise app, or a consent policy requiring admin approval).

Options:
  1. Ask your IT/Entra admin to approve/assign the "Microsoft Graph Command
     Line Tools" enterprise app for your account, or grant admin consent for
     the Calendars.Read / User.Read delegated permissions.
  2. If they won't, ask them to register a small first-party app for you
     instead (App registrations -> New registration -> enable "Allow public
     client flows" -> API permissions: Calendars.Read, User.Read delegated).
     Once you have a client ID, put it in:
         {config_file}
     as: {{"client_id": "<id>", "tenant_id": "<your tenant id or 'common'>"}}
  3. As a no-OAuth fallback, Outlook can publish a read-only ICS calendar
     link (Outlook web -> Settings -> Calendar -> Shared calendars ->
     Publish a calendar). That URL can be polled directly without any
     Azure app or sign-in — ask me to switch the widget to that approach.
""".strip()


def load_config():
    """
    config.json is optional. If present it can override client_id/tenant_id
    (e.g. to point at your own app registration, or a specific tenant ID
    instead of "common"). If absent, fall back to Microsoft's own public
    "Graph Command Line Tools" client — no registration needed.

    Read through secure_io so a symlinked, foreign-owned or absurdly large
    config.json is refused rather than followed — this file can name the
    identity we authenticate as, so it is worth being fussy about.
    """
    try:
        with secure_io.PrivateDir(CONFIG_DIR, create=False) as d:
            data = d.read_json(CONFIG_NAME, secure_io.MAX_CONFIG_BYTES)
    except FileNotFoundError:
        return DEFAULT_CLIENT_ID, DEFAULT_TENANT_ID
    except (secure_io.SecureIOError, OSError) as e:
        print(f"Refusing to use {CONFIG_FILE}: {e}", file=sys.stderr)
        sys.exit(1)
    client_id = str(data.get("client_id") or "").strip() or DEFAULT_CLIENT_ID
    tenant_id = str(data.get("tenant_id") or "").strip() or DEFAULT_TENANT_ID
    return client_id, tenant_id


def build_app(client_id: str, tenant_id: str):
    """
    Build the MSAL app around a cache directory we have verified is ours.

    The returned PrivateDir stays open for the life of the run so the cache
    write at the end lands in the same directory we checked, not in
    whatever the path happens to point at by then.
    """
    cache_dir = secure_io.PrivateDir(CACHE_DIR, create=True)
    cache = msal.SerializableTokenCache()
    try:
        serialized = cache_dir.read_text(CACHE_NAME, secure_io.MAX_TOKEN_CACHE_BYTES)
    except secure_io.SecureIOError as e:
        cache_dir.close()
        print(f"Refusing to use {CACHE_FILE}: {e}", file=sys.stderr)
        sys.exit(1)
    if serialized:
        cache.deserialize(serialized)

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )
    return app, cache, cache_dir


def save_cache(cache: "msal.SerializableTokenCache", cache_dir: "secure_io.PrivateDir"):
    """
    Persist the token cache atomically at mode 0600.

    The old write-then-chmod left the file world-readable for the window
    between the two calls; secure_io creates the temp file at 0600 up front
    and renames it into place, so there is no such window and no risk of a
    half-written cache if this run is interrupted.
    """
    if not cache.has_state_changed:
        return
    try:
        cache_dir.write_text(CACHE_NAME, cache.serialize())
    except (secure_io.SecureIOError, OSError) as e:
        print(f"Could not save token cache: {net.redact(e)}", file=sys.stderr)


def get_token(app: msal.PublicClientApplication, cache, cache_dir):
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            # The flow dict echoes request details back; redact before printing.
            print(f"Failed to create device flow: {net.redact(flow)}", file=sys.stderr)
            sys.exit(1)
        print(flow["message"])
        sys.stdout.flush()
        result = app.acquire_token_by_device_flow(flow)

    save_cache(cache, cache_dir)

    if "access_token" not in result:
        err = str(result.get("error") or "")
        desc = net.redact(result.get("error_description") or "")
        print(f"Auth failed: {err}: {desc}", file=sys.stderr)
        blocked_markers = ("AADSTS50105", "AADSTS53000", "AADSTS53003", "AADSTS65004", "AADSTS90094")
        if any(marker in desc for marker in blocked_markers) or err == "access_denied":
            print("\n" + BLOCKED_HINT.format(config_file=CONFIG_FILE), file=sys.stderr)
        sys.exit(1)

    return result["access_token"]


def fetch_next_events(token: str, top: int = 5):
    body, _ = net.fetch(
        f"{GRAPH_ROOT}/me/events",
        max_bytes=MAX_GRAPH_BYTES,
        timeout=15,
        headers={
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.timezone="UTC"',
        },
        params={
            "$select": "subject,start,end,isCancelled,showAs",
            "$orderby": "start/dateTime",
            "$top": str(top),
            "$filter": "isCancelled eq false",
        },
        allowed_content_types=GRAPH_CONTENT_TYPES,
        deadline=None,
    )
    parsed = json.loads(body)
    events = parsed.get("value") if isinstance(parsed, dict) else None
    return events if isinstance(events, list) else []


def main():
    client_id, tenant_id = load_config()
    app, cache, cache_dir = build_app(client_id, tenant_id)
    try:
        token = get_token(app, cache, cache_dir)
    finally:
        cache_dir.close()

    print("\n✅ Auth succeeded. Fetching events...\n")
    try:
        events = fetch_next_events(token)
    except net.FetchError as e:
        print(f"Could not fetch events: {e}", file=sys.stderr)
        sys.exit(1)
    if not events:
        print("No upcoming events found.")
        return

    for ev in events:
        start = ev.get("start") or {}
        print(f"- {ev.get('subject')!r} starts {start.get('dateTime')} "
              f"({start.get('timeZone')})")


if __name__ == "__main__":
    main()
