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
import requests

CONFIG_DIR = Path.home() / ".config" / "omarchy" / "next-meeting"
CONFIG_FILE = CONFIG_DIR / "config.json"
CACHE_DIR = Path.home() / ".cache" / "omarchy" / "next-meeting"
CACHE_FILE = CACHE_DIR / "token_cache.bin"

SCOPES = ["Calendars.Read", "User.Read"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"

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
    """
    if not CONFIG_FILE.exists():
        return DEFAULT_CLIENT_ID, DEFAULT_TENANT_ID
    try:
        data = json.loads(CONFIG_FILE.read_text())
    except json.JSONDecodeError as e:
        print(f"Could not parse {CONFIG_FILE}: {e}", file=sys.stderr)
        sys.exit(1)
    client_id = (data.get("client_id") or "").strip() or DEFAULT_CLIENT_ID
    tenant_id = (data.get("tenant_id") or "").strip() or DEFAULT_TENANT_ID
    return client_id, tenant_id


def build_app(client_id: str, tenant_id: str):
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache = msal.SerializableTokenCache()
    if CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text())

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )
    return app, cache


def save_cache(cache: "msal.SerializableTokenCache"):
    if cache.has_state_changed:
        CACHE_FILE.write_text(cache.serialize())
        CACHE_FILE.chmod(0o600)


def get_token(app: msal.PublicClientApplication, cache):
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            print(f"Failed to create device flow: {flow}", file=sys.stderr)
            sys.exit(1)
        print(flow["message"])
        sys.stdout.flush()
        result = app.acquire_token_by_device_flow(flow)

    save_cache(cache)

    if "access_token" not in result:
        err = str(result.get("error") or "")
        desc = str(result.get("error_description") or "")
        print(f"Auth failed: {err}: {desc}", file=sys.stderr)
        blocked_markers = ("AADSTS50105", "AADSTS53000", "AADSTS53003", "AADSTS65004", "AADSTS90094")
        if any(marker in desc for marker in blocked_markers) or err == "access_denied":
            print("\n" + BLOCKED_HINT.format(config_file=CONFIG_FILE), file=sys.stderr)
        sys.exit(1)

    return result["access_token"]


def fetch_next_events(token: str, top: int = 5):
    resp = requests.get(
        f"{GRAPH_ROOT}/me/events",
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
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json().get("value", [])


def main():
    client_id, tenant_id = load_config()
    app, cache = build_app(client_id, tenant_id)
    token = get_token(app, cache)

    print("\n✅ Auth succeeded. Fetching events...\n")
    events = fetch_next_events(token)
    if not events:
        print("No upcoming events found.")
        return

    for ev in events:
        print(f"- {ev['subject']!r} starts {ev['start']['dateTime']} ({ev['start']['timeZone']})")


if __name__ == "__main__":
    main()
