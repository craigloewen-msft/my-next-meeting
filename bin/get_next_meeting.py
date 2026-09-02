#!/usr/bin/env python3
"""
Stage 2 backend: non-interactive poll used by the bar widget (BarWidget.qml).

Always exits 0 and prints exactly one JSON line, so the QML Process handler
can parse it unconditionally:

  {"ok": true,  "subject": "...", "start": "2024-01-01T12:00:00+00:00",
   "end": "..."}
  {"ok": true,  "subject": null}                 # nothing upcoming
  {"ok": false, "error": "not_authenticated"}     # (Graph mode) run test_auth.py first
  {"ok": false, "error": "blocked_by_org"}         # (Graph mode) org policy blocked sign-in
  {"ok": false, "error": "fetch_error", "detail": "..."}
  {"ok": false, "error": "graph_error", "detail": "..."}

Two independent backends, picked automatically based on config.json:

  - ICS feed (preferred if "ics_url" is set): just an HTTPS GET of a published
    Outlook calendar link. No sign-in, no OAuth, so it isn't affected by
    Conditional Access policies that block device-code flow / unmanaged
    devices (common on locked-down corporate tenants). See README.md for how
    to get this URL from Outlook on the web.

  - Microsoft Graph (used if no "ics_url" configured): silent/cached OAuth
    token only — never starts a device-code flow itself. Run test_auth.py
    once by hand first to populate the token cache.
"""
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests

CONFIG_FILE = Path.home() / ".config" / "omarchy" / "next-meeting" / "config.json"
CACHE_FILE = Path.home() / ".cache" / "omarchy" / "next-meeting" / "token_cache.bin"

SCOPES = ["Calendars.Read", "User.Read"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
DEFAULT_TENANT_ID = "common"

BLOCKED_MARKERS = ("AADSTS50105", "AADSTS53000", "AADSTS53003", "AADSTS65004", "AADSTS90094")

WINDOW_DAYS = 14


def emit(payload):
    print(json.dumps(payload), flush=True)
    sys.exit(0)


def load_config():
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text())
    except json.JSONDecodeError:
        return {}


# ---------------------------------------------------------------------------
# ICS feed backend (no auth)
# ---------------------------------------------------------------------------

ICS_FETCH_ATTEMPTS = 3
ICS_FETCH_RETRY_DELAY_SECONDS = 2
# A bare "python-requests/x.y" UA is occasionally rejected by Office 365's
# front doors; a normal-looking UA avoids that class of transient block.
ICS_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) omarchy-next-meeting/1.0"


def fetch_next_meeting_ics(ics_url: str):
    import icalendar
    import recurring_ical_events
    import time

    last_error = None
    resp = None
    for attempt in range(1, ICS_FETCH_ATTEMPTS + 1):
        try:
            resp = requests.get(
                ics_url,
                timeout=15,
                headers={"User-Agent": ICS_USER_AGENT},
            )
            resp.raise_for_status()
            last_error = None
            break
        except requests.RequestException as e:
            last_error = e
            if attempt < ICS_FETCH_ATTEMPTS:
                time.sleep(ICS_FETCH_RETRY_DELAY_SECONDS)

    if last_error is not None:
        emit({"ok": False, "error": "fetch_error", "detail": str(last_error)})

    try:
        calendar = icalendar.Calendar.from_ical(resp.content)
    except ValueError as e:
        emit({"ok": False, "error": "fetch_error", "detail": f"Bad ICS data: {e}"})

    now = datetime.now(timezone.utc)
    window_end = now + timedelta(days=WINDOW_DAYS)
    events = recurring_ical_events.of(calendar).between(now, window_end)

    candidates = []
    for ev in events:
        status = str(ev.get("STATUS", "")).upper()
        if status == "CANCELLED":
            continue
        transp = str(ev.get("TRANSP", "")).upper()
        if transp == "TRANSPARENT":  # marked "Free" in Outlook
            continue
        start = ev.get("DTSTART").dt
        end = ev.get("DTEND").dt if ev.get("DTEND") else start
        # All-day entries use date (not datetime) values - skip those.
        if not isinstance(start, datetime):
            continue
        if start.tzinfo is None:
            start = start.replace(tzinfo=timezone.utc)
        if isinstance(end, datetime) and end.tzinfo is None:
            end = end.replace(tzinfo=timezone.utc)
        if start < now and isinstance(end, datetime) and end < now:
            continue
        candidates.append((start, end, str(ev.get("SUMMARY", "")) or "(No subject)"))

    if not candidates:
        emit({"ok": True, "subject": None})

    candidates.sort(key=lambda c: c[0])
    start, end, subject = candidates[0]
    emit({
        "ok": True,
        "subject": subject,
        "start": start.isoformat(),
        "end": end.isoformat() if isinstance(end, datetime) else start.isoformat(),
    })


# ---------------------------------------------------------------------------
# Microsoft Graph backend (silent OAuth token only)
# ---------------------------------------------------------------------------

def get_token_silent(client_id: str, tenant_id: str):
    import msal

    if not CACHE_FILE.exists():
        emit({"ok": False, "error": "not_authenticated"})

    cache = msal.SerializableTokenCache()
    cache.deserialize(CACHE_FILE.read_text())

    app = msal.PublicClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    if not accounts:
        emit({"ok": False, "error": "not_authenticated"})

    result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if cache.has_state_changed:
        CACHE_FILE.write_text(cache.serialize())
        CACHE_FILE.chmod(0o600)

    if not result or "access_token" not in result:
        desc = str((result or {}).get("error_description") or "")
        if any(marker in desc for marker in BLOCKED_MARKERS):
            emit({"ok": False, "error": "blocked_by_org", "detail": desc})
        emit({"ok": False, "error": "not_authenticated"})
    return result["access_token"]


def fetch_next_meeting_graph(token: str):
    now = datetime.now(timezone.utc)
    window_end = now + timedelta(days=WINDOW_DAYS)
    try:
        resp = requests.get(
            f"{GRAPH_ROOT}/me/calendarView",
            headers={
                "Authorization": "Bearer " + token,
                "Prefer": 'outlook.timezone="UTC"',
            },
            params={
                "startDateTime": now.strftime("%Y-%m-%dT%H:%M:%S"),
                "endDateTime": window_end.strftime("%Y-%m-%dT%H:%M:%S"),
                "$select": "subject,start,end,isCancelled,showAs,isAllDay",
                "$orderby": "start/dateTime",
                "$top": "10",
            },
            timeout=10,
        )
        resp.raise_for_status()
    except requests.RequestException as e:
        status = getattr(e.response, "status_code", None)
        if status in (401, 403):
            emit({"ok": False, "error": "blocked_by_org", "detail": str(e)})
        emit({"ok": False, "error": "graph_error", "detail": str(e)})

    events = resp.json().get("value", [])
    for ev in events:
        if ev.get("isCancelled"):
            continue
        if ev.get("showAs") == "free":
            continue
        if ev.get("isAllDay"):
            continue
        emit({
            "ok": True,
            "subject": ev.get("subject") or "(No subject)",
            "start": ev["start"]["dateTime"] + "Z",
            "end": ev["end"]["dateTime"] + "Z",
        })

    emit({"ok": True, "subject": None})


def main():
    config = load_config()
    ics_url = (config.get("ics_url") or "").strip()

    if ics_url:
        fetch_next_meeting_ics(ics_url)
        return

    client_id = (config.get("client_id") or "").strip() or DEFAULT_CLIENT_ID
    tenant_id = (config.get("tenant_id") or "").strip() or DEFAULT_TENANT_ID
    token = get_token_silent(client_id, tenant_id)
    fetch_next_meeting_graph(token)


if __name__ == "__main__":
    main()
