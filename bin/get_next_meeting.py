#!/usr/bin/env python3
"""
Stage 2 backend: non-interactive poll used by the bar widget (BarWidget.qml).

Always exits 0 and prints exactly one JSON line, so the QML Process handler
can parse it unconditionally:

  {"ok": true,  "subject": "...", "start": "2024-01-01T12:00:00+00:00",
   "end": "...", "agenda": [ ... ]}
  {"ok": true,  "subject": null, "agenda": []}     # nothing upcoming
  {"ok": false, "error": "not_authenticated"}      # (Graph mode) run test_auth.py first
  {"ok": false, "error": "blocked_by_org"}         # (Graph mode) org policy blocked sign-in
  {"ok": false, "error": "config_error", "detail": "..."}
  {"ok": false, "error": "fetch_error", "detail": "..."}
  {"ok": false, "error": "graph_error", "detail": "..."}

``agenda`` is every meeting still to come *today* in the local timezone -
including one already in progress - each with its full (untruncated by
display concerns) subject, start, and end. The widget renders it in a
popup when you click the bar item; putting it in the same JSON line as the
countdown means opening that popup costs nothing, because the data is
already on hand from the regular poll.

Two independent backends, picked automatically based on config.json:

  - ICS feed (preferred if "ics_url" is set): just an HTTPS GET of a published
    Outlook calendar link. No sign-in, no OAuth, so it isn't affected by
    Conditional Access policies that block device-code flow / unmanaged
    devices (common on locked-down corporate tenants). See README.md for how
    to get this URL from Outlook on the web.

  - Microsoft Graph (used if no "ics_url" configured): silent/cached OAuth
    token only — never starts a device-code flow itself. Run test_auth.py
    once by hand first to populate the token cache.

Everything this script reads is treated as untrusted:

  * config.json and the token cache are opened through :mod:`secure_io`,
    which refuses symlinks and foreign-owned files instead of following
    them, caps the read, and rewrites the cache atomically at mode 0600.
  * HTTP bodies are streamed through :mod:`net`, which enforces HTTPS,
    bounds redirects, and abandons a response the moment it crosses a byte
    cap - a calendar feed is remote input and gets no more trust than that.
  * ICS expansion is bounded three ways (event count, occurrence count, and
    a hard SIGALRM parsing deadline), because a recurrence rule is a tiny
    piece of text that can describe an unbounded number of events - a
    single ``RRULE:FREQ=SECONDLY`` would otherwise spin this poll forever
    inside the shell's process table.
  * Anything that reaches the ``detail`` field is passed through
    ``net.redact`` first. The published ICS URL is a bearer secret and
    ``requests`` puts it in almost every exception message it raises.
"""
import json
import re
import signal
import sys
import time
from contextlib import contextmanager
from datetime import datetime, time as dtime, timedelta, timezone
from pathlib import Path

import net
import secure_io

CONFIG_DIR = Path.home() / ".config" / "omarchy" / "next-meeting"
CONFIG_NAME = "config.json"
CACHE_DIR = Path.home() / ".cache" / "omarchy" / "next-meeting"
CACHE_NAME = "token_cache.bin"

SCOPES = ["Calendars.Read", "User.Read"]
GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
DEFAULT_TENANT_ID = "common"

BLOCKED_MARKERS = ("AADSTS50105", "AADSTS53000", "AADSTS53003", "AADSTS65004", "AADSTS90094")

# How far ahead to look for "the next meeting". The agenda popup only ever
# shows today, but the countdown should still work on a Friday afternoon.
WINDOW_DAYS = 14
# Widening ICS expansion windows, tried in order until one yields a meeting.
# Starting small keeps the common case (a meeting later today) cheap, while
# the last step still finds the next thing on a sparse calendar.
LOOKAHEAD_STEPS = (2, WINDOW_DAYS, 90)
# Largest span handed to a single between() call. Expansion is materialised
# per call, so this - not the lookahead - is what bounds peak memory.
CHUNK_DAYS = 7

# ---------------------------------------------------------------------------
# Resource limits
#
# Every one of these bounds untrusted remote input. They are deliberately
# far above any real calendar: the goal is to make a hostile or broken feed
# fail fast, not to second-guess a busy week.
# ---------------------------------------------------------------------------

# A year of a heavily-booked calendar publishes well under a megabyte.
MAX_ICS_BYTES = 8 * 1024 * 1024
# Graph replies with at most $top events and a fixed $select.
MAX_GRAPH_BYTES = 2 * 1024 * 1024
# Distinct VEVENT blocks in the raw feed, checked before any parsing.
MAX_VEVENTS = 5000
# Expanded occurrences pulled out of the recurrence engine.
MAX_OCCURRENCES = 2000
# Wall-clock ceiling on parse + recurrence expansion specifically.
ICS_PARSE_SECONDS = 12
# Wall-clock ceiling covering all fetch attempts including retries.
ICS_FETCH_SECONDS = 40
GRAPH_FETCH_SECONDS = 25
# Subjects are shown to the user; a megabyte-long SUMMARY is not a subject.
MAX_SUBJECT_CHARS = 200
# Entries returned in the agenda popup.
MAX_AGENDA_ITEMS = 25
# Error text shown to the user. Parse failures quote the input, so this is
# a cap on attacker-chosen text reaching the screen, not just on verbosity.
MAX_DETAIL_CHARS = 400

# RRULE lines in the raw feed, screened before the parser ever sees them.
_RRULE_RE = re.compile(rb"^RRULE:(.*)$", re.MULTILINE)

ICS_CONTENT_TYPES = frozenset({
    "text/calendar", "text/plain", "application/calendar",
    "application/ics", "text/ics", "application/octet-stream",
})
GRAPH_CONTENT_TYPES = frozenset({"application/json"})


def emit(payload):
    print(json.dumps(payload), flush=True)
    sys.exit(0)


def emit_error(error, detail=None):
    """
    Emit a failure. ``detail`` is redacted here, at the single choke point,
    rather than at each call site - it only takes one forgotten call to put
    a published-calendar URL on screen.

    It is also truncated here. Parser errors quote the offending property
    value, so a malformed feed can otherwise turn an error message into
    hundreds of kilobytes of attacker-chosen text that the widget would
    faithfully word-wrap into a tooltip.
    """
    payload = {"ok": False, "error": error}
    if detail:
        payload["detail"] = _clip(net.redact(detail), MAX_DETAIL_CHARS)
    emit(payload)


def _clip(text, limit):
    text = str(text)
    return text if len(text) <= limit else text[:limit] + "…"


def load_config():
    """
    Read config.json through the hardened reader.

    A missing file or directory is normal (Graph mode with the default
    client needs no config at all), so that returns {}. A file that fails
    the safety checks is *not* normal and surfaces as an error the user can
    see, instead of being silently treated as absent.
    """
    try:
        with secure_io.PrivateDir(CONFIG_DIR, create=False) as d:
            return d.read_json(CONFIG_NAME, secure_io.MAX_CONFIG_BYTES)
    except FileNotFoundError:
        return {}
    except secure_io.SecureIOError as e:
        emit_error("config_error", str(e))
    except OSError as e:
        emit_error("config_error", f"Could not read {CONFIG_NAME}: {e}")


# ---------------------------------------------------------------------------
# Shared shaping of results
# ---------------------------------------------------------------------------

def clean_subject(raw):
    """Collapse whitespace and cap length - this string goes on screen."""
    text = re.sub(r"\s+", " ", str(raw or "")).strip()
    if not text:
        return "(No subject)"
    if len(text) > MAX_SUBJECT_CHARS:
        return text[:MAX_SUBJECT_CHARS] + "…"
    return text


def local_day_end(now):
    """
    Local midnight at the end of today, as an aware (comparable) value.

    Built from the local calendar date rather than by adding a day to a
    fixed-offset datetime: astimezone() pins the offset that applies *now*,
    so on a DST-transition night the arithmetic version lands an hour either
    side of real midnight and the agenda gains or loses a meeting.
    """
    local_now = now.astimezone()
    tomorrow = local_now.date() + timedelta(days=1)
    return datetime.combine(tomorrow, dtime.min).astimezone()


def build_payload(candidates, now):
    """
    Turn a list of (start, end, subject) tuples into the widget's JSON.

    ``candidates`` need not be sorted and may contain events that already
    ended; both are handled here so the two backends don't each have to.
    """
    upcoming = sorted(
        (c for c in candidates if c[1] > now),
        key=lambda c: (c[0], c[1]),
    )
    if not upcoming:
        return {"ok": True, "subject": None, "agenda": []}

    day_end = local_day_end(now)
    agenda = [
        {
            "subject": subject,
            "start": start.isoformat(),
            "end": end.isoformat(),
            "inProgress": start <= now < end,
        }
        for start, end, subject in upcoming
        if start < day_end
    ][:MAX_AGENDA_ITEMS]

    start, end, subject = upcoming[0]
    return {
        "ok": True,
        "subject": subject,
        "start": start.isoformat(),
        "end": end.isoformat(),
        "agenda": agenda,
    }


class ParseTimeout(Exception):
    """The ICS parse/expansion budget ran out."""


@contextmanager
def time_limit(seconds, what):
    """
    Hard wall-clock limit on a block of pure-CPU work.

    A cooperative check between loop iterations is not enough here: the
    recurrence engine can spend an unbounded amount of time inside a
    *single* ``next()`` call while it winds a long-running rule forward to
    the present, never handing control back for us to look at a clock.
    SIGALRM interrupts it from outside, which is the only thing that
    actually bounds this.
    """
    if not hasattr(signal, "SIGALRM"):
        yield
        return

    def _fire(_signum, _frame):
        raise ParseTimeout(f"{what} took longer than {seconds}s")

    previous = signal.signal(signal.SIGALRM, _fire)
    signal.setitimer(signal.ITIMER_REAL, seconds)
    try:
        yield
    finally:
        signal.setitimer(signal.ITIMER_REAL, 0)
        signal.signal(signal.SIGALRM, previous)


# ---------------------------------------------------------------------------
# ICS feed backend (no auth)
# ---------------------------------------------------------------------------

ICS_FETCH_ATTEMPTS = 3
ICS_FETCH_RETRY_DELAY_SECONDS = 2
# A bare "python-requests/x.y" UA is occasionally rejected by Office 365's
# front doors; a normal-looking UA avoids that class of transient block.
ICS_USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) omarchy-next-meeting/1.0"


def fetch_ics_body(ics_url: str) -> bytes:
    """
    Fetch the feed, retrying transient failures within one overall budget.

    The retry loop shares a single deadline with the requests it makes, so
    three attempts can't stack up into three separate timeouts.
    """
    deadline = time.monotonic() + ICS_FETCH_SECONDS
    last_error = None

    for attempt in range(1, ICS_FETCH_ATTEMPTS + 1):
        try:
            body, _ = net.fetch(
                ics_url,
                max_bytes=MAX_ICS_BYTES,
                timeout=15,
                headers={"User-Agent": ICS_USER_AGENT},
                allowed_content_types=ICS_CONTENT_TYPES,
                deadline=deadline,
            )
            return body
        except net.FetchError as e:
            last_error = e
            # Retrying a rejected scheme, an oversized body or a bad
            # content type just burns the budget - the answer won't change.
            if not _worth_retrying(e):
                break
            if attempt < ICS_FETCH_ATTEMPTS:
                if time.monotonic() + ICS_FETCH_RETRY_DELAY_SECONDS >= deadline:
                    break
                time.sleep(ICS_FETCH_RETRY_DELAY_SECONDS)

    emit_error("fetch_error", str(last_error))


def _worth_retrying(err) -> bool:
    if err.status is not None:
        # 4xx is a durable "no"; 5xx and 408/429 may well clear up.
        return err.status >= 500 or err.status in (408, 429)
    text = str(err)
    return not any(marker in text for marker in (
        "must use https", "byte limit", "unexpected content type",
        "too many redirects", "has no host",
    ))


def parse_ics(body: bytes):
    """
    Parse the feed and expand recurrences into concrete occurrences.

    The limits applied here are the reason this is separate from the fetch:
    a feed can be tiny on the wire and still describe an effectively
    infinite number of events.
    """
    import icalendar
    import recurring_ical_events

    if b"BEGIN:VCALENDAR" not in body[:4096]:
        emit_error("fetch_error", "Response is not an iCalendar feed")

    vevents = body.count(b"BEGIN:VEVENT")
    if vevents > MAX_VEVENTS:
        emit_error(
            "fetch_error",
            f"Calendar has {vevents} events, over the {MAX_VEVENTS} limit",
        )

    _screen_recurrence_rules(body)

    now = datetime.now(timezone.utc)
    candidates = []

    try:
        with time_limit(ICS_PARSE_SECONDS, "Calendar parsing"):
            calendar = icalendar.Calendar.from_ical(body)
            query = recurring_ical_events.of(calendar)
            # Expand a closed window rather than walking the open-ended
            # after() generator. after() has to grind forward through every
            # series in the feed to guarantee global start ordering, which on
            # a real calendar costs minutes; a bounded between() answers the
            # same question in well under a second. It is also the cap the
            # security review asked for: the work is a function of the window
            # we chose, not of how far a hostile RRULE reaches.
            for days in LOOKAHEAD_STEPS:
                candidates = _expand_window(query, now, now + timedelta(days=days))
                # The first window that turns anything up wins: the nearest
                # meeting cannot be further out than the window that found it.
                if candidates:
                    break
    except ParseTimeout as e:
        emit_error("fetch_error", str(e))
    except ValueError as e:
        emit_error("fetch_error", f"Bad ICS data: {e}")

    return candidates, now


def _screen_recurrence_rules(body: bytes):
    """
    Reject recurrence rules that expand faster than we can ever use.

    This runs on the raw bytes, before any parsing, because the expansion
    cost is what we are trying to avoid - once the engine is handed a
    ``FREQ=SECONDLY`` rule it will happily allocate gigabytes inside a
    single call. A calendar of *meetings* has no legitimate use for
    sub-hourly recurrence, and an unbounded hourly rule is only marginally
    less absurd, so both are refused outright rather than truncated.
    """
    for rule in _RRULE_RE.findall(body):
        rule = rule.upper()
        if b"FREQ=SECONDLY" in rule or b"FREQ=MINUTELY" in rule:
            emit_error(
                "fetch_error",
                "Calendar contains a sub-hourly recurring event, "
                "which this widget will not expand",
            )
        if (b"FREQ=HOURLY" in rule
                and b"COUNT=" not in rule and b"UNTIL=" not in rule):
            emit_error(
                "fetch_error",
                "Calendar contains an hourly recurring event with no end date",
            )


def _expand_window(query, now, window_end):
    """
    Expand ``[now, window_end)`` in chunks, enforcing MAX_OCCURRENCES.

    between() builds its entire result list - and converts every occurrence
    into a component - before it returns, so slicing what comes back caps
    nothing at all: the allocation has already happened. Chunking bounds the
    size of any single expansion, and the running total then bounds the
    whole window. Without this, MAX_OCCURRENCES is decoration.
    """
    candidates = []
    seen = set()
    total = 0
    chunk_start = now

    while chunk_start < window_end:
        chunk_end = min(chunk_start + timedelta(days=CHUNK_DAYS), window_end)
        for ev in query.between(chunk_start, chunk_end):
            total += 1
            if total > MAX_OCCURRENCES:
                emit_error(
                    "fetch_error",
                    f"Calendar expands to more than {MAX_OCCURRENCES} "
                    "occurrences in the window checked",
                )
            record = _ics_record(ev, now)
            # An event straddling a chunk boundary is returned by both
            # chunks; the tuple is exactly what the payload is built from,
            # so deduplicating on it is sufficient.
            if record is not None and record not in seen:
                seen.add(record)
                candidates.append(record)
        chunk_start = chunk_end

    return candidates


def _ics_record(ev, now):
    """Normalise one occurrence, or None if it should be ignored."""
    if str(ev.get("STATUS", "")).upper() == "CANCELLED":
        return None
    if str(ev.get("TRANSP", "")).upper() == "TRANSPARENT":  # "Free" in Outlook
        return None

    dtstart = ev.get("DTSTART")
    if dtstart is None:
        return None
    start = dtstart.dt
    # All-day entries use date (not datetime) values - skip those.
    if not isinstance(start, datetime):
        return None
    if start.tzinfo is None:
        start = start.replace(tzinfo=timezone.utc)

    dtend = ev.get("DTEND")
    end = dtend.dt if dtend is not None else start
    if not isinstance(end, datetime):
        end = start
    elif end.tzinfo is None:
        end = end.replace(tzinfo=timezone.utc)
    if end < start:
        end = start

    if end < now:
        return None
    return start, end, clean_subject(ev.get("SUMMARY", ""))


def run_ics_backend(ics_url: str):
    body = fetch_ics_body(ics_url)
    candidates, now = parse_ics(body)
    emit(build_payload(candidates, now))


# ---------------------------------------------------------------------------
# Microsoft Graph backend (silent OAuth token only)
# ---------------------------------------------------------------------------

def get_token_silent(client_id: str, tenant_id: str):
    import msal

    try:
        cache_dir = secure_io.PrivateDir(CACHE_DIR, create=False)
    except FileNotFoundError:
        emit({"ok": False, "error": "not_authenticated"})
    except secure_io.SecureIOError as e:
        emit_error("config_error", str(e))

    with cache_dir:
        try:
            serialized = cache_dir.read_text(CACHE_NAME, secure_io.MAX_TOKEN_CACHE_BYTES)
        except secure_io.SecureIOError as e:
            emit_error("config_error", str(e))
        if serialized is None:
            emit({"ok": False, "error": "not_authenticated"})

        cache = msal.SerializableTokenCache()
        cache.deserialize(serialized)

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
            try:
                cache_dir.write_text(CACHE_NAME, cache.serialize())
            except (secure_io.SecureIOError, OSError) as e:
                # A cache we couldn't persist is a warning, not a failure:
                # the token we just acquired still works for this poll.
                print(f"warning: could not update token cache: {net.redact(e)}",
                      file=sys.stderr)

    if not result or "access_token" not in result:
        desc = str((result or {}).get("error_description") or "")
        if any(marker in desc for marker in BLOCKED_MARKERS):
            emit_error("blocked_by_org", desc)
        emit({"ok": False, "error": "not_authenticated"})
    return result["access_token"]


def run_graph_backend(token: str):
    now = datetime.now(timezone.utc)
    window_end = now + timedelta(days=WINDOW_DAYS)
    try:
        body, _ = net.fetch(
            f"{GRAPH_ROOT}/me/calendarView",
            max_bytes=MAX_GRAPH_BYTES,
            timeout=10,
            headers={
                "Authorization": "Bearer " + token,
                "Prefer": 'outlook.timezone="UTC"',
            },
            params={
                "startDateTime": now.strftime("%Y-%m-%dT%H:%M:%S"),
                "endDateTime": window_end.strftime("%Y-%m-%dT%H:%M:%S"),
                "$select": "subject,start,end,isCancelled,showAs,isAllDay",
                "$orderby": "start/dateTime",
                # Enough to cover a full day for the agenda popup, while
                # still bounding what the server may send back.
                "$top": "50",
            },
            allowed_content_types=GRAPH_CONTENT_TYPES,
            deadline=time.monotonic() + GRAPH_FETCH_SECONDS,
        )
    except net.FetchError as e:
        if e.status in (401, 403):
            emit_error("blocked_by_org", str(e))
        emit_error("graph_error", str(e))

    try:
        parsed = json.loads(body)
    except json.JSONDecodeError as e:
        emit_error("graph_error", f"Malformed Graph response: {e}")

    events = parsed.get("value") if isinstance(parsed, dict) else None
    if not isinstance(events, list):
        emit_error("graph_error", "Graph response had no event list")

    candidates = []
    for ev in events[:MAX_OCCURRENCES]:
        record = _graph_record(ev)
        if record is not None:
            candidates.append(record)

    emit(build_payload(candidates, now))


def _graph_record(ev):
    """Normalise one Graph event, or None if it should be ignored."""
    if not isinstance(ev, dict):
        return None
    if ev.get("isCancelled") or ev.get("isAllDay"):
        return None
    if ev.get("showAs") == "free":
        return None

    start = _graph_time(ev.get("start"))
    if start is None:
        return None
    end = _graph_time(ev.get("end")) or start
    if end < start:
        end = start
    return start, end, clean_subject(ev.get("subject"))


def _graph_time(slot):
    """
    Parse Graph's {"dateTime": "...", "timeZone": "UTC"} shape defensively.

    We ask for UTC via the Prefer header, but the response is remote input
    like any other, so a missing or malformed value is skipped rather than
    raising out of the poll.
    """
    if not isinstance(slot, dict):
        return None
    raw = slot.get("dateTime")
    if not isinstance(raw, str) or len(raw) > 64:
        return None
    # Graph emits seven fractional digits ("...T18:00:00.0000000"), which
    # fromisoformat only learned to tolerate in 3.11. Trim to six so this
    # keeps working on the 3.10 floor the README advertises.
    raw = re.sub(r"(\.\d{6})\d+", r"\1", raw)
    try:
        parsed = datetime.fromisoformat(raw)
    except ValueError:
        return None
    return parsed.replace(tzinfo=timezone.utc) if parsed.tzinfo is None else parsed


def main():
    config = load_config()
    ics_url = str(config.get("ics_url") or "").strip()

    if ics_url:
        run_ics_backend(ics_url)
        return

    client_id = str(config.get("client_id") or "").strip() or DEFAULT_CLIENT_ID
    tenant_id = str(config.get("tenant_id") or "").strip() or DEFAULT_TENANT_ID
    token = get_token_silent(client_id, tenant_id)
    run_graph_backend(token)


if __name__ == "__main__":
    main()
