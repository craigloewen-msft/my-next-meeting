#!/usr/bin/env python3
"""
Bounded, redaction-aware HTTP for the two calendar backends.

Both backends used to call ``requests.get(...)`` and then touch
``resp.content``/``resp.json()``. That is fine against a well-behaved
server and unbounded against anything else: ``resp.content`` buffers the
*entire* body into memory before we get a say, so a hostile or simply
broken endpoint (or a captive portal handing back a giant HTML page) can
make a background poll allocate arbitrarily much. Redirects were followed
blindly too, which means a plain-HTTP hop - and with it, a published
calendar URL sent in the clear - was one ``Location:`` header away.

``fetch`` addresses that by doing every hop itself:

  * the scheme must be HTTPS, on the first request and on every redirect;
  * redirects are followed a bounded number of times, and any ``Authorization``
    header is dropped the moment the host changes, so a redirect can't be
    used to harvest a Graph token;
  * the body is streamed and abandoned the instant it crosses a hard byte
    cap, so nothing oversized is ever fully materialised;
  * a declared ``Content-Length`` over the cap is refused before any body
    is read at all;
  * the response's content type must be one we asked for, so an error page
    doesn't get handed to the ICS parser;
  * a single wall-clock deadline covers all hops, not just each one
    individually - otherwise N redirects each just under the socket timeout
    add up to an unbounded stall.

The second job of this module is ``redact``. The ICS backend's URL *is*
the credential: anyone holding it can read the calendar. ``requests``
puts the full URL into the string form of nearly every exception it
raises, and those strings were being passed straight out to the widget as
``detail`` and rendered in a tooltip on screen. Everything user-visible
goes through ``redact`` first, which keeps the host (useful: it tells you
who is unreachable) and throws away the path, query, fragment and any
bearer credentials (not useful, and secret).
"""
import re
import time
from urllib.parse import urljoin, urlsplit

import requests

DEFAULT_MAX_REDIRECTS = 3
_CHUNK = 64 * 1024


class FetchError(Exception):
    """
    A request failed. The message is already redacted and safe to display.

    ``status`` carries the HTTP status code when the failure was an error
    response, so callers can distinguish "your org revoked this" (401/403)
    from a generic transport problem without re-parsing the message.
    """

    def __init__(self, message, status=None):
        super().__init__(message)
        self.status = status


# --------------------------------------------------------------------------
# Redaction
# --------------------------------------------------------------------------

# Any scheme://host/rest-of-it. The host is kept, everything after it is not.
_URL_RE = re.compile(r"(?i)\b([a-z][a-z0-9+.\-]*)://([^\s/?#'\"<>]*)([^\s'\"<>]*)")

# "Bearer eyJ0..." / "authorization: Bearer ...", however it got stringified.
_BEARER_RE = re.compile(r"(?i)\b(bearer)\s+[A-Za-z0-9._~+/=\-]+")

# A userinfo component (https://user:pass@host) is a credential too.
_USERINFO_RE = re.compile(r"(?i)(://)[^/\s'\"<>@]+@")

# urllib3's MaxRetryError renders the *path only*, detached from its scheme
# and host: "Max retries exceeded with url: /owa/calendar/<secret>/x.ics".
# _URL_RE cannot see that as a URL, so it needs its own rule.
_BARE_PATH_RE = re.compile(r"(?i)\b(url:\s*)(/[^\s'\"<>]*)")

# Literal secrets seen by this process (see remember_secret). Pattern
# matching alone is a losing game against every library's idea of how to
# format a URL into an exception message; scrubbing the exact strings we
# know to be sensitive is the part that does not depend on guessing.
_SECRETS = set()

# Below this length a "secret" is more likely to be "/" or "/cal" and
# blanket-replacing it would mangle unrelated text.
_MIN_SECRET_CHARS = 8

MARKER = "<redacted>"

# Redaction has to be idempotent: a message can pass through here more than
# once (fetch redacts, then the caller redacts the detail again before
# emitting it). The character classes above deliberately exclude "<" and
# ">", so a marker left by an earlier pass sits just outside the next
# match and would otherwise accumulate - "/<redacted><redacted>...". This
# collapses any run of markers, with or without leading slashes, back to
# exactly one, which makes a second pass a no-op.
_MARKER_RUN_RE = re.compile(r"(?:/*" + re.escape(MARKER) + r")+")


def remember_secret(url):
    """
    Register a URL's non-public parts so ``redact`` can scrub them by value.

    A published Outlook calendar link carries its credential in the path,
    so the path (and any query) is what must never be shown.
    """
    try:
        parts = urlsplit(str(url))
    except ValueError:
        return
    for piece in (parts.path, parts.query, parts.fragment):
        if piece and len(piece) >= _MIN_SECRET_CHARS:
            _SECRETS.add(piece)


def redact(text):
    """
    Strip secrets out of a string that is about to be logged or displayed.

    Keeps enough to diagnose a problem (scheme, host, status codes, error
    class) and drops the parts that authenticate: URL paths and queries -
    a published Outlook calendar link is a bearer secret in its path -
    plus bearer tokens and any embedded userinfo.
    """
    if text is None:
        return ""
    out = str(text)
    for secret in _SECRETS:
        out = out.replace(secret, "/" + MARKER)
    out = _USERINFO_RE.sub(r"\1", out)
    out = _BEARER_RE.sub(r"\1 " + MARKER, out)
    out = _BARE_PATH_RE.sub(r"\1/" + MARKER, out)

    def _url(match):
        scheme, host, rest = match.group(1), match.group(2), match.group(3)
        if not rest or rest == "/":
            return f"{scheme}://{host}"
        return f"{scheme}://{host}/{MARKER}"

    out = _URL_RE.sub(_url, out)
    return _MARKER_RUN_RE.sub("/" + MARKER, out)


def redact_exception(exc):
    """Redacted, human-readable form of an exception (type included)."""
    detail = redact(exc)
    name = type(exc).__name__
    return f"{name}: {detail}" if detail else name


# --------------------------------------------------------------------------
# Bounded fetch
# --------------------------------------------------------------------------

def _require_https(url, what):
    parts = urlsplit(url)
    if parts.scheme.lower() != "https":
        raise FetchError(
            f"{what} must use https (got {parts.scheme or 'no'} scheme); "
            "refusing to send calendar credentials in the clear"
        )
    if not parts.netloc:
        raise FetchError(f"{what} has no host")
    return parts


def _remaining(deadline):
    if deadline is None:
        return None
    left = deadline - time.monotonic()
    if left <= 0:
        raise FetchError("timed out")
    return left


class _NoRedirectSession(requests.Session):
    """
    A Session that never recognises a redirect target.

    This is not cosmetic. With ``allow_redirects=False``, ``Session.send``
    still primes ``response._next`` by pulling one item out of
    ``resolve_redirects(...)``, and the first thing that generator does is
    touch ``resp.content`` - an unbounded read of the whole redirect body,
    inside ``session.get()``, before this module's byte cap can see a single
    byte. Returning None keeps ``resolve_redirects`` out of its ``while url``
    loop entirely, so nothing is buffered. We read ``Location`` ourselves.
    """

    def get_redirect_target(self, resp):
        return None


def fetch(url, *, max_bytes, timeout, headers=None, params=None,
          allowed_content_types=(), max_redirects=DEFAULT_MAX_REDIRECTS,
          deadline=None):
    """
    HTTPS GET ``url`` and return at most ``max_bytes`` of body.

    Returns (body_bytes, final_response). ``final_response`` has had its
    body consumed and connection released; only its headers/status are
    meaningful afterwards.

    Raises FetchError (message already redacted) on any transport error,
    HTTP error status, disallowed scheme, redirect overrun, oversized
    body, unexpected content type, or deadline overrun.
    """
    _require_https(url, "Calendar URL")
    remember_secret(url)
    session = _NoRedirectSession()
    # NB: do not touch session.max_redirects here. Even with
    # allow_redirects=False, requests primes response._next by pulling one
    # item from resolve_redirects(), which raises TooManyRedirects if the
    # session limit is already reached - setting it to 0 would turn every
    # redirect into that error before our own handling below ever runs.

    current = url
    hop_headers = dict(headers or {})
    hop_params = params

    try:
        for hop in range(max_redirects + 1):
            try:
                resp = session.get(
                    current,
                    headers=hop_headers,
                    params=hop_params,
                    timeout=_effective_timeout(timeout, deadline),
                    allow_redirects=False,
                    stream=True,
                )
            except requests.RequestException as e:
                raise FetchError(redact_exception(e)) from None

            if resp.is_redirect or resp.status_code in (301, 302, 303, 307, 308):
                location = resp.headers.get("Location", "")
                resp.close()
                if hop >= max_redirects:
                    raise FetchError(
                        f"too many redirects (stopped after {max_redirects})"
                    )
                if not location:
                    raise FetchError(
                        f"HTTP {resp.status_code} redirect with no Location header"
                    )
                nxt = urljoin(current, location)
                _require_https(nxt, "Redirect target")
                remember_secret(nxt)
                # Credentials are scoped to the host we chose to trust; a
                # redirect elsewhere does not inherit them.
                if urlsplit(nxt).netloc != urlsplit(current).netloc:
                    hop_headers.pop("Authorization", None)
                # Query params belong to the original request only; the
                # Location URL carries whatever the server wants.
                hop_params = None
                current = nxt
                _remaining(deadline)
                continue

            with resp:
                if resp.status_code >= 400:
                    raise FetchError(
                        f"HTTP {resp.status_code} {resp.reason or ''}".strip(),
                        status=resp.status_code,
                    )
                _check_declared_length(resp, max_bytes)
                body = _read_capped(resp, max_bytes, deadline)
                _check_content_type(resp, allowed_content_types)
            return body, resp

        raise FetchError(f"too many redirects (stopped after {max_redirects})")
    finally:
        session.close()


def _effective_timeout(timeout, deadline):
    left = _remaining(deadline)
    return timeout if left is None else min(timeout, left)


def _check_declared_length(resp, max_bytes):
    declared = resp.headers.get("Content-Length")
    if declared is None:
        return
    try:
        size = int(declared)
    except ValueError:
        return
    if size > max_bytes:
        raise FetchError(
            f"response declares {size} bytes, over the {max_bytes} byte limit"
        )


def _check_content_type(resp, allowed):
    if not allowed:
        return
    raw = resp.headers.get("Content-Type", "")
    kind = raw.split(";", 1)[0].strip().lower()
    if kind not in allowed:
        raise FetchError(
            f"unexpected content type {kind or '(none)'}; "
            f"expected one of {', '.join(sorted(allowed))}"
        )


def _read_capped(resp, max_bytes, deadline):
    """
    Stream the body, giving up as soon as it exceeds ``max_bytes``.

    ``iter_content`` is what keeps this bounded - touching ``resp.content``
    anywhere in this function would defeat the entire point.
    """
    chunks = []
    total = 0
    try:
        for chunk in resp.iter_content(chunk_size=_CHUNK):
            if not chunk:
                continue
            total += len(chunk)
            if total > max_bytes:
                raise FetchError(
                    f"response exceeded the {max_bytes} byte limit"
                )
            chunks.append(chunk)
            _remaining(deadline)
    except requests.RequestException as e:
        raise FetchError(redact_exception(e)) from None
    return b"".join(chunks)
