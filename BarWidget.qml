import QtQuick
import Quickshell
import Quickshell.Io
import qs.Commons
import qs.Ui

// Bar widget: countdown to the next Outlook/Microsoft 365 calendar meeting,
// plus a click-through popup listing the rest of today's meetings.
//
// Backed by two Python scripts under bin/ (run inside a private venv so we
// don't need system packages for msal/requests):
//   - test_auth.py        interactive, one-time device-code sign-in
//   - get_next_meeting.py non-interactive poll; always emits one JSON line
//
// install.sh must have been run first - see its header. The venv is not
// optional; pythonBin below is referenced by absolute path.
//
// See bin/test_auth.py for one-time Azure app registration setup.
BarWidget {
  id: root
  moduleName: "craig.next-meeting"

  readonly property string pluginDir: Quickshell.env("HOME") + "/.config/omarchy/plugins/craig.next-meeting"
  readonly property string pythonBin: pluginDir + "/venv/bin/python"
  readonly property string pollScript: pluginDir + "/bin/get_next_meeting.py"
  readonly property string authScript: pluginDir + "/bin/test_auth.py"
  readonly property string configDir: Quickshell.env("HOME") + "/.config/omarchy/next-meeting"

  // ------------------------------------------------------------------
  // Watchdog budget
  //
  // The helper talks to the network, so "it finished" is not something we
  // get to assume. It has its own internal budget (~55s worst case: ICS
  // fetch retries plus recurrence parsing), and these numbers sit outside
  // that, in two layers:
  //
  //   backendTimeoutSeconds - enforced by coreutils `timeout`, which runs
  //     the helper in its own process group and signals the *group*. That
  //     is what reaps descendants; Process.signal() below can only reach
  //     the one pid we spawned, so a wedged grandchild would otherwise
  //     survive and keep the pipe open forever.
  //   watchdogSeconds - enforced here in QML, and deliberately later. It
  //     is the backstop for `timeout` itself failing to do its job (or not
  //     existing), and it is what puts a visible error on the bar instead
  //     of leaving "Loading…" up permanently.
  //
  // Both escalate TERM -> KILL rather than going straight to KILL, so the
  // helper gets a chance to die cleanly and flush what it has.
  // ------------------------------------------------------------------
  readonly property int backendTimeoutSeconds: 60
  readonly property int killGraceSeconds: 5
  readonly property int watchdogSeconds: backendTimeoutSeconds + killGraceSeconds + 10

  // Streaming ceilings. The helper emits one JSON line of at most a few KB;
  // these are orders of magnitude above that and exist so a helper stuck in
  // a print loop (or a Python traceback storm on stderr) can't grow the
  // shell's heap without bound. Enforced *as bytes arrive*, which is why
  // stdout is a SplitParser and not a StdioCollector - StdioCollector
  // buffers the whole stream before handing it over, so by the time we
  // could look at it, the allocation has already happened.
  readonly property int maxStdoutBytes: 256 * 1024
  readonly property int maxStderrBytes: 32 * 1024
  // Mirrors MAX_DETAIL_CHARS in bin/get_next_meeting.py.
  readonly property int maxDetailChars: 400

  // Poll cadence. The countdown is recomputed locally far more often, since
  // that's free.
  readonly property int pollIntervalMs: 1200000   // 20 minutes
  readonly property int retryIntervalMs: 60000    // after a transient error
  readonly property int tickIntervalMs: 15000
  // Clicking the widget re-polls if what we have is older than this.
  readonly property int stalenessMs: 120000

  // "loading" | "ok" | "no-meeting" | "not-authenticated" | "blocked-by-org"
  // | "fetch-error" | "timeout" | "error"
  property string status: "loading"
  property string subject: ""
  property string startIso: ""
  property string endIso: ""
  property string errorDetail: ""
  property var agenda: []
  property double lastUpdatedMs: 0
  property int nowTick: 0
  property bool popupOpen: false

  // Per-run streaming state, reset by refresh().
  property int stdoutBytes: 0
  property int stderrBytes: 0
  property bool resultSeen: false
  property bool aborting: false
  property string abortReason: ""

  function close() { popupOpen = false }

  // ------------------------------------------------------------------
  // Redaction
  //
  // The helper already redacts everything it puts in `detail`. This is the
  // same scrub repeated on the display side, because this is the last point
  // before a string becomes pixels: the published ICS URL is a bearer
  // secret, and a tooltip is a very easy thing to screenshot or screen
  // share. Doing it in both places means a future change to either side
  // can't quietly put a credential on the bar.
  //
  // Mirrors bin/net.py's redact(); keep the two in step.
  // ------------------------------------------------------------------
  function redact(text) {
    var out = String(text === undefined || text === null ? "" : text)
    // Collapse markers from an earlier pass first so this stays idempotent.
    out = out.replace(/(?:\/*<redacted>)+/g, "/<redacted>")
    out = out.replace(/(:\/\/)[^/\s'"<>@]+@/gi, "$1")
    out = out.replace(/\b(bearer)\s+[A-Za-z0-9._~+/=-]+/gi, "$1 <redacted>")
    out = out.replace(/\b(url:\s*)\/[^\s'"<>]*/gi, "$1/<redacted>")
    out = out.replace(/\b([a-z][a-z0-9+.-]*):\/\/([^\s/?#'"<>]*)([^\s'"<>]*)/gi,
      function(whole, scheme, host, rest) {
        return (!rest || rest === "/")
          ? scheme + "://" + host
          : scheme + "://" + host + "/<redacted>"
      })
    return out.replace(/(?:\/*<redacted>)+/g, "/<redacted>")
  }

  // The helper clamps `detail` too (MAX_DETAIL_CHARS). Repeated here for the
  // same reason as redact(): parse errors quote the feed, so this string is
  // partly attacker-chosen, and it lands in a word-wrapping tooltip.
  function clampDetail(text) {
    var out = String(text === undefined || text === null ? "" : text)
    return out.length <= root.maxDetailChars
      ? out
      : out.substring(0, root.maxDetailChars) + "…"
  }

  // UTF-8 length of a QML (UTF-16) string, so the ceilings above are
  // genuinely byte ceilings and not character counts that a feed full of
  // emoji or CJK subjects could quietly walk past.
  function utf8Length(text) {
    var bytes = 0
    for (var i = 0; i < text.length; i++) {
      var c = text.charCodeAt(i)
      if (c < 0x80) bytes += 1
      else if (c < 0x800) bytes += 2
      else if (c >= 0xd800 && c <= 0xdbff) { bytes += 4; i++ }  // surrogate pair
      else bytes += 3
    }
    return bytes
  }

  // ------------------------------------------------------------------
  // Polling
  // ------------------------------------------------------------------

  function refresh() {
    if (pollProc.running) return
    root.stdoutBytes = 0
    root.stderrBytes = 0
    root.resultSeen = false
    root.aborting = false
    root.abortReason = ""
    watchdogTimer.restart()
    pollProc.running = true
  }

  // Escalating shutdown of a poll that overstayed its welcome.
  //
  // TERM first so Python can unwind and close its socket; KILL only if it
  // ignores that. `timeout` normally gets there first and takes the whole
  // process group with it - this path exists for when it doesn't.
  function abortPoll(reason) {
    if (!pollProc.running || root.aborting) return
    root.aborting = true
    root.abortReason = reason
    pollProc.signal(15)   // SIGTERM
    killTimer.restart()
  }

  function stopWatchdogs() {
    watchdogTimer.stop()
    killTimer.stop()
    reapTimer.stop()
    settleTimer.stop()
  }

  function onStdoutLine(line) {
    // +1 for the newline the split marker consumed.
    root.stdoutBytes += utf8Length(line) + 1
    if (root.stdoutBytes > root.maxStdoutBytes) {
      abortPoll("overrun")
      return
    }
    if (root.resultSeen) return   // one JSON line is the contract; ignore extras

    var trimmed = String(line).trim()
    if (trimmed === "") return
    root.resultSeen = true
    root.handleResult(trimmed)
    watchdogTimer.stop()
    // We have the answer. The helper exits immediately after printing, so
    // if it's still around shortly from now something is wrong with it.
    settleTimer.restart()
  }

  function onStderrLine(line) {
    root.stderrBytes += utf8Length(line) + 1
    if (root.stderrBytes > root.maxStderrBytes) abortPoll("overrun")
    // Deliberately not stored or displayed. The helper redacts its own
    // stderr, but an unexpected Python traceback would quote local
    // variables - including the ICS URL - and this is the one stream we
    // have no control over the formatting of. Bounded and dropped.
  }

  function handleResult(raw) {
    var data
    try {
      data = JSON.parse(raw)
    } catch (e) {
      root.status = "error"
      root.errorDetail = "Bad output from get_next_meeting.py"
      root.subject = ""
      root.agenda = []
      return
    }

    root.lastUpdatedMs = Date.now()

    if (!data || !data.ok) {
      var err = data ? data.error : ""
      root.status = err === "not_authenticated" ? "not-authenticated"
        : err === "blocked_by_org" ? "blocked-by-org"
        : err === "fetch_error" ? "fetch-error"
        : err === "config_error" ? "config-error" : "error"
      root.errorDetail = root.clampDetail(
        root.redact((data && (data.detail || data.error)) || ""))
      root.subject = ""
      root.agenda = []
      // Fetch errors are usually transient (network blip, brief WAF hiccup,
      // waking from sleep before the network is back) - retry soon instead
      // of waiting the full poll interval, which would otherwise leave a
      // stale "Calendar feed error" on the bar for up to 20 minutes.
      if (root.status === "fetch-error") retryTimer.restart()
      return
    }

    root.agenda = root.sanitizeAgenda(data.agenda)

    if (!data.subject) {
      root.status = "no-meeting"
      root.subject = ""
      root.startIso = ""
      root.endIso = ""
      root.errorDetail = ""
      return
    }

    root.status = "ok"
    root.subject = String(data.subject)
    root.startIso = String(data.start)
    root.endIso = String(data.end)
    root.errorDetail = ""
  }

  // The helper is trusted-ish, but it is still a separate process whose
  // output we parse; take only the fields we understand, in the shape we
  // expect, so a malformed entry can't break the delegate at paint time.
  function sanitizeAgenda(list) {
    if (!Array.isArray(list)) return []
    var out = []
    for (var i = 0; i < list.length && out.length < 50; i++) {
      var item = list[i]
      if (!item || !item.subject || !item.start || !item.end) continue
      var start = new Date(item.start).getTime()
      var end = new Date(item.end).getTime()
      if (isNaN(start) || isNaN(end)) continue
      out.push({
        subject: String(item.subject),
        startMs: start,
        endMs: end
      })
    }
    return out
  }

  // Today's meetings that haven't ended yet. Recomputed on every tick so
  // the popup keeps shedding finished meetings between polls, instead of
  // showing a 10:00 standup as "upcoming" until the next 20-minute poll.
  readonly property var visibleAgenda: {
    void root.nowTick
    var now = Date.now()
    var out = []
    for (var i = 0; i < root.agenda.length; i++) {
      if (root.agenda[i].endMs > now) out.push(root.agenda[i])
    }
    return out
  }

  // What the bar label counts down to: the first meeting that hasn't ended,
  // falling back to the helper's "next" when that one is beyond today.
  readonly property var currentNext: {
    void root.nowTick
    if (root.status !== "ok") return null
    var live = root.visibleAgenda
    if (live.length > 0) return live[0]
    var end = new Date(root.endIso).getTime()
    if (!isNaN(end) && end > Date.now()) {
      return {
        subject: root.subject,
        startMs: new Date(root.startIso).getTime(),
        endMs: end
      }
    }
    return null
  }

  // ------------------------------------------------------------------
  // Timeline geometry
  //
  // The popup draws what's left of the day proportionally rather than as a
  // flat list: in a list a 10-minute gap and a 3-hour gap look identical,
  // which is exactly the thing you open a day view to find out. Everything
  // below derives from visibleAgenda, so it re-settles on every tick as
  // meetings end and drop off.
  // ------------------------------------------------------------------

  // Vertical scale bounds, in pixels per hour. A short day is stretched up
  // to maxPxPerHour so two back-to-back meetings aren't a smear; a long one
  // is compressed only as far as minPxPerHour and then scrolls, rather than
  // shrinking into an unreadable sliver.
  readonly property int minPxPerHour: Style.space(34)
  readonly property int maxPxPerHour: Style.space(72)
  readonly property int preferredTimelineHeight: Style.space(360)
  // A 15-minute meeting is ~9px at full scale, which cannot hold a title.
  // Short blocks are drawn taller than their true duration and may overhang
  // the next one slightly - the same compromise every calendar app makes.
  readonly property int minBlockHeight: Style.space(26)

  // Top of the timeline: the start of the hour containing the earliest of
  // "now" and the meeting currently in progress (which began before now).
  readonly property double timelineStartMs: {
    void root.nowTick
    var anchor = Date.now()
    var items = root.visibleAgenda
    if (items.length > 0 && items[0].startMs < anchor) anchor = items[0].startMs
    var hour = new Date(anchor)
    hour.setMinutes(0, 0, 0)
    return hour.getTime()
  }

  // Whole hours spanned. At least two, so a single short meeting still gets
  // a grid to sit against instead of filling the whole card.
  readonly property int timelineHours: {
    void root.nowTick
    var items = root.visibleAgenda
    var last = Date.now() + 3600000   // always show an hour past now
    for (var i = 0; i < items.length; i++) last = Math.max(last, items[i].endMs)
    var hours = Math.ceil((last - root.timelineStartMs) / 3600000)
    return Math.max(2, Math.min(24, hours))
  }

  readonly property real pxPerHour: {
    var fit = root.preferredTimelineHeight / Math.max(1, root.timelineHours)
    return Math.max(root.minPxPerHour, Math.min(root.maxPxPerHour, fit))
  }

  readonly property int timelineHeight: Math.round(root.timelineHours * root.pxPerHour)
  readonly property int timelineViewHeight: Math.min(root.timelineHeight, root.preferredTimelineHeight)

  function timelineY(ms) {
    return ((ms - root.timelineStartMs) / 3600000) * root.pxPerHour
  }

  // Meetings placed into side-by-side columns where they overlap: a run of
  // transitively overlapping meetings forms a cluster, and each one takes
  // the leftmost column that is free at its start. Without this, a
  // double-booked hour would draw two blocks on top of each other and read
  // as one meeting.
  readonly property var timelineBlocks: {
    void root.nowTick
    var items = root.visibleAgenda   // already sorted by start
    var out = []
    var i = 0
    while (i < items.length) {
      var clusterEnd = items[i].endMs
      var j = i + 1
      while (j < items.length && items[j].startMs < clusterEnd) {
        clusterEnd = Math.max(clusterEnd, items[j].endMs)
        j++
      }
      var columnEnds = []
      var cluster = []
      for (var k = i; k < j; k++) {
        var column = -1
        for (var c = 0; c < columnEnds.length; c++) {
          if (items[k].startMs >= columnEnds[c]) { column = c; break }
        }
        if (column < 0) {
          columnEnds.push(items[k].endMs)
          column = columnEnds.length - 1
        } else {
          columnEnds[column] = items[k].endMs
        }
        cluster.push({
          subject: items[k].subject,
          startMs: items[k].startMs,
          endMs: items[k].endMs,
          column: column
        })
      }
      for (var p = 0; p < cluster.length; p++) {
        cluster[p].columns = columnEnds.length
        out.push(cluster[p])
      }
      i = j
    }
    return out
  }

  // ------------------------------------------------------------------
  // Formatting
  // ------------------------------------------------------------------

  function formatDuration(ms) {
    var mins = Math.round(ms / 60000)
    if (mins < 1) return "<1m"
    var h = Math.floor(mins / 60)
    var m = mins % 60
    if (h >= 24) {
      var d = Math.floor(h / 24)
      return d + "d " + (h % 24) + "h"
    }
    return h > 0 ? (h + "h " + m + "m") : (m + "m")
  }

  function formatCountdown(startMs, endMs, nowMs) {
    if (nowMs >= startMs && nowMs < endMs) return "now"
    if (startMs - nowMs <= 0) return "now"
    return formatDuration(startMs - nowMs)
  }

  // "time left until that meeting starts", phrased for a list row.
  function formatLead(startMs, endMs, nowMs) {
    if (nowMs >= endMs) return "ended"
    if (nowMs >= startMs) return "in progress, " + formatDuration(endMs - nowMs) + " left"
    return "in " + formatDuration(startMs - nowMs)
  }

  function clockRange(startMs, endMs) {
    var start = new Date(startMs)
    var end = new Date(endMs)
    return start.toLocaleTimeString(Qt.locale(), Locale.ShortFormat) + " – " +
      end.toLocaleTimeString(Qt.locale(), Locale.ShortFormat)
  }

  function clockAt(ms) {
    return new Date(ms).toLocaleTimeString(Qt.locale(), Locale.ShortFormat)
  }

  readonly property string displayText: {
    void root.nowTick
    if (root.status === "loading") return "Loading…"
    if (root.status === "not-authenticated") return "Sign in to Outlook"
    if (root.status === "blocked-by-org") return "Outlook blocked by org"
    if (root.status === "fetch-error") return "Calendar feed error"
    if (root.status === "config-error") return "Calendar config error"
    if (root.status === "timeout") return "Calendar check timed out"
    if (root.status === "error") return "Meeting: error"
    var next = root.currentNext
    if (!next) return "No upcoming meetings"
    var countdown = formatCountdown(next.startMs, next.endMs, Date.now())
    var subj = next.subject.length > 24 ? (next.subject.substring(0, 24) + "…") : next.subject
    return countdown === "now" ? ("▶ " + subj) : (subj + " in " + countdown)
  }

  readonly property string tooltipText: {
    void root.nowTick
    if (root.status === "not-authenticated")
      return "Not signed in to Outlook yet.\nClick to sign in (device code flow)."
    if (root.status === "blocked-by-org")
      return "Your org blocked sign-in (Conditional Access / consent policy).\nClick to retry & see details in a terminal."
    if (root.status === "fetch-error")
      return "Could not fetch the published ICS calendar link.\n" + root.errorDetail
    if (root.status === "config-error")
      return "Refused to read the calendar config.\n" + root.errorDetail
    if (root.status === "timeout")
      return "The calendar check ran too long and was stopped.\nMiddle-click to try again."
    if (root.status === "error")
      return "Error checking calendar: " + root.errorDetail
    var next = root.currentNext
    if (next) {
      var count = root.visibleAgenda.length
      return next.subject + "\n" + root.clockRange(next.startMs, next.endMs) +
        "\nClick for the rest of today" + (count > 1 ? (" (" + count + " meetings)") : "")
    }
    return "No more meetings today.\nClick for details."
  }

  visible: true
  implicitWidth: label.implicitWidth + Style.spacing.controlPaddingX * 2
  implicitHeight: barSize

  Component.onCompleted: refresh()

  Process {
    id: pollProc

    // `timeout` is the outer guarantee: it puts the helper in its own
    // process group and, on expiry, signals that whole group - first TERM,
    // then KILL after the grace period. That group-wide reach is the part
    // Process.signal() cannot provide, since it only knows the one pid.
    // (coreutils ships with the base system, so this is always present.)
    command: [
      "timeout",
      "--signal=TERM",
      "--kill-after=" + root.killGraceSeconds,
      String(root.backendTimeoutSeconds),
      root.pythonBin,
      root.pollScript
    ]

    // SplitParser, not StdioCollector: we need to see bytes as they arrive
    // to enforce a ceiling on them. See maxStdoutBytes above.
    stdout: SplitParser {
      splitMarker: "\n"
      onRead: function (line) { root.onStdoutLine(line) }
    }

    stderr: SplitParser {
      splitMarker: "\n"
      onRead: function (line) { root.onStderrLine(line) }
    }

    onExited: function (exitCode) {
      root.stopWatchdogs()
      if (root.resultSeen) return

      if (root.abortReason === "overrun") {
        root.status = "error"
        root.errorDetail = "Backend produced far more output than expected and was stopped"
      } else if (root.abortReason !== "" || exitCode === 124 || exitCode === 137) {
        // 124: `timeout` expired. 137: killed (128 + SIGKILL).
        root.status = "timeout"
        root.errorDetail = "Calendar check exceeded " + root.backendTimeoutSeconds + "s"
        retryTimer.restart()
      } else {
        root.status = "error"
        root.errorDetail = "Backend script failed (exit " + exitCode + ")"
        retryTimer.restart()
      }
      root.subject = ""
      root.agenda = []
    }
  }

  // Overall deadline for one poll. Fires only if `timeout` didn't.
  Timer {
    id: watchdogTimer
    interval: root.watchdogSeconds * 1000
    repeat: false
    onTriggered: root.abortPoll("deadline")
  }

  // Escalation step two: TERM was ignored, so KILL.
  Timer {
    id: killTimer
    interval: root.killGraceSeconds * 1000
    repeat: false
    onTriggered: {
      if (pollProc.running) pollProc.signal(9)   // SIGKILL
      reapTimer.restart()
    }
  }

  // Escalation step three: even KILL didn't clear it (an uninterruptible
  // wait, most likely). Drop our handle so the next poll isn't blocked
  // forever by a process we can no longer do anything about.
  Timer {
    id: reapTimer
    interval: 2000
    repeat: false
    onTriggered: if (pollProc.running) pollProc.running = false
  }

  // We already have the JSON line but the helper hasn't exited. Nudge it
  // out rather than letting it hold a pipe open until the next poll.
  Timer {
    id: settleTimer
    interval: 3000
    repeat: false
    onTriggered: if (pollProc.running) root.abortPoll("lingering")
  }

  Timer {
    interval: root.pollIntervalMs
    running: true
    repeat: true
    onTriggered: root.refresh()
  }

  // Fast retry after a transient failure, instead of waiting up to 20
  // minutes for the timer above.
  Timer {
    id: retryTimer
    interval: root.retryIntervalMs
    repeat: false
    onTriggered: root.refresh()
  }

  // Watch the config directory (not the file directly - FileView can't
  // observe a path that doesn't exist yet) so saving/editing config.json
  // triggers an immediate refresh instead of waiting for the poll timer.
  FileView {
    path: root.configDir
    watchChanges: true
    printErrors: false
    onFileChanged: root.refresh()
  }

  Timer {
    interval: root.tickIntervalMs
    running: true
    repeat: true
    onTriggered: root.nowTick = root.nowTick + 1
  }

  Item {
    anchors.fill: parent

    Text {
      id: label
      anchors.centerIn: parent
      textFormat: Text.PlainText
      text: root.displayText
      color: root.bar ? root.bar.barForeground : Color.foreground
      font.family: root.bar ? root.bar.fontFamily : Style.font.family
      font.pixelSize: Style.font.body
      elide: Text.ElideRight
      opacity: root.status === "ok" ? 1.0 : 0.7
    }
  }

  MouseArea {
    anchors.fill: parent
    hoverEnabled: true
    cursorShape: Qt.PointingHandCursor
    acceptedButtons: Qt.LeftButton | Qt.MiddleButton | Qt.RightButton

    onClicked: function (mouse) {
      if (mouse.button !== Qt.LeftButton) {
        root.refresh()
        return
      }
      // Signing in needs a terminal, so that stays the click action while
      // there's nothing to show yet.
      if (root.status === "not-authenticated" || root.status === "blocked-by-org") {
        if (root.bar) root.bar.run(
          "omarchy-launch-floating-terminal-with-presentation \"" +
          root.pythonBin + " " + root.authScript + "\""
        )
        return
      }
      root.popupOpen = !root.popupOpen
      // Opening the day's agenda is exactly when a stale answer is most
      // annoying, so top it up if the last poll is getting old.
      if (root.popupOpen && Date.now() - root.lastUpdatedMs > root.stalenessMs) root.refresh()
    }
    onEntered: if (root.bar) root.bar.showTooltip(root, root.tooltipText)
    onExited: if (root.bar) root.bar.hideTooltip(root)
  }

  // ------------------------------------------------------------------
  // Today's agenda popup
  // ------------------------------------------------------------------

  PopupCard {
    id: popup
    anchorItem: root
    bar: root.bar
    owner: root
    open: root.popupOpen
    contentWidth: popup.fittedContentWidth(Style.space(400))
    contentHeight: popup.fittedContentHeight(column.implicitHeight)

    readonly property color fg: root.bar ? root.bar.foreground : Color.foreground
    readonly property color accent: Color.accent
    readonly property color urgent: root.bar ? root.bar.urgent : Color.urgent
    readonly property color dim: root.bar ? Qt.darker(root.bar.foreground, 1.4) : Color.foreground
    readonly property color surface: Color.popups.background
    readonly property string face: root.bar ? root.bar.fontFamily : Style.font.family

    // Width of the hour-label column. Sized from the widest label the locale
    // can produce rather than the current one, so the grid doesn't shift
    // sideways as the day crosses from "9:00 AM" to "10:00 AM".
    readonly property int gutterWidth: hourMetrics.width + Style.space(10)

    readonly property real nowY: {
      void root.nowTick
      return root.timelineY(Date.now())
    }

    Column {
      id: column
      width: parent.width
      spacing: Style.space(8)

      Item {
        width: parent.width
        height: Math.max(todayLabel.implicitHeight, dateLabel.implicitHeight)

        Text {
          id: todayLabel
          anchors.left: parent.left
          anchors.verticalCenter: parent.verticalCenter
          textFormat: Text.PlainText
          text: "Today"
          color: root.bar ? root.bar.foreground : Color.foreground
          font.family: popup.face
          font.pixelSize: Style.font.subtitle
          font.bold: true
        }

        Text {
          id: dateLabel
          anchors.right: parent.right
          anchors.verticalCenter: parent.verticalCenter
          textFormat: Text.PlainText
          text: {
            void root.nowTick
            var date = new Date().toLocaleDateString(Qt.locale(), "ddd d MMM")
            var n = root.visibleAgenda.length
            if (n === 0) return date
            return date + "  ·  " + n + (n === 1 ? " meeting left" : " meetings left")
          }
          color: popup.dim
          font.family: popup.face
          font.pixelSize: Style.font.bodySmall
        }
      }

      PanelSeparator { width: parent.width }

      // Error states get the popup too, so a click always explains itself.
      Text {
        width: parent.width
        visible: root.status !== "ok" && root.status !== "no-meeting" && root.status !== "loading"
        textFormat: Text.PlainText
        text: root.tooltipText
        color: popup.dim
        font.family: popup.face
        font.pixelSize: Style.font.bodySmall
        wrapMode: Text.WordWrap
      }

      Text {
        width: parent.width
        visible: root.status === "loading"
        textFormat: Text.PlainText
        text: "Checking your calendar…"
        color: popup.dim
        font.family: popup.face
        font.pixelSize: Style.font.bodySmall
      }

      Text {
        width: parent.width
        visible: (root.status === "ok" || root.status === "no-meeting")
                 && root.visibleAgenda.length === 0
        textFormat: Text.PlainText
        text: "No more meetings today."
        color: popup.dim
        font.family: popup.face
        font.pixelSize: Style.font.body
        wrapMode: Text.WordWrap
      }

      // The day itself. A proportional timeline: hour gridlines down the
      // left, each meeting a block whose height is its real duration, and a
      // marker for where "now" sits - so gaps, overlaps and back-to-backs
      // are visible at a glance instead of being inferred from clock times.
      Item {
        id: timeline
        width: parent.width
        visible: root.visibleAgenda.length > 0
        height: visible ? root.timelineViewHeight : 0

        // Not laid out - it exists only to measure the hour gutter.
        TextMetrics {
          id: hourMetrics
          font.family: popup.face
          font.pixelSize: Style.font.caption
          text: new Date(2000, 0, 1, 22, 0).toLocaleTimeString(Qt.locale(), Locale.ShortFormat)
        }

        Flickable {
          id: timelineFlick
          anchors.fill: parent
          contentWidth: width
          contentHeight: root.timelineHeight
          clip: true
          boundsBehavior: Flickable.StopAtBounds
          interactive: contentHeight > height

          Item {
            id: timelineBody
            width: timelineFlick.width
            height: root.timelineHeight

            // Hour grid. One extra line closes off the bottom of the last
            // hour; it gets no label, since its label would sit past the
            // end of the content.
            Repeater {
              model: root.timelineHours + 1

              delegate: Item {
                id: hourRow
                required property int index
                y: hourRow.index * root.pxPerHour
                width: timelineBody.width
                height: 1

                readonly property double hourMs: root.timelineStartMs + hourRow.index * 3600000

                Rectangle {
                  x: popup.gutterWidth
                  width: parent.width - x
                  height: Math.max(1, Style.space(1))
                  color: Qt.rgba(popup.fg.r, popup.fg.g, popup.fg.b, 0.16)
                }

                Text {
                  id: hourLabel
                  y: Style.space(2)
                  width: popup.gutterWidth - Style.space(6)
                  horizontalAlignment: Text.AlignRight
                  visible: hourRow.index < root.timelineHours
                  // Yield to the "now" label when the two would physically
                  // overlap, rather than drawing two times on top of each
                  // other. Compared as real label boxes, since the hour label
                  // hangs below its line while the now label straddles its own.
                  opacity: (hourRow.y + hourLabel.y < popup.nowY + hourMetrics.height / 2
                            && hourRow.y + hourLabel.y + hourMetrics.height > popup.nowY - hourMetrics.height / 2)
                    ? 0 : 1
                  textFormat: Text.PlainText
                  text: root.clockAt(hourRow.hourMs)
                  color: popup.dim
                  font.family: popup.face
                  font.pixelSize: Style.font.caption
                }
              }
            }

            // The "now" rule. Kept below the meeting blocks (which carry
            // z: 1) so it reads as a background gridline marking the gaps
            // you're between, rather than striking through the title of the
            // meeting you're in - that one already has its own highlight.
            Rectangle {
              x: popup.gutterWidth
              width: timelineBody.width - x
              height: Math.max(1, Style.space(1))
              y: popup.nowY - height / 2
              color: popup.urgent
            }

            Repeater {
              model: root.timelineBlocks

              delegate: Rectangle {
                id: block
                required property var modelData

                readonly property bool live: {
                  void root.nowTick
                  var now = Date.now()
                  return now >= block.modelData.startMs && now < block.modelData.endMs
                }

                readonly property real slotWidth:
                  (timelineBody.width - popup.gutterWidth) / Math.max(1, block.modelData.columns)
                readonly property real trueHeight:
                  root.timelineY(block.modelData.endMs) - root.timelineY(block.modelData.startMs)

                x: popup.gutterWidth + block.modelData.column * block.slotWidth
                width: Math.max(Style.space(40), block.slotWidth - Style.space(4))
                y: Math.max(0, root.timelineY(block.modelData.startMs))
                height: Math.max(root.minBlockHeight, block.trueHeight - Style.space(2))
                z: 1
                radius: Style.space(4)
                clip: true
                // Tinted rather than translucent: a see-through block would
                // let the hour grid and the "now" rule underneath show up as
                // lines drawn across the meeting's own title.
                color: block.live
                  ? Qt.tint(popup.surface, Qt.rgba(popup.accent.r, popup.accent.g, popup.accent.b, 0.22))
                  : Qt.tint(popup.surface, Qt.rgba(popup.fg.r, popup.fg.g, popup.fg.b, 0.09))

                // Left edge stripe, the one part that stays legible when a
                // block is squeezed down to minBlockHeight.
                Rectangle {
                  width: Style.space(3)
                  height: parent.height
                  radius: Style.space(2)
                  color: block.live ? popup.urgent : Qt.rgba(popup.fg.r, popup.fg.g, popup.fg.b, 0.45)
                }

                Column {
                  anchors.fill: parent
                  anchors.leftMargin: Style.space(9)
                  anchors.rightMargin: Style.space(6)
                  anchors.topMargin: Style.space(3)
                  anchors.bottomMargin: Style.space(3)
                  spacing: Style.space(1)

                  Text {
                    width: parent.width
                    textFormat: Text.PlainText
                    text: block.modelData.subject
                    color: popup.fg
                    font.family: popup.face
                    font.pixelSize: Style.font.bodySmall
                    font.bold: block.live
                    // A block is only as tall as its meeting is long, so the
                    // title wraps where there's room for it and elides where
                    // there isn't, instead of overflowing into the next one.
                    maximumLineCount: Math.max(1, Math.floor(
                      (block.height - Style.space(6)) / (Style.font.bodySmall * 1.35)) - 1)
                    wrapMode: Text.WordWrap
                    elide: Text.ElideRight
                  }

                  Text {
                    width: parent.width
                    visible: block.height >= root.minBlockHeight + Style.space(12)
                    textFormat: Text.PlainText
                    text: {
                      void root.nowTick
                      return root.clockRange(block.modelData.startMs, block.modelData.endMs) +
                        "  ·  " + root.formatLead(block.modelData.startMs, block.modelData.endMs, Date.now())
                    }
                    color: popup.dim
                    font.family: popup.face
                    font.pixelSize: Style.font.caption
                    elide: Text.ElideRight
                  }
                }
              }
            }

            // Where you are in the day. Only the gutter half sits above the
            // blocks - the gutter is always empty, so this can never cover a
            // meeting title the way a full-width rule would.
            Item {
              id: nowMarker
              width: popup.gutterWidth
              height: Math.max(hourMetrics.height, Style.space(8))
              y: popup.nowY - height / 2
              z: 10

              Rectangle {
                anchors.verticalCenter: parent.verticalCenter
                x: parent.width - width
                width: Style.space(6)
                height: width
                radius: width / 2
                color: popup.urgent
              }

              Text {
                anchors.verticalCenter: parent.verticalCenter
                width: popup.gutterWidth - Style.space(9)
                horizontalAlignment: Text.AlignRight
                textFormat: Text.PlainText
                text: {
                  void root.nowTick
                  return root.clockAt(Date.now())
                }
                color: popup.urgent
                font.family: popup.face
                font.pixelSize: Style.font.caption
                font.bold: true
              }
            }
          }
        }
      }

      // The timeline is height-capped so a heavy day can't grow the card off
      // the screen, which means the last visible block is often cut
      // mid-title. Say so explicitly rather than leaving a half-drawn row to
      // be interpreted as a rendering bug.
      Text {
        width: parent.width
        visible: timeline.visible && timelineFlick.contentHeight > timelineFlick.height
        horizontalAlignment: Text.AlignHCenter
        textFormat: Text.PlainText
        text: timelineFlick.atYEnd ? "⌃  scroll up for earlier" : "⌄  scroll for later"
        color: popup.dim
        font.family: popup.face
        font.pixelSize: Style.font.bodySmall
      }

      // A next meeting that falls outside today still deserves a mention,
      // otherwise an empty agenda reads as "nothing scheduled, ever".
      Text {
        width: parent.width
        visible: root.status === "ok" && root.visibleAgenda.length === 0 && root.currentNext !== null
        textFormat: Text.PlainText
        text: {
          void root.nowTick
          var next = root.currentNext
          if (!next) return ""
          return "Next: " + next.subject + "\n" +
            new Date(next.startMs).toLocaleDateString(Qt.locale(), "ddd d MMM") + ", " +
            root.clockRange(next.startMs, next.endMs) + "  ·  " +
            root.formatLead(next.startMs, next.endMs, Date.now())
        }
        color: popup.dim
        font.family: popup.face
        font.pixelSize: Style.font.bodySmall
        wrapMode: Text.WordWrap
      }

      PanelSeparator { width: parent.width }

      Item {
        width: parent.width
        height: Math.max(updatedLabel.implicitHeight, refreshButton.implicitHeight)

        Text {
          id: updatedLabel
          anchors.left: parent.left
          anchors.verticalCenter: parent.verticalCenter
          textFormat: Text.PlainText
          text: {
            void root.nowTick
            if (pollProc.running) return "Checking…"
            if (root.lastUpdatedMs <= 0) return "Not checked yet"
            return "Updated " + root.formatDuration(Date.now() - root.lastUpdatedMs) + " ago"
          }
          color: popup.dim
          font.family: popup.face
          font.pixelSize: Style.font.caption
        }

        Button {
          id: refreshButton
          anchors.right: parent.right
          anchors.verticalCenter: parent.verticalCenter
          text: "Refresh"
          iconText: "󰑐"
          foreground: root.bar ? root.bar.foreground : Color.foreground
          fontSize: Style.font.bodySmall
          enabled: !pollProc.running
          opacity: enabled ? 1.0 : 0.5
          onClicked: root.refresh()
        }
      }
    }
  }
}
