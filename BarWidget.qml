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
    contentWidth: popup.fittedContentWidth(Style.space(340))
    contentHeight: popup.fittedContentHeight(column.implicitHeight)

    readonly property color dim: root.bar ? Qt.darker(root.bar.foreground, 1.4) : Color.foreground
    readonly property string face: root.bar ? root.bar.fontFamily : Style.font.family

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

      // The list itself. Capped and clipped so a 20-meeting day scrolls
      // rather than growing a popup taller than the screen.
      Item {
        width: parent.width
        visible: root.visibleAgenda.length > 0
        height: visible ? Math.min(list.contentHeight, Style.space(340)) : 0

        ListView {
          id: list
          anchors.fill: parent
          model: root.visibleAgenda
          clip: true
          spacing: Style.space(10)
          boundsBehavior: Flickable.StopAtBounds
          interactive: contentHeight > height

          delegate: Column {
            id: entry
            required property var modelData
            width: list.width
            spacing: Style.space(2)

            readonly property bool live: {
              void root.nowTick
              var now = Date.now()
              return now >= entry.modelData.startMs && now < entry.modelData.endMs
            }

            Row {
              width: parent.width
              spacing: Style.space(6)

              Text {
                id: marker
                textFormat: Text.PlainText
                // Filled dot for the meeting you're in, hollow for the rest.
                text: entry.live ? "●" : "○"
                color: root.bar ? root.bar.foreground : Color.foreground
                font.family: popup.face
                font.pixelSize: Style.font.bodySmall
              }

              Text {
                width: parent.width - marker.implicitWidth - Style.space(6)
                textFormat: Text.PlainText
                // Full title - truncation is the bar label's job, not this
                // popup's; seeing the whole subject is the point of opening it.
                text: entry.modelData.subject
                color: root.bar ? root.bar.foreground : Color.foreground
                font.family: popup.face
                font.pixelSize: Style.font.body
                font.bold: entry.live
                wrapMode: Text.WordWrap
              }
            }

            Text {
              x: marker.implicitWidth + Style.space(6)
              width: parent.width - x
              textFormat: Text.PlainText
              text: {
                void root.nowTick
                return root.clockRange(entry.modelData.startMs, entry.modelData.endMs) + "  ·  " +
                  root.formatLead(entry.modelData.startMs, entry.modelData.endMs, Date.now())
              }
              color: popup.dim
              font.family: popup.face
              font.pixelSize: Style.font.bodySmall
              elide: Text.ElideRight
            }
          }
        }
      }

      // The list is height-capped so a heavy day can't grow the card off the
      // screen, which means the last visible row is often cut mid-sentence.
      // Say so explicitly rather than leaving a half-drawn line to be
      // interpreted as a rendering bug.
      Text {
        width: parent.width
        visible: list.contentHeight > list.height
        horizontalAlignment: Text.AlignHCenter
        textFormat: Text.PlainText
        text: list.atYEnd ? "⌃  scroll up for earlier" : "⌄  scroll for more"
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
