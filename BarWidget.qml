import QtQuick
import Quickshell
import Quickshell.Io
import qs.Commons
import qs.Ui

// Bar widget: countdown to the next Outlook/Microsoft 365 calendar meeting.
//
// Backed by two Python scripts under bin/ (run inside a private venv so we
// don't need system packages for msal/requests):
//   - test_auth.py        interactive, one-time device-code sign-in
//   - get_next_meeting.py non-interactive poll; always emits one JSON line
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

  // "loading" | "ok" | "no-meeting" | "not-authenticated" | "blocked-by-org" | "fetch-error" | "error"
  property string status: "loading"
  property string subject: ""
  property string startIso: ""
  property string endIso: ""
  property string errorDetail: ""
  property int nowTick: 0

  function refresh() {
    if (!pollProc.running) pollProc.running = true
  }

  function handleResult(raw) {
    var data
    try {
      data = JSON.parse(raw)
    } catch (e) {
      root.status = "error"
      root.errorDetail = "Bad output from get_next_meeting.py"
      return
    }

    if (!data.ok) {
      root.status = data.error === "not_authenticated" ? "not-authenticated"
        : data.error === "blocked_by_org" ? "blocked-by-org"
        : data.error === "fetch_error" ? "fetch-error" : "error"
      root.errorDetail = String(data.detail || data.error || "")
      root.subject = ""
      return
    }

    if (!data.subject) {
      root.status = "no-meeting"
      root.subject = ""
      return
    }

    root.status = "ok"
    root.subject = data.subject
    root.startIso = data.start
    root.endIso = data.end
    root.errorDetail = ""
  }

  function formatCountdown(startIso, endIso, nowMs) {
    var start = new Date(startIso).getTime()
    var end = new Date(endIso).getTime()
    if (nowMs >= start && nowMs < end) return "now"
    var diffMs = start - nowMs
    if (diffMs <= 0) return "now"
    var diffMin = Math.round(diffMs / 60000)
    if (diffMin < 1) return "<1m"
    var h = Math.floor(diffMin / 60)
    var m = diffMin % 60
    return h > 0 ? (h + "h " + m + "m") : (m + "m")
  }

  readonly property string displayText: {
    void root.nowTick // force re-eval every tick
    if (root.status === "loading") return "Loading…"
    if (root.status === "not-authenticated") return "Sign in to Outlook"
    if (root.status === "blocked-by-org") return "Outlook blocked by org"
    if (root.status === "fetch-error") return "Calendar feed error"
    if (root.status === "error") return "Meeting: error"
    if (root.status === "no-meeting") return "No upcoming meetings"
    var countdown = formatCountdown(root.startIso, root.endIso, Date.now())
    var subj = root.subject.length > 24 ? (root.subject.substring(0, 24) + "…") : root.subject
    return countdown === "now" ? ("▶ " + subj) : (subj + " in " + countdown)
  }

  readonly property string tooltipText: {
    if (root.status === "not-authenticated")
      return "Not signed in to Outlook yet.\nClick to sign in (device code flow)."
    if (root.status === "blocked-by-org")
      return "Your org blocked sign-in (Conditional Access / consent policy).\nClick to retry & see details in a terminal."
    if (root.status === "fetch-error")
      return "Could not fetch the published ICS calendar link.\n" + root.errorDetail
    if (root.status === "error")
      return "Error checking calendar: " + root.errorDetail
    if (root.status === "ok") {
      var start = new Date(root.startIso)
      var end = new Date(root.endIso)
      return root.subject + "\n" + start.toLocaleString(Qt.locale(), "ddd MMM d, HH:mm") +
        " – " + end.toLocaleString(Qt.locale(), "HH:mm")
    }
    return "Click to refresh"
  }

  visible: true
  implicitWidth: label.implicitWidth + Style.spacing.controlPaddingX * 2
  implicitHeight: barSize

  Component.onCompleted: refresh()

  Process {
    id: pollProc
    command: [root.pythonBin, root.pollScript]
    stdout: StdioCollector {
      waitForEnd: true
      onStreamFinished: root.handleResult(text)
    }
    onExited: function(exitCode) {
      if (exitCode !== 0 && root.status === "loading") {
        root.status = "error"
        root.errorDetail = "Backend script failed to start (exit " + exitCode + ")"
      }
    }
  }

  // Re-poll periodically; recompute the on-screen countdown far more often
  // since it's cheap and purely local (no network call).
  Timer {
    interval: 180000
    running: true
    repeat: true
    onTriggered: root.refresh()
  }

  // Watch the config directory (not the file directly - FileView can't
  // observe a path that doesn't exist yet) so saving/editing config.json
  // triggers an immediate refresh instead of waiting up to 3 minutes for
  // the timer above.
  FileView {
    path: root.configDir
    watchChanges: true
    printErrors: false
    onFileChanged: root.refresh()
  }

  Timer {
    interval: 15000
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

    onClicked: {
      if (root.status === "not-authenticated" || root.status === "blocked-by-org") {
        if (root.bar) root.bar.run(
          "omarchy-launch-floating-terminal-with-presentation \"" +
          root.pythonBin + " " + root.authScript + "\""
        )
      } else {
        root.refresh()
      }
    }
    onEntered: if (root.bar) root.bar.showTooltip(root, root.tooltipText)
    onExited: if (root.bar) root.bar.hideTooltip(root)
  }
}
