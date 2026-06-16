// wrapper-heartbeat.js — connection-lost watchdog (outer edit.html page).
//
// Receives `ping` from the main app, replies `pong`, and shows the
// #connection-lost-modal if pings stop while the editor tab is visible.
//
// Loaded by overlay/edit.html via its OWN <script> tag — deliberately NOT part
// of the editor-load list (api.js / wrapper-*.js). Keeping the watchdog
// independent means a failure loading the heavy editor/protocol code can't also
// kill disconnect detection.
//
// Depends only on:
//   • window.HOST_ORIGIN — the baked main app origin, set by edit.html's
//     pre-mount guard. If unset (guard failed: no opener / bad origin), this
//     no-ops.
//   • #connection-lost-modal + #cl-close-btn in the DOM — edit.html loads this
//     script after that markup, so they exist when this runs.
//
// Wire format is the cross-origin postMessage protocol (v: 'edit-1'); see
// wrapper-postmessage.js + the cross-origin postMessage protocol.

(function () {
  "use strict";

  if (!window.HOST_ORIGIN) return; // pre-mount validation failed

  var modal = document.getElementById("connection-lost-modal");
  if (!modal) return; // markup missing — nothing to drive

  var lastPing = 0; // 0 = never received any ping yet
  var shown = false;
  var isSavingFailed = false;
  var droppedPings = 0; // diagnostic: messages from other origins
  // We only judge the connection while the editor tab is VISIBLE.
  // A hidden tab is (a) unreliable to measure — Chrome throttles its timers +
  // inbound message delivery to ~1/min after a few minutes, which is what
  // produced the false-positive "connection lost" — and (b) pointless to alert,
  // since the user isn't looking at it. A disconnect that happened while hidden
  // surfaces when the user returns to the tab (see the watcher). Dirty changes
  // are already autosaved on tab-hide, so nothing is lost in the meantime.
  // The SENDER matters too: the main app pings every 10s, but while the user
  // is focused on THIS (editor) tab, the main app tab is backgrounded, and
  // Chrome throttles a background tab's setInterval to ~once per minute. So in
  // normal use pings arrive ~60s apart, not 10s. TIMEOUT_MS must comfortably
  // exceed that throttled cadence or the modal flickers every cycle (a 45s
  // timeout vs ~60s throttled pings showed the modal for ~15s, then a ping
  // cleared it — repeatedly). 120s tolerates a throttled sender plus one fully
  // missed ping; a genuinely-closed main app still surfaces within ~2 min.
  var TIMEOUT_MS = 120000; // no ping for this long, while visible → lost
  var SETTLE_MS = 15000; // grace after (re)gaining visibility for a ping to land

  // ---- Diagnostic instrumentation (temporary) -------------------------
  // Exposes ping cadence + tab-visibility so a recurrence is self-explaining.
  // Filter the console with "[edit][hb]". Remove once the fix is confirmed in
  // the field (keep the single CONNECTION LOST warn if you want a breadcrumb).
  var pingCount = 0;
  var firstPingAt = 0;
  var bootAt = Date.now();
  var becameVisibleAt = Date.now(); // settle window after foregrounding
  function hbLog() {
    if (window.console)
      console.log.apply(console, ["[edit][hb]"].concat([].slice.call(arguments)));
  }
  function secs(ms) { return (ms / 1000).toFixed(1) + "s"; }
  function vis() { return document.visibilityState; }
  // --------------------------------------------------------------------

  function showConnectionLost(reason) {
    if (shown)
      return;
    // With the main app gone, the diskette can no longer reach the host. If
    // there's unsaved work, surface it as a save error (red-badge diskette,
    // still clickable) instead of leaving it looking idle/greyed. The save-ack
    // watchdog in wrapper-postmessage.js does the same for an in-flight save;
    // this also covers a doc that's merely dirty with no save attempted.
    isSavingFailed = window.__editorDirty && typeof window.skSetSaveState === "function";
    shown = true;

    if (!isSavingFailed) {
      var warning = modal.querySelector("div.cl-dialog-warning");

      warning.innerText = "All changes were successfully saved."
      warning.style.color = "#2FA0AF";
    }

    modal.style.display = "flex";

    if (isSavingFailed) {
      window.skSetSaveState("error");
    }
    if (window.console)
      console.warn(
        "[edit][hb] CONNECTION LOST:", reason,
        "| sinceLastPing=", lastPing ? secs(Date.now() - lastPing) : "never",
        "| pingsReceived=", pingCount,
        "| droppedPings=", droppedPings,
        "| visibility=", vis(),
        "| uptime=", secs(Date.now() - bootAt),
      );
  }

  // NOTE: don't poll window.opener.closed for connection-lost detection.
  // Cross-origin popups under certain COOP configs (or browsers that lazily
  // apply COOP) may report opener as null/inaccessible even when the main app
  // is alive. False positives are worse than slower detection — the
  // heartbeat-timeout below is the reliable signal.

  hbLog("heartbeat armed; HOST_ORIGIN=", window.HOST_ORIGIN,
        "timeout=", secs(TIMEOUT_MS), "(visible-only)", "visibility=", vis());

  document.addEventListener("visibilitychange", function () {
    if (document.visibilityState === "visible") {
      becameVisibleAt = Date.now();
    }
    hbLog("visibility →", vis(),
          "| sinceLastPing=", lastPing ? secs(Date.now() - lastPing) : "never");
  });

  // Reject anything not from the baked-in allowed origin. This is the security
  // boundary: even if a malicious page opens this editor, its messages are
  // silently dropped here.
  window.addEventListener("message", function (ev) {
    // Allowed-origin gate. window.matchHostOrigin (set by edit.html) handles
    // exact origins AND `*.suffix` wildcards; fall back to exact compare if it's
    // somehow absent (older edit.html).
    var originOk = window.matchHostOrigin
      ? window.matchHostOrigin(ev.origin)
      : ev.origin === window.HOST_ORIGIN;
    if (!originOk) {
      // Diagnostic: surface origin/substitution mismatches instead of failing
      // silently.
      if (ev.data && ev.data.type === "ping") {
        droppedPings++;
        if (droppedPings === 1 && window.console) {
          console.warn(
            "[edit] DROPPING ping from", ev.origin, "— expected", window.HOST_ORIGIN,
            ". The editor was built with a HOST_ORIGIN that does not match the",
            "actual main app origin. Check the __ALLOWED_HOST_ORIGIN__",
            "substitution in edit.html.",
          );
        }
      }
      return;
    }
    if (!ev.data || ev.data.v !== "edit-1") return;
    if (ev.data.type === "ping") {
      var now = Date.now();
      var gap = lastPing ? now - lastPing : 0;
      pingCount++;
      if (firstPingAt === 0) {
        firstPingAt = now;
        hbLog("first ping", secs(now - bootAt), "after boot; visibility=", vis());
      } else if (gap > 20000) {
        // Only the interesting (slow) pings are logged: a hidden tab's timers +
        // message delivery get clamped to ~1/min by Chrome throttling.
        hbLog("SLOW ping #" + pingCount, "gap=", secs(gap), "visibility=", vis());
      }
      lastPing = now;
      // Pong back to the CONCRETE sender we just validated — window.HOST_ORIGIN
      // may be a wildcard rule, which is not a valid postMessage targetOrigin.
      ev.source.postMessage({ v: "edit-1", type: "pong" }, ev.origin);
      // A ping proves the link is alive — if we'd shown the modal (e.g. a
      // throttling false-positive), take it back.
      if (shown) {
        shown = false;
        modal.style.display = "none";
        hbLog("connection restored (ping resumed) — hiding modal");
      }
    }
  });

  // Timeout watcher. Notes:
  //  • Don't fire until at least one ping has arrived — slow cold-cache bundle
  //    loads (~20 MB) can exceed the window before the main app's heartbeat
  //    even starts.
  //  • Only judge while visible, with a SETTLE_MS grace after foregrounding so a
  //    live main app's resumed ping clears a stale lastPing from the throttled
  //    hidden period.
  setInterval(function () {
    if (lastPing === 0) return; // no pings yet, no timeout
    if (document.visibilityState !== "visible") return;
    if (Date.now() - becameVisibleAt < SETTLE_MS) return;
    var idle = Date.now() - lastPing;
    if (idle > 20000 && !shown) {
      hbLog("no ping for", secs(idle), "(timeout", secs(TIMEOUT_MS) + ", visible)");
    }
    if (idle > TIMEOUT_MS) {
      showConnectionLost("heartbeat-timeout");
    }
  }, 5000);

  var closeBtn = document.getElementById("cl-close-btn");
  if (closeBtn) {
    closeBtn.onclick = function () {
      try { window.close(); } catch (e) { /* ignore */ }
    };
  }
})();
