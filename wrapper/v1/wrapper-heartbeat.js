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

  if (!window.SK_DESKTOP_TRANSPORT && !window.HOST_ORIGIN) { // pre-mount validation failed
    return;
  }

  var cannotReconnectModal = document.getElementById("cannot-reconnect-modal");
  var reconnectingModal = document.getElementById("reconnecting-modal");
  var closedHostModal = document.getElementById("main-app-closed-modal");
  var loggedOutModal = document.getElementById("main-app-logged-out-modal");

  var cannotReconnectCloseBtn = document.getElementById("crm-close-btn");
  var reconnectingCloseBtn = document.getElementById("rm-close-btn");

  if (!reconnectingModal || !cannotReconnectModal) {
    return; // markup missing — nothing to drive
  }

  var lastPing = 0; // 0 = never received any ping yet
  var isCannotReconnectModalShown = false;
  var isReconnectingModalShow = false;
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
    if (window.console) {
      console.log.apply(console, ["[edit][hb]"].concat([].slice.call(arguments)));
    }
  }

  function secs(ms) {
    return (ms / 1000).toFixed(1) + "s";
  }

  function vis() {
    return document.visibilityState;
  }

  function handlePing() {
    hideReconnecting();

    var now = Date.now();
    var gap = lastPing ? now - lastPing : 0;

    pingCount++;

    if (firstPingAt === 0) {
      firstPingAt = now;
      hbLog('first ping', secs(now - bootAt), 'after boot; visibility=', vis());
    } else if (gap > 20000) {
      hbLog('SLOW ping #' + pingCount, 'gap=', secs(gap), 'visibility=', vis());
    }

    lastPing = now;

    if (isCannotReconnectModalShown) {
      isCannotReconnectModalShown = false;
      cannotReconnectModal.style.display = 'none';
      hbLog('connection restored (ping resumed) — hiding modal');
    }
  }

  function showConnectionLost(reason) {
    var shouldShowModal = !isCannotReconnectModalShown &&
        (!loggedOutModal || getComputedStyle(loggedOutModal).display === 'none') &&
        (!closedHostModal || getComputedStyle(closedHostModal).display === 'none');

    if (!shouldShowModal) {
      return;
    }

    hideReconnecting();

    // With the main app gone, the diskette can no longer reach the host. If
    // there's unsaved work, surface it as a save error (red-badge diskette,
    // still clickable) instead of leaving it looking idle/greyed. The save-ack
    // watchdog in wrapper-postmessage.js does the same for an in-flight save;
    // this also covers a doc that's merely dirty with no save attempted.
    isSavingFailed = window.__editorDirty && typeof window.skSetSaveState === "function";
    isCannotReconnectModalShown = true;

    if (isSavingFailed) {
      var warning = cannotReconnectModal.querySelector("div.cm-footnote");

      warning.innerText = "The latest changes made in this document could NOT be saved\nbefore the connection was interrupted and will be lost.";
      warning.style.color = "#FF274B";
    }

    if (window.SK_DESKTOP_TRANSPORT) {
      var title = cannotReconnectModal.querySelector('div.cm-title');
      var descriptions = cannotReconnectModal.querySelectorAll('div.cm-description');

      if (title) {
        title.innerText = 'Could Not Reconnect to Sharekey Main App';
      }

      if (descriptions[0]) {
        descriptions[0].innerText = 'The connection to the Main App you opened this document from\ncould not be restored.';
      }

      if (descriptions[1]) {
        descriptions[1].innerHTML = "<strong>This document needs to be reopened.</strong> Please close this<br>document, then open the Main App and reopen the document<br>from there. If you are offline, reconnect to the internet first."
      }

      if (cannotReconnectCloseBtn) {
        cannotReconnectCloseBtn.innerText = 'Close Document';
      }
    }

    cannotReconnectModal.style.display = "flex";

    if (isSavingFailed) {
      window.skSetSaveState("error");
    }

    if (window.console) {
      console.warn(
          "[edit][hb] CONNECTION LOST:", reason,
          "| sinceLastPing=", lastPing ? secs(Date.now() - lastPing) : "never",
          "| pingsReceived=", pingCount,
          "| droppedPings=", droppedPings,
          "| visibility=", vis(),
          "| uptime=", secs(Date.now() - bootAt)
      );
    }
  }

  function showReconnecting() {
    var shouldShowModal = !isReconnectingModalShow &&
        (!loggedOutModal || getComputedStyle(loggedOutModal).display === 'none') &&
        (!closedHostModal || getComputedStyle(closedHostModal).display === 'none');

    if (!shouldShowModal) {
      return;
    }

    // With the main app gone, the diskette can no longer reach the host. If
    // there's unsaved work, surface it as a save error (red-badge diskette,
    // still clickable) instead of leaving it looking idle/greyed. The save-ack
    // watchdog in wrapper-postmessage.js does the same for an in-flight save;
    // this also covers a doc that's merely dirty with no save attempted.
    isSavingFailed = window.__editorDirty && typeof window.skSetSaveState === "function";
    isReconnectingModalShow = true;

    if (isSavingFailed) {
      var warning = reconnectingModal.querySelector("div.cm-footnote");

      warning.classList.remove('cm-footnote');
      warning.innerHTML = '<span class="cm-footnote" style="color: #FF274B">The latest changes made in this document could NOT be saved<br>and will be lost if you close this tab.</span> They will be saved if the<br>connection is restored.';

      if (reconnectingCloseBtn) {
        reconnectingCloseBtn.classList.remove('cm-button--positive');
        reconnectingCloseBtn.classList.add('cm-button--negative');
      }
    }

    if (window.SK_DESKTOP_TRANSPORT) {
      var title = reconnectingModal.querySelector('div.cm-title');
      var descriptionFirstLine = reconnectingModal.querySelector('div.cm-description');

      if (title) {
        title.innerText = 'Connection to Sharekey Main App Interrupted';
      }

      if (descriptionFirstLine) {
        descriptionFirstLine.innerText = 'The connection to the Main App you opened this document from\nwas interrupted.';
      }

      if (reconnectingCloseBtn) {
        reconnectingCloseBtn.innerText = 'Close Document';
      }
    }

    reconnectingModal.style.display = "flex";

    if (isSavingFailed) {
      window.skSetSaveState("error");
    }
  }

  function hideReconnecting() {
    if (!isReconnectingModalShow) {
      return;
    }

    reconnectingModal.style.display = "none";
    isReconnectingModalShow = false;
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

  if (window.SK_DESKTOP_TRANSPORT) {
    window.addEventListener('host-ping', handlePing);
  }

  // Reject anything not from the baked-in allowed origin. This is the security
  // boundary: even if a malicious page opens this editor, its messages are
  // silently dropped here.
  window.addEventListener("message", function (ev) {
    // Allowed-origin gate. window.matchHostOrigin (set by edit.html) handles
    // exact origins AND `*.suffix` wildcards; fall back to exact compare if it's
    // somehow absent (older edit.html).
    var originOk = window.matchHostOrigin ? window.matchHostOrigin(ev.origin) : ev.origin === window.HOST_ORIGIN;

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
            "substitution in edit.html."
          );
        }
      }

      return;
    }

    if (!ev.data || ev.data.v !== "edit-1") {
      return;
    }

    if (ev.data.type === "ping") {
      handlePing();

      ev.source.postMessage(
          { v: 'edit-1', type: 'pong' },
          ev.origin
      );
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
    if (lastPing === 0) {
      return; // no pings yet, no timeout
    }

    if (document.visibilityState !== "visible") {
      return;
    }

    var idle = Date.now() - lastPing;

    if (idle > TIMEOUT_MS) {
      showConnectionLost("heartbeat-timeout");

      return;
    }

    if (Date.now() - becameVisibleAt < SETTLE_MS) {
      return;
    }

    if (idle > 60000) {
      showReconnecting();
      hbLog("no ping for", secs(idle), "(timeout", secs(TIMEOUT_MS) + ", visible)");
    }
  }, 5000);

  function closeWindow() {
    if (window.SK_DESKTOP_TRANSPORT) {
      window.postMessage(
          { __skForceCloseWindow: true },
          window.location.origin
      );

      return;
    }

    window.__editorDirty = false;
    window.close();
  }

  if (cannotReconnectCloseBtn) {
    cannotReconnectCloseBtn.onclick = closeWindow;
  }

  if (reconnectingCloseBtn) {
    reconnectingCloseBtn.onclick = closeWindow;
  }
})();
