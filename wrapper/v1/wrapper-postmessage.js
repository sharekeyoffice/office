// wrapper-postmessage.js — cross-tab postMessage protocol between the host
// (main app, which opened this tab via window.open) and the editor.
//
// Wire format (versioned to allow evolution):
//   { v: 'edit-1', type: '...', requestId?: string, ... }
//
// Inbound from host:
//   { type: 'ping' }
//   { type: 'load',     bytes: ArrayBuffer|Uint8Array, fileName?, requestId?, userId, userName }
//                                                     // userId/userName: who is editing; recorded
//                                                     // as the author of every tracked change.
//                                                     // Optional — a fallback applies. @adr-0001
//   { type: 'save-request', requestId? }            // 10.1.2 step 3
//   { type: 'save-ack', saveId, ok: bool, reason? } // v0 collab
//   { type: 'close',    requestId? }
//   { type: 'set-mode', mode: 'view'|'edit', lockHolder?: {userId?, userName, isSelf?} }  // v0 collab (isSelf ⇒ held by current user, e.g. another tab)
//   { type: 'permissions', canEdit: bool }            // role-gated: canEdit=false hides the Edit button + editing label (sent before `load`)
//   { type: 'conflict' }                              // v0 collab: another user saved
//   { type: 'editor-name', userId, userName }         // reply to request-editor-name: resolved holder display name
//   { type: 'conflict-cleared' }                       // reserved; modal hides on its own after `reload`
//   { type: 'reload' }                                  // v0 collab: main app asks us to location.reload()
//
// Outbound to host:
//   { type: 'pong', editedSinceLastPing: bool }  // true if user interacted (key/click/paste) since last ping (L1.8 idle signal)
//   { type: 'ready',     editorType: 'word'|'cell'|'slide' }   // sent on boot
//   { type: 'progress',  stage: 'converting'|'opening'|'saving', requestId? }
//   { type: 'opened',    requestId? }                          // doc rendered
//   { type: 'dirty',     dirty: bool }                         // 10.1.2 step 3
//   { type: 'saved',     bytes: Uint8Array, fileName, saveId, requestId? }
//   { type: 'close-request' }                                  // user File→Close
//   { type: 'request-edit-mode' }                              // v0 collab: user clicked Edit
//   { type: 'request-edit-state' }                             // v0 collab: ask host for current lock/mode (sent on boot; fixes stale Edit button after a reload)
//   { type: 'request-editor-name', userId }                    // v0 collab: ask host to resolve a lock holder's display name by id (when the label is stuck at "Someone")
//   { type: 'mode-changed', mode: 'view'|'edit' }              // v0 collab: user toggled native dropdown
//   { type: 'reload-request' }                         // v0 collab: user clicked Reload in conflict modal
//   { type: 'error',     code: string, message: string, requestId? }
//
// 10.1.2 step 2 scope: load only. Save and dirty-tracking land in step 3.

(function (global) {
  'use strict';

  var POST_VERSION = 'edit-1';

  function log() {
    if (window.console) {
      console.log.apply(console, ['[wrapper-pm]'].concat([].slice.call(arguments)));
    }
  }

  function WrapperPostMessage(opts) {
    this.editor       = opts.editor;       // DocsAPI.DocEditor instance
    this.editorType   = opts.editorType;   // 'word' | 'cell' | 'slide'
    this.iframe       = null;              // resolved when iframe is ready
    this.hostWindow   = window.opener || window.parent;   // host tab/iframe
    this.hostOrigin   = '*';               // tightened on first inbound msg
    this.requestId    = null;              // current load/save in flight
    this.fontsPath    = '/fonts/';

    // Autosave state (v0-save)
    this.autosaveDebounceMs = 30000;       // 30s of edit-inactivity → autosave
    this.autosaveTimer      = null;        // setTimeout handle (null when not pending)
    this.dirty              = false;       // editor's last-reported dirty state
    this.pendingSaveId      = null;        // saveId we sent on `saved`, awaiting save-ack
    this.editedSincePending = false;       // user edited after pendingSaveId was sent
    this.saveAckTimer       = null;        // watchdog: if no save-ack lands, the host is gone → error
    this.saveAckTimeoutMs   = 25000;       // no save-ack within this → assume host unreachable, flip to 'error'
    this.everSaved          = false;       // a save has been confirmed this session → clean shows 'saved' (else 'idle')

    // Chunk-diff self-check state. lastFullBytes is the OOXML we sent on
    // the previous `saved`; on the next save we run a round-trip check via
    // ChunkDiff to measure delta size on real edits. No wire-protocol
    // change yet — we still send full bytes. The log lines tell us
    // whether the algorithm is viable for OOXML before we commit to it.
    this.lastFullBytes      = null;

    this.editedSinceLastPing   = false;   // set by iframe interaction listeners; reported + reset on each pong (L1.8 edit-lock idle signal)
    this._activityListenersAttached = false;

    this.installListener();
  }

  WrapperPostMessage.prototype.findIframe = function () {
    if (this.iframe && this.iframe.contentWindow) return this.iframe;
    this.iframe = document.querySelector('iframe[name="frameEditor"]');
    return this.iframe;
  };

  // Lazy-attach interaction listeners to the editor iframe's contentDocument.
  // Idempotent via the _activityListenersAttached guard. Any keydown/mousedown/
  // paste inside the editor sets editedSinceLastPing, which we report + reset
  // on each pong (L1.8 edit-lock idle signal).
  WrapperPostMessage.prototype.ensureActivityListeners = function () {
    if (this._activityListenersAttached) return;
    var iframe = this.findIframe();
    if (!iframe || !iframe.contentDocument) return;
    var self = this;
    var markActive = function () { self.editedSinceLastPing = true; };
    // Capture phase so we see events regardless of where focus sits inside
    // the editor. keydown = typing, mousedown = toolbar/canvas interaction,
    // paste = clipboard insert. Heuristic "user is interacting" signal.
    ['keydown', 'mousedown', 'paste'].forEach(function (evt) {
      iframe.contentDocument.addEventListener(evt, markActive, true);
    });
    this._activityListenersAttached = true;
  };

  WrapperPostMessage.prototype.toHost = function (msg, transfer) {
    if (!this.hostWindow) return;
    msg.v = POST_VERSION;
    try {
      this.hostWindow.postMessage(msg, this.hostOrigin || '*', transfer || []);
    } catch (e) {
      log('postToHost failed:', e.message);
    }
  };

  WrapperPostMessage.prototype.error = function (code, message, requestId) {
    log('error', code, message);
    this.toHost({ type: 'error', code: code, message: message, requestId: requestId });
  };

  WrapperPostMessage.prototype.installListener = function () {
    var self = this;
    // Store the handler as an instance property so destroy() can remove it.
    // Without this, mode-switching leaks listeners — a stale view-mode pm
    // still receives `load` and races the fresh edit-mode pm, with the
    // dead listener trying to call openDocument on a destroyed editor.
    this._messageHandler = function (ev) {
      // Defensive: if this instance was destroyed but the listener is still
      // somehow attached (e.g. during a race window), bail.
      if (self._destroyed) return;

      // Only handle messages explicitly tagged as our wire format.
      var d = ev.data;
      if (!d || typeof d !== 'object' || d.v !== POST_VERSION) return;

      // Allowed-origin gate. window.matchHostOrigin (set by edit.html's guard)
      // is the single source of truth — it handles an exact baked origin AND a
      // `*.suffix` wildcard. If it's somehow absent (older edit.html), fall back
      // to the prior behaviour of trusting the opener so nothing breaks.
      if (window.matchHostOrigin && !window.matchHostOrigin(ev.origin)) return;

      // Lock the host origin to the first ALLOWED origin we hear from. This is
      // the CONCRETE origin (a wildcard rule can't be a postMessage targetOrigin),
      // and every subsequent send uses it.
      if (self.hostOrigin === '*' && ev.origin) {
        self.hostOrigin = ev.origin;
        log('locked host origin =', ev.origin);
      }

      switch (d.type) {
        case 'ping':
          self.ensureActivityListeners();
          self.toHost({ type: 'pong', editedSinceLastPing: self.editedSinceLastPing });
          self.editedSinceLastPing = false;
          return;
        case 'load':
          self.onLoad(d);
          return;
        case 'save-request':
          self.onSaveRequest(d);
          return;
        case 'save-ack':
          self.onSaveAck(d);
          return;
        case 'close':
          window.close();
          return;
        case 'set-mode':
          // Mode change is authoritative — only the trusted main app
          // (origin-pinned above) can demand a view↔edit switch. Dispatched
          // to wrapper-mount.js's handleSetMode, which tears down the
          // current editor instance and reconstructs in the new mode.
          if (typeof window.handleSetMode === 'function') {
            window.handleSetMode(d.mode, d.lockHolder);
          } else {
            log('set-mode received before wrapper-mount.js bound it — ignoring');
          }
          return;
        case 'permissions':
          // Role-gated capability from the main app (canEdit ⇐ EDIT_CONTENT
          // right). Sent before `load` so the Edit button/label gate before the
          // toolbar renders. UX/defense-in-depth only — the server enforces edit
          // rights on acquireEditLock + appendDiffChunk.
          if (typeof window.handlePermissions === 'function') {
            window.handlePermissions(d);
          } else {
            log('permissions received before wrapper-mount.js bound it — ignoring');
          }
          return;
        case 'conflict':
          // Another user saved changes to this document while we have it
          // open. Show the blocking modal; user must click Reload to refresh
          // their view of the document. The Reload click posts `reload-request`
          // back to the main app, which sends us a `reload` message that
          // triggers `location.reload()` in this same tab (no popup-blocker
          // re-trip from window.open).
          window.handleConflict();

          return;
        case 'editor-name':
          if (typeof window.handleEditorName === 'function') {
            window.handleEditorName(d.userId, d.userName);
          } else {
            log('editor-name received before wrapper-mount.js bound it — ignoring');
          }
          return;
        case 'conflict-cleared':
          // Reserved: the main app could send this if it ever wants to
          // dismiss the modal without triggering a full reload (e.g. for
          // a future in-editor-state-patch flow). Today's in-place reload
          // doesn't use this path because the whole page reload tears the
          // modal down naturally. Handle defensively so the modal hides
          // if it ever arrives.
          window.handleConflictCleared();
          return;
        case 'reload':
          // Main app has fresh bytes for us. Reload the page in-place
          // (same tab, same Window object — main app's session map stays
          // valid across the reload, and the post-reload `ready` event
          // will trigger maybeSendLoad with the new bytes). Clear the
          // dirty flag first so the editor's beforeunload guard doesn't
          // re-prompt — the user already acknowledged the loss in the
          // conflict modal.
          window.__editorDirty = false;
          window.location.reload();
          return;
        default:
          log('unknown inbound type:', d.type);
      }
    };

    window.addEventListener('message', this._messageHandler);
  };

  // Tear down: remove the message listener and clear pending state. Called
  // from wrapper-mount.js before destroying the editor instance during a
  // mode switch, so the old pm doesn't race the new one for inbound `load`.
  WrapperPostMessage.prototype.destroy = function () {
    if (this._destroyed) return;
    this._destroyed = true;
    if (this._messageHandler) {
      window.removeEventListener('message', this._messageHandler);
      this._messageHandler = null;
    }
    if (this.autosaveTimer) {
      clearTimeout(this.autosaveTimer);
      this.autosaveTimer = null;
    }
    // pendingSaveId / lastFullBytes intentionally left as-is; they'll be GC'd
    // with the rest of the instance when wrapper-mount.js drops its reference.
    log('destroyed');
  };

  WrapperPostMessage.prototype.signalReady = function () {
    this.toHost({ type: 'ready', editorType: this.editorType });
  };

  // Configure x2t-bridge once. Idempotent.
  WrapperPostMessage.prototype.ensureX2T = function () {
    if (!global.X2TBridge) {
      throw new Error('X2TBridge not loaded; check /x2t-bridge.js script tag');
    }
    if (!this._x2tConfigured) {
      global.X2TBridge.configure({ wasmDir: '/x2t/', verbose: false });
      this._x2tConfigured = true;
    }
    return global.X2TBridge;
  };

  // Inside the iframe context, set the SDK's font path so font loading
  // resolves against our dev-server `/fonts/` route. Done before document
  // open so the editor doesn't 404-loop on fonts.
  WrapperPostMessage.prototype.primeFontPath = function () {
    var iframe = this.findIframe();
    if (!iframe || !iframe.contentWindow) return false;
    var w = iframe.contentWindow;
    if (w.AscCommon && w.AscCommon.g_font_loader) {
      w.AscCommon.g_font_loader.fontFilesPath = this.fontsPath;
      return true;
    }
    return false;
  };

  // Store the editing user's identity from a `load` message. @adr-0001
  // Only non-empty strings are taken, so a host that sends `userName: ''`
  // does not overwrite a good value with a blank one. A blank name would
  // reach Common.Utils.getUserInitials, which calls .split() with no guard.
  WrapperPostMessage.prototype.rememberHostUser = function (msg) {
    var id   = (msg && typeof msg.userId   === 'string') ? msg.userId.trim()   : '';
    var name = (msg && typeof msg.userName === 'string') ? msg.userName.trim() : '';
    if (!id && !name) return null;

    var user = global.__skHostUser || (global.__skHostUser = {});
    if (id)   user.userId   = id;
    if (name) user.userName = name;
    log('host user =', user.userName || '(no name)', user.userId || '(no id)');
    return user;
  };

  // Copy the identity into the iframe so editor-stubs can put it on the
  // editor's DocInfo before the document opens. @adr-0001
  //
  // The editor is constructed at boot, but the host sends who the user is
  // with its `load` message, so the identity is not known in time for
  // buildEditorConfig. This is the same "prime the iframe before
  // openDocument" step the font path and the extracted media already use.
  //
  // Same-origin access is fine: edit.html and the editor iframe share an
  // origin.
  WrapperPostMessage.prototype.primeUserIdentity = function () {
    if (!global.__skHostUser)
      return false;

    var iframe = this.findIframe();
    if (!iframe || !iframe.contentWindow)
      return false;

    iframe.contentWindow.__skUser = {
      id:   global.__skHostUser.userId || '',
      name: global.__skHostUser.userName || ''
    };

    return true;
  };

  // x2t extracts embedded media (images, primarily) from OOXML containers
  // into its virtual FS, then x2t-bridge wraps each as a blob URL. The SDK's
  // image loader looks up these media references via
  // `AscCommon.g_oDocumentUrls.getImageUrl(name)`; without registration the
  // lookup returns null and slide images render as blank/broken.
  //
  // Same-origin access is fine here — both edit.html (this window) and the
  // editor iframe live on the dev-server origin, so blob URLs created here
  // are also dereferenceable from inside the iframe.
  WrapperPostMessage.prototype.registerExtractedMedia = function () {
    if (!global.X2TBridge || typeof global.X2TBridge.getLastMedia !== 'function') return 0;
    var media = global.X2TBridge.getLastMedia() || {};
    var iframe = this.findIframe();
    if (!iframe || !iframe.contentWindow) return 0;
    var w = iframe.contentWindow;
    var urls = w.AscCommon && w.AscCommon.g_oDocumentUrls;
    if (!urls || typeof urls.addImageUrl !== 'function') return 0;
    var names = Object.keys(media);
    for (var i = 0; i < names.length; i++) {
      urls.addImageUrl(names[i], media[names[i]]);
    }
    if (names.length) log('registered ' + names.length + ' media file(s): ' + names.join(', '));
    return names.length;
  };

  WrapperPostMessage.prototype.onLoad = function (msg) {
    var self = this;
    var requestId = msg.requestId || null;
    var shouldShowOnlyOfficeWelcomeScreen = msg.shouldShowOnlyOfficeWelcomeScreen === true;

    self.requestId = requestId;
    self.shouldShowOnlyOfficeWelcomeScreen = shouldShowOnlyOfficeWelcomeScreen;

    var ab = (msg.bytes instanceof ArrayBuffer) ? msg.bytes :
             (msg.bytes && msg.bytes.buffer instanceof ArrayBuffer) ? msg.bytes.buffer :
             null;
    if (!ab) {
      return self.error('BAD_REQUEST', 'load.bytes must be ArrayBuffer or Uint8Array', requestId);
    }
    var uint8 = new Uint8Array(ab);
    var fileName = msg.fileName || 'document';

    // Who is editing. @adr-0001
    // Recorded on every tracked change this session produces. The host may
    // omit it; buildEditorConfig then applies its fallback. Kept on the
    // window so a later reconstruct of the editor picks up the same person.
    self.rememberHostUser(msg);

    // Show the filename in the browser tab. Favicon was already set
    // from ?type= when the page loaded (see edit.html).
    document.title = fileName;

    var x2t;
    try { x2t = self.ensureX2T(); }
    catch (e) { return self.error('X2T_NOT_LOADED', e.message, requestId); }

    var fmt;
    try { fmt = x2t.detectFormat(uint8); }
    catch (e) { return self.error('FORMAT_DETECT_FAILED', e.message, requestId); }

    log('load', fileName, fmt, '(' + uint8.length + ' bytes)');

    var binPromise;
    if (fmt === 'bin') {
      binPromise = Promise.resolve(uint8);
    } else if (fmt === 'ooxml') {
      self.toHost({ type: 'progress', stage: 'converting', requestId: requestId });
      binPromise = x2t.convertToBin(uint8, fileName);
    } else {
      return self.error('UNSUPPORTED_FORMAT', 'expected OOXML or Editor.bin (got: ' + fmt + ')', requestId);
    }

    self._lastFmt = fmt;
    self._lastFileName = fileName;

    // ── Load-stall diagnostics (temporary) ───────────────────────────────
    // Intermittently the doc never opens: host gets `load`, but `loadBinary
    // intercepted` never appears in the editor and the UI sticks on "loading".
    // These logs pinpoint WHERE the pipeline stalls — x2t conversion vs. the
    // editor open-handoff — and capture tab visibility (a backgrounded tab can
    // throttle the async chain). Remove once the cause is fixed.
    var _convT0 = Date.now();
    var _convSettled = false;
    var _convWatch = setTimeout(function () {
      if (!_convSettled)
        log('⚠ LOAD STALLED: x2t convertToBin still pending after 10s (fmt=' + fmt +
            ', ' + uint8.length + ' bytes, visibility=' + document.visibilityState + ')');
    }, 10000);

    binPromise.then(function (binBytes) {
      _convSettled = true; clearTimeout(_convWatch);
      log('convert resolved in ' + (Date.now() - _convT0) + 'ms → ' + binBytes.length +
          ' bin bytes; opening… (visibility=' + document.visibilityState + ')');
      // Prime the font path before invoking openDocument so the editor's
      // font loader doesn't fire 404s during render.
      self.primeFontPath();
      // Register x2t-extracted media into the iframe's image-URL map so
      // slide/word/cell image references resolve to the in-memory blob URLs
      // x2t produced. Must happen before openDocument because the SDK's
      // ImageLoader.LoadDocumentImages reads the URL map during the open
      // path; missing entries would render as blank shapes.
      self.registerExtractedMedia();
      // Put the editing user into the iframe before the document opens, so
      // editor-stubs can set it on DocInfo. Every revision created in this
      // session copies its author from there. @adr-0001
      self.primeUserIdentity();
      self.toHost({ type: 'progress', stage: 'opening', requestId: requestId });

      // Hand off to DocsAPI: editor.openDocument(buffer) sends
      // {command:'openDocumentFromBinary', data:buffer} to the iframe, where
      // editor-stubs' Common.Gateway 'opendocumentfrombinary' handler calls the
      // patched Main.loadBinary → OpenDocumentFromBin.
      //
      // CRITICAL: wait until that handler is actually registered before
      // dispatching. editor-stubs sets `window.__wrapperOpenReady` once its
      // listener + loadBinary patch are installed. Without this gate, a fast
      // x2t convert (or a throttled, backgrounded editor-bundle init) lets
      // openDocument fire first → the event is dropped → the doc never opens
      // ("stuck on loading"). Fall back to dispatching anyway after ~20s so a
      // missing flag (older editor-stubs) can't wedge the load permanently.
      var buf = binBytes.buffer.slice(binBytes.byteOffset, binBytes.byteOffset + binBytes.byteLength);
      var openTries = 0;
      (function dispatchWhenReady() {
        var ifw = (function () { var f = self.findIframe(); return f && f.contentWindow; })();
        var ready = !!(ifw && ifw.__wrapperOpenReady === true);
        if (!ready && openTries++ < 200) {
          return setTimeout(dispatchWhenReady, 100);
        }
        if (!ready) log('⚠ editor open-listener never signalled ready after 20s — dispatching anyway');
        try {
          self.editor.openDocument(new Uint8Array(buf));
          log('openDocument dispatched (openReady=' + ready + ', waited=' + (openTries * 100) +
              'ms) — awaiting "[editor-stubs] loadBinary intercepted"');
        } catch (err) {
          self.error('OPEN_FAILED', err.message, requestId);
        }
      })();
    }).catch(function (err) {
      _convSettled = true; clearTimeout(_convWatch);
      log('✗ convert/open FAILED: ' + (err && err.message || err));
      self.error('CONVERT_FAILED', err.message || String(err), requestId);
    });
  };

  // Save handler — handles both explicit `save-request` from the host AND
  // editor-side autosave (timer/visibilitychange). Both call the shared
  // `captureAndSend()` below.
  //
  // Pulls Editor.bin bytes from the iframe via the `__captureSave()` hook
  // editor-stubs.js installs (Layer 7), runs them through x2t in reverse
  // to produce OOXML bytes, then ships the result back to the host via
  // `saved`. Same x2t module instance as the load path — no extra worker
  // round-trip, no DocServer.
  //
  // The host's `save-request` may include `format: 'docx'|'xlsx'|'pptx'`
  // to override the default; otherwise we infer from the editor type.
  WrapperPostMessage.prototype.onSaveRequest = function (msg) {
    var requestId = msg.requestId || ('req-' + Date.now());
    // Explicit save supersedes any pending autosave debounce.
    if (this.autosaveTimer) {
      clearTimeout(this.autosaveTimer);
      this.autosaveTimer = null;
    }
    this.captureAndSend(requestId, msg.format);
  };

  // Shared byte-capture + send path. Used by:
  //   - onSaveRequest (explicit save from host)
  //   - triggerAutosave (timer-driven / visibilitychange)
  // `saveId` is the correlation id echoed back in `save-ack`.
  WrapperPostMessage.prototype.captureAndSend = function (saveId, formatOverride) {
    var self = this;
    var formatByEditor = { word: 'docx', cell: 'xlsx', slide: 'pptx' };
    var ext = formatOverride || formatByEditor[self.editorType] || 'docx';

    var iframe = self.findIframe();
    if (!iframe || !iframe.contentWindow) {
      self.setSaveState('error');
      return self.error('IFRAME_NOT_READY', 'editor iframe not available', saveId);
    }
    var capture = iframe.contentWindow.__captureSave;
    if (typeof capture !== 'function') {
      self.setSaveState('error');
      return self.error('CAPTURE_NOT_INSTALLED', '__captureSave missing — editor-stubs.js may not have loaded', saveId);
    }

    self.toHost({ type: 'progress', stage: 'saving', requestId: saveId });
    self.setSaveState('saving');

    var binBytes;
    try {
      binBytes = capture();
    } catch (e) {
      self.setSaveState('error');
      return self.error('SERIALIZE_FAILED', e.message, saveId);
    }

    var x2t;
    try { x2t = self.ensureX2T(); }
    catch (e) {
      self.setSaveState('error');
      return self.error('X2T_NOT_LOADED', e.message, saveId);
    }

    // Track this save so onSaveAck can correlate
    self.pendingSaveId      = saveId;
    self.editedSincePending = false;

    x2t.convertFromBin(binBytes, ext).then(function (ooxmlBytes) {
      var fileName = (self._lastFileName || 'document').replace(/\.[^.]+$/, '') + '.' + ext;
      log('saved ' + fileName + ' (' + ooxmlBytes.length + ' bytes) saveId=' + saveId);

      // ── Chunk-diff self-check ─────────────────────────────────────
      // Measure how the diff/merge round-trip behaves on real OOXML.
      // Logs only — doesn't change the bytes we send to the host.
      // Snapshot the OOXML now so it survives the postMessage transfer
      // (the transferable below detaches `ooxmlBytes.buffer`).
      var snapshot = new Uint8Array(ooxmlBytes.length);
      snapshot.set(ooxmlBytes);
      if (self.lastFullBytes && global.ChunkDiff) {
        var prevBytes = self.lastFullBytes;
        global.ChunkDiff.roundTrip(prevBytes, snapshot).then(function (r) {
          var pct = r.ratio == null ? '—' : (r.ratio * 100).toFixed(1) + '%';
          var algo = 'v' + (r.version != null ? r.version : '?');
          if (r.ok) {
            var eq = r.byteIdentical ? 'byte-identical'
                  : r.contentEquivalent ? 'content-equivalent (rezip)'
                  : 'ok';
            log('[chunk-diff] saveId=' + saveId + ' algo=' + algo +
                ' delta=' + r.deltaSize + 'B full=' + r.fullSize + 'B ratio=' + pct +
                ' (' + eq + ')');
          } else {
            log('[chunk-diff] saveId=' + saveId + ' algo=' + algo + ' FAILED roundtrip: ' +
                (r.error || ('mismatch — new=' + r.newHash + ' rec=' + r.reconstructedHash)));
          }
        });
      } else if (global.ChunkDiff) {
        log('[chunk-diff] saveId=' + saveId + ' first save — no previous snapshot to diff against');
      }
      self.lastFullBytes = snapshot;
      // ──────────────────────────────────────────────────────────────

      // Transfer the underlying buffer to avoid copying. The host gets a
      // detached Uint8Array on the inbound side.
      var ab = ooxmlBytes.buffer.slice(ooxmlBytes.byteOffset, ooxmlBytes.byteOffset + ooxmlBytes.byteLength);
      self.toHost({
        type:      'saved',
        bytes:     ab,
        fileName:  fileName,
        saveId:    saveId,
        // legacy alias so existing host code that reads `requestId` keeps working
        requestId: saveId
      }, [ab]);

      // Watchdog: the host (main app) must ack within saveAckTimeoutMs. If it
      // doesn't — e.g. the main app tab was closed mid-edit — we'd otherwise be
      // stuck in 'saving' forever (the diskette dims + disables, looking "greyed
      // out"). Flip to 'error' and clear pendingSaveId so the diskette goes
      // black + red-badge and becomes clickable again, letting the user retry.
      if (self.saveAckTimer) clearTimeout(self.saveAckTimer);
      self.saveAckTimer = setTimeout(function () {
        self.saveAckTimer = null;
        if (self.pendingSaveId !== saveId) return;   // already ack'd / superseded
        log('save-ack timeout — no confirmation from host for saveId=' + saveId + ' → error');
        self.pendingSaveId      = null;
        self.editedSincePending = false;
        self.setSaveState('error');
      }, self.saveAckTimeoutMs);
    }).catch(function (err) {
      if (self.saveAckTimer) { clearTimeout(self.saveAckTimer); self.saveAckTimer = null; }
      self.setSaveState('error');
      self.pendingSaveId = null;
      self.error('CONVERT_FROM_BIN_FAILED', err.message || String(err), saveId);
    });
  };

  // Called from wrapper-mount.js's onDocumentStateChange. Drives the
  // autosave debounce timer AND relays the dirty state to the host.
  WrapperPostMessage.prototype.onDirtyChanged = function (isDirty) {
    var self = this;
    // Relay to host (existing protocol — informational)
    self.toHost({ type: 'dirty', dirty: isDirty });

    // If a save is in flight and the user keeps editing, mark the upload
    // as "stale" so a successful save-ack doesn't wrongly clear our
    // save-state indicator.
    if (isDirty && self.pendingSaveId) {
      self.editedSincePending = true;
    }

    // Diskette → "unsaved changes" state. Only when no save is in flight;
    // mid-save we keep 'saving' (editedSincePending drives the post-ack
    // transition back to 'dirty'). This is also the path that returns the
    // button from 'saved' (green) back to 'dirty' when the user edits again.
    if (isDirty && !self.pendingSaveId) {
      self.setSaveState('dirty');
    }

    // Not modified, at rest ⇒ the document content matches the saved point.
    // This fires when the user undoes back to the last save (now that we
    // advance History.SavedIndex on save-ack) or back to the originally-opened
    // state. Mirror it on the diskette: 'saved' if we've saved this session
    // (a clean state == all changes saved), else 'idle'. dirty=false can only
    // happen at exactly the saved point — undoing past it reads dirty again —
    // so 'saved' is accurate here. Skip mid-save (pendingSaveId) — that's 'saving'.
    if (!isDirty && !self.pendingSaveId) {
      self.setSaveState(self.everSaved ? 'saved' : 'idle');
    }

    self.dirty = isDirty;

    // Restart debounce. Any new edit cancels the prior timer.
    if (self.autosaveTimer) {
      clearTimeout(self.autosaveTimer);
      self.autosaveTimer = null;
    }
    if (isDirty) {
      self.autosaveTimer = setTimeout(function () {
        self.triggerAutosave();
      }, self.autosaveDebounceMs);
    }
  };

  // Fire an autosave immediately (timer expired OR tab being backgrounded).
  // Skips if not dirty, or if a previous save is still in flight.
  WrapperPostMessage.prototype.triggerAutosave = function () {
    if (this.autosaveTimer) {
      clearTimeout(this.autosaveTimer);
      this.autosaveTimer = null;
    }
    if (!this.dirty) return;
    if (this.pendingSaveId !== null) {
      // Already saving — let the in-flight one finish. The next dirty
      // event after save-ack will restart the timer.
      log('triggerAutosave skipped — save already in flight (saveId=' + this.pendingSaveId + ')');
      return;
    }
    var saveId = 'auto-' + Date.now();
    log('triggerAutosave fires, saveId=' + saveId);
    this.captureAndSend(saveId, null);
  };

  // Explicit user-initiated save (diskette click). Runs the same capture →
  // x2t → `saved` path as autosave; captureAndSend flips the state to
  // 'saving' (disabling the button) and the eventual save-ack drives it to
  // 'saved'/'error'. Skipped if a save is already in flight. Forces a save
  // even when not dirty so the click always gives feedback.
  WrapperPostMessage.prototype.requestManualSave = function () {
    if (this.pendingSaveId !== null) {
      log('requestManualSave skipped — save already in flight (saveId=' + this.pendingSaveId + ')');
      return;
    }
    if (this.autosaveTimer) {
      clearTimeout(this.autosaveTimer);
      this.autosaveTimer = null;
    }
    var saveId = 'manual-' + Date.now();
    log('requestManualSave fires, saveId=' + saveId);
    this.captureAndSend(saveId, null);
  };

  // Host-to-editor: main app ack'd a save we sent. Update save-state UI
  // and clear pending bookkeeping.
  WrapperPostMessage.prototype.onSaveAck = function (msg) {
    var saveId = msg.saveId;
    if (!saveId || saveId !== this.pendingSaveId) {
      log('save-ack ignored (stale): got saveId=' + saveId + ', pending=' + this.pendingSaveId);
      return;
    }
    var wasEdited = this.editedSincePending;
    this.pendingSaveId      = null;        // cleared BEFORE markEditorSaved so the
    this.editedSincePending = false;       // modified event it fires is treated as at-rest
    if (this.saveAckTimer) { clearTimeout(this.saveAckTimer); this.saveAckTimer = null; }

    if (msg.ok) {
      this.everSaved = true;   // a clean state from here on means "all changes saved"
      // Advance the editor's history "saved point" to the bytes we just
      // persisted. This is what stops a mere cursor click from re-flagging the
      // doc dirty: the editor recomputes "modified" from History.Index vs
      // SavedIndex on every interface update, and until now SavedIndex was
      // never advanced. Returns whether the user edited PAST that point during
      // the upload (→ still dirty). Falls back to the pre-edit flag if the
      // iframe hook isn't available.
      var stillModified = this.markEditorSaved(wasEdited);
      if (stillModified) {
        log('save-ack ok, but edited past the saved point during upload — back to dirty');
        this.setSaveState('dirty');
      } else {
        log('save-ack ok — saved point advanced, all clean');
        this.setSaveState('saved');
      }
    } else {
      log('save-ack failed: ' + (msg.reason || '(no reason)'));
      this.setSaveState('error');
      // The next dirty event will restart the timer; or, if the doc is
      // still dirty, the timer might already be queued by a recent edit.
    }
  };

  // Advance the editor's history saved point to the last captured save via the
  // iframe hook (editor-stubs window.__skMarkSaved). Returns true if the doc is
  // still modified (edited past the saved point). Does NOT affect undo/redo —
  // only the SavedIndex marker. `fallbackEdited` preserves the old behaviour if
  // the hook isn't present (older editor-stubs).
  WrapperPostMessage.prototype.markEditorSaved = function (fallbackEdited) {
    var iframe = this.findIframe();
    var cw = iframe && iframe.contentWindow;
    if (cw && typeof cw.__skMarkSaved === 'function') {
      try { return !!cw.__skMarkSaved(); }
      catch (e) { log('markEditorSaved threw: ' + (e && e.message)); }
    }
    return !!fallbackEdited;
  };

  // Single funnel for the save lifecycle. Drives (1) our injected diskette
  // button (wrapper-mount.js owns it via window.skSetSaveState) and (2) the
  // legacy bottom-right text indicator. Because every save path — manual click,
  // autosave timer, tab-hide, edit-mode-off — runs through captureAndSend /
  // onSaveAck, the diskette always mirrors the true state.
  // States: 'idle' | 'dirty' | 'saving' | 'saved' | 'error'.
  WrapperPostMessage.prototype.setSaveState = function (state) {
    // Diskette button (no-op if wrapper-mount hasn't injected it yet).
    if (typeof window.skSetSaveState === 'function') window.skSetSaveState(state);

    // Legacy text indicator — no-op if absent (e.g. standalone test page).
    var el = document.getElementById('save-state');
    if (!el) return;
    el.dataset.state = state;
    var text;
    switch (state) {
      case 'saving': text = 'Saving…';                   break;
      case 'saved':  text = 'All changes saved';         break;
      case 'error':  text = 'Couldn’t save — retrying'; break;
      default:       text = '';                          // 'idle' / 'dirty' → no text
    }
    el.textContent = text;
  };

  global.WrapperPostMessage = WrapperPostMessage;
})(window);
