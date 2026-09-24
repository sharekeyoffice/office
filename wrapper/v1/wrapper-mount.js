// wrapper-mount.js — constructs the editor via DocsAPI.DocEditor and bridges
// the host main app to the editor via wrapper-postmessage.js.
//
// Lifecycle:
//   1. boot() runs on DOMContentLoaded.
//   2. Always constructs the editor with editorConfig.mode='edit' and
//      permissions.edit=true. Security is enforced dynamically: until the
//      main app sends `set-mode: edit` (which it does only AFTER acquiring
//      the server-side lock), we apply `asc_setRestriction(View)` to keep
//      the document read-only. URL fiddling can't grant edit privileges
//      because the main app remains the sole source of truth for mode.
//   3. After DocsAPI's `onAppReady`, we cache the iframe's internal api
//      reference (DE/SSE/PE.getController('Viewport').getApi()) and apply
//      the pending restriction.
//   4. Mode toggles after open use `asc_setRestriction(None|View)` directly —
//      NO destroy, NO reconstruct, NO host re-send of bytes. Scroll/cursor/
//      undo are preserved across the toggle.
//   5. The native "Editing/Viewing" dropdown is hidden permanently in
//      favor of our own Edit button, which we INJECT INTO the iframe header
//      after onAppReady (mountHeaderControls; approach B), since the native
//      dropdown only exists for the word editor. Our button posts
//      `request-edit-mode` / `mode-changed` to the host so the main app can
//      acquire/release its lock — uniform across word/cell/slide. The
//      `asc_onChangeRestrictions` callback stays wired as a safety net for
//      any other code path that might flip restrictions on us.
//   6. Host sends `{type:'load', bytes, fileName}` → WrapperPostMessage runs
//      x2t-bridge → Editor.bin → editor.openDocument(buffer).
//   7. On DocsAPI's `onDocumentReady`, we send `opened` to the host.
//
// State machine:
//   currentMode: 'view' | 'edit' — last mode that was applied/confirmed.
//   lockHolder:  null | { userId?, userName, isSelf? } — when set, the edit
//                lock is held; `isSelf` true ⇒ held by the current user (e.g.
//                from another tab). Drives the in-header "editing" label.
/* jshint -W106 */
/* jshint -W003 */
/* jshint -W104 */
/* jshint -W119 */
(function () {
  'use strict';

  function getQueryParam(name) {
    var match = new RegExp('[?&]' + name + '=([^&]*)').exec(window.location.search);
    return match ? decodeURIComponent(match[1]) : null;
  }

  function log() {
    if (window.console) {
      console.log.apply(console, ['[wrapper-mount]'].concat([].slice.call(arguments)));
    }
  }

  function boot() {
    var type = getQueryParam('type') || 'word';

    if (typeof window.DocsAPI === 'undefined' || typeof window.DocsAPI.DocEditor !== 'function') {
      document.body.innerHTML = '<p style="font-family:sans-serif;padding:1em;color:#a00">' +
        'DocsAPI not loaded.</p>';
      return;
    }
    if (typeof window.WrapperPostMessage !== 'function') {
      document.body.innerHTML = '<p style="font-family:sans-serif;padding:1em;color:#a00">' +
        'wrapper-postmessage.js not loaded.</p>';
      return;
    }

    // Standalone mode: when edit.html is opened directly (no window.opener
    // from a host page), there's nobody to send us a `load` postMessage.
    // Detect that and auto-load a test fixture so the user sees the editor
    // working immediately. This is the dev-friendly path; production hosts
    // will always have an opener.
    // "Standalone" = the wrapper page was visited directly with no host on
    // either side (no `window.open` opener AND no parent iframe wrapping
    // us). Embed-via-iframe hosts have `window.parent !== window` even
    // though `window.opener` is null, so checking opener alone wrongly
    // pulls a test fixture in when the host is just a few ms slow to send
    // its `load` postMessage.
    var hasOpener = !!(window.opener && window.opener !== window);
    var hasParent = !!(window.parent && window.parent !== window);
    var isStandalone = !hasOpener && !hasParent && !window.SK_DESKTOP_TRANSPORT;
    // Pick the fixture matching the editor type so the standalone smoke test
    // exercises the cascade for the right document family.
    var fixtureExt = type === 'cell' ? 'xlsx' : (type === 'slide' ? 'pptx' : 'docx');
    var fixtureUrl = '/test-fixtures/sample.' + fixtureExt;

    // ── State ───────────────────────────────────────────────────────────
    var currentMode    = 'view';   // host promotes via set-mode after acquiring lock
    var lockHolder     = null;     // { userId?, userName } when held by someone else
    var editorInstance = null;
    var pm             = null;     // WrapperPostMessage — lives for the page lifetime
    var events;                    // declared below; closed over by constructEditor
    var headerEditTooltip = null;
    var headerDownloadTooltip = null;
    var headerSaveTooltip = null;
    var headerEditBtn  = null;     // our Edit button injected into the iframe header
                                   // (approach B); null until mountHeaderControls runs
    var headerSaveBtn  = null;     // our Save (diskette) button, same approach
    var headerDownloadBtn = null;
    var headerMainAppBtn = null;   // "Main App" button in the header-right area
    var headerEditingLabel = null; // our "<who> is editing the document…" label,
                                   // injected into the tab row right of the Edit button
    var currentSaveState = 'idle'; // 'idle'|'dirty'|'saving'|'saved'|'error' — drives
                                   // the diskette icon; updated via window.skSetSaveState
    var conflictState = null; // null | { updatedBy?: string, userId?: string|null }
    var editorNameRequestedFor = null;
    // Last "X is editing…" name we were told about. The conflict label derives
    // its name from the lock holder, but a save-and-exit releases the lock at the
    // SAME moment its save lands — so the set-mode:view (which clears lockHolder)
    // can be processed just before the conflict arrives, losing the name. We
    // remember it here so handleConflict still knows who edited.
    var lastLockHolderName = null;
    var lastLockHolderId = null;
    var canDownload = false;
    var isLargeFile = false;
    var isDownloading = false;
    var editModeTransition = null; // null | 'opening' | 'exiting'
    var isDownloadStarting = false;
    var canEdit        = true;     // role-gated edit capability, set by the main app's
                                   // `permissions` message (EDIT_CONTENT right). false ⇒ the
                                   // Edit button + "editing" label are never shown and
                                   // edit-mode requests are refused. Default true (standalone
                                   // dev + backward-compat). UX/defense-in-depth only — the
                                   // server (acquireEditLock + appendDiffChunk EDIT_CONTENT
                                   // checks) is the real gate against a tampering client.
    var lastPresentationPointerDownAt = 0;
    var lastPresentationEditModalAt = 0;
    var lastBlockedEditAttemptFocus = null;
    var isDesktopClosing = false;
    var isDesktopLoggingOut = false;

    // ── Restriction-API state (hot view↔edit toggle, no destroy) ────────
    var editorApi      = null;     // iframe-internal api, cached after onAppReady
    var editorApiNs    = null;     // iframe-internal Asc namespace (for constants)
    var pendingRestrict = 'view';  // restriction to apply once api is cached; default view
    var lastAppliedRestriction = null;  // last value WE programmatically passed to
                                        // asc_setRestriction; used to suppress the
                                        // feedback loop in asc_onChangeRestrictions
    var initialRestrictionApplied = false; // tracks whether the full editing:disable
                                           // notification has been dispatched. For
                                           // all three editor types we defer the
                                           // notification to onDocumentReady because
                                           // the Toolbar controller's setApi runs in
                                           // onDocumentContentReady (AFTER appReady)
                                           // and the toolbar DOM (.toolbar / mask
                                           // attach point) isn't rendered until
                                           // app:face fires. Triggering
                                           // editing:disable from onAppReady would
                                           // run before the toolbar view has DOM —
                                           // the mask append target is empty and
                                           // the disable side-effects are lost.

    // ── Overlay / header UI ─────────────────────────────────────────────
    // Both the mode-switch affordance (Edit button) AND the "<who> is editing"
    // label now live INSIDE the iframe header (approach B — see
    // mountHeaderControls), replacing the old absolutely-positioned outer-page
    // #mode-button and #lock-held-banner overlays. updateOverlayUI drives both
    // in-header controls: the Edit button (renderEditButton) and the editing
    // label (renderEditingLabel).
    function updateOverlayUI() {
      renderEditButton();
      renderEditingLabel();
      renderDownloadButton();
    }

      // Determines the relevant DOM element for an edit attempt event.
      function getEditAttemptTarget(e) {
          if (!e || !e.target) {
              return null;
          }

          if (e.target.nodeType === 1) {
              return e.target;
          }

          return e.target.parentElement || null;
      }

      // Checks if an element is a native editable element (input, textarea, select, or contentEditable).
      function isNativeEditableElement(element) {
          if (!element) {
              return false;
          }

          var tagName = element.tagName ? element.tagName.toLowerCase() : '';

          return tagName === 'input' ||
              tagName === 'textarea' ||
              tagName === 'select' ||
              element.isContentEditable;
      }

      function handleBlockedContentCopy(e) {
          if (canEdit) {
              return;
          }

          e.preventDefault();
          e.stopPropagation();

          if (typeof e.stopImmediatePropagation === 'function') {
              e.stopImmediatePropagation();
          }

          if (e.clipboardData && typeof e.clipboardData.setData === 'function') {
              e.clipboardData.setData('text/plain', '');
          }
      }

      function handleBlockedContentCopyEvent(e) {
          if (canEdit) {
              return;
          }

          const shouldBlockContentCopy = e.type !== 'pointerdown' &&
              e.type !== 'mousedown' &&
              e.type !== 'mouseup' ||
              e.button === 2;

          if (shouldBlockContentCopy) {
              handleBlockedContentCopy(e);
          }
      }

      function handleBlockedContentCopyKeyDown(e) {
          if (canEdit) {
              return;
          }

          const key = e.key ? e.key.toLowerCase() : '';
          const isCopyShortcut = (e.metaKey || e.ctrlKey) && (key === 'c' || key === 'x');
          const isContextMenuShortcut = e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10');

          if (isCopyShortcut || isContextMenuShortcut) {
              handleBlockedContentCopy(e);
          }
      }

      function bindBlockedContentCopyListeners(targetDocument) {
          if (!targetDocument || targetDocument.__blockedContentCopyListenersBound) {
              return;
          }

          targetDocument.__blockedContentCopyListenersBound = true;

          ['copy', 'cut', 'contextmenu', 'dragstart', 'pointerdown', 'mousedown', 'mouseup', 'auxclick'].forEach(function (eventName) {
              targetDocument.addEventListener(eventName, handleBlockedContentCopyEvent, true);
          });
          targetDocument.addEventListener('keydown', handleBlockedContentCopyKeyDown, true);
      }

      function showNeedEditModeModalFromCellEditor() {
          if (currentMode === 'edit' || !canEdit) {
              return false;
          }

          showNeedEditModeModal();

          log('blocked cell editor edit attempt: user is not in edit mode');

          return true;
      }

      function bindCellEditorEditAttemptListener() {
          if (type !== 'cell') {
              return;
          }

          var iframe = document.querySelector('iframe[name="frameEditor"]');

          if (!iframe || !iframe.contentDocument) {
              return;
          }

          if (iframe.contentDocument.__cellEditorEditAttemptListenerBound) {
              return;
          }

          iframe.contentDocument.__cellEditorEditAttemptListenerBound = true;

          iframe.contentDocument.addEventListener('beforeinput', function (e) {
              const target = getEditAttemptTarget(e);
              const isCellEditorInput = !!(target && target.id === 'ce-cell-content');

              if (!isCellEditorInput || !showNeedEditModeModalFromCellEditor()) {
                  return;
              }

              e.preventDefault();
              e.stopPropagation();
          }, true);
      }

      function bindSheetTabEditAttemptListener() {
          if (type !== 'cell') {
              return;
          }

          var iframe = document.querySelector('iframe[name="frameEditor"]');

          if (!iframe || !iframe.contentDocument) {
              return;
          }

          if (iframe.contentDocument.__sheetTabEditAttemptListenerBound) {
              return;
          }

          iframe.contentDocument.__sheetTabEditAttemptListenerBound = true;

          iframe.contentDocument.addEventListener('dblclick', function (e) {
              var target = getEditAttemptTarget(e);

              if (!target || !target.closest || !target.closest('.statusbar .list-item')) {
                  return;
              }

              if (!showBlockedEditAttempt('SheetTab.rename')) {
                  return;
              }

              e.preventDefault();
              e.stopPropagation();

              if (typeof e.stopImmediatePropagation === 'function') {
                  e.stopImmediatePropagation();
              }
          }, true);
      }

      function getPresentationController(name, editorWindow) {
          try {
              if (!editorWindow || !editorWindow.PE || typeof editorWindow.PE.getController !== 'function') {
                  return null;
              }

              return editorWindow.PE.getController(name);
          } catch (e) {
              return null;
          }
      }

      function handlePresentationCanvasClick(e) {
          if (type !== 'slide' || !isPresentationCanvasElement(getEditAttemptTarget(e))) {
              return;
          }

          if (hasRecentPresentationPointerDown() && isPresentationEmptySlidePlaceholderFocused()) {
              showBlockedPresentationEditAttemptOnce('Canvas.emptySlidePlaceholder');
          }
      }

      function showBlockedEditAttempt(label) {
          if (currentMode === 'edit') {
              return false;
          }

          if (!canEdit) {
              showViewOnlyModeModal(false);
          } else if (lockHolder && !lockHolder.isSelf) {
              showViewerModeModal(false);

              log('blocked edit attempt: ' + label + ', lock held by ' + (lockHolder.userName || 'Someone'));
          } else {
              showNeedEditModeModal();

              log('blocked edit attempt: ' + label + ', user is not in edit mode');
          }

          return true;
      }

      function wrapPresentationEditAttemptMethod(object, methodName, label) {
          if (!object || typeof object[methodName] !== 'function') {
              return false;
          }

          if (object[methodName].__sharekeyBlockedEditAttemptWrapped) {
              return true;
          }

          var originalMethod = object[methodName];

          object[methodName] = function () {
              if (showBlockedEditAttempt(label + '.' + methodName)) {
                  return;
              }

              return originalMethod.apply(this, arguments);
          };

          object[methodName].__sharekeyBlockedEditAttemptWrapped = true;

          return true;
      }

      function isTurnOnEditModeModalVisible() {
          var modal = document.getElementById('turn-on-edit-mode');

          return !!(modal && modal.style.display === 'flex');
      }

      function focusTurnOnEditModeModal() {
          var modal = document.getElementById('turn-on-edit-mode');

          if (!modal) {
              return;
          }

          modal.setAttribute('tabindex', '-1');
          modal.focus();
      }

      function handleTurnOnEditModeModalKeyDown(e) {
          if (!isTurnOnEditModeModalVisible()) {
              return;
          }

          if (!e || e.key !== 'Enter') {
              return;
          }

          e.preventDefault();
          e.stopPropagation();

          if (typeof e.stopImmediatePropagation === 'function') {
              e.stopImmediatePropagation();
          }

          var editButton = document.getElementById('toem-edit-btn');

          if (editButton) {
              editButton.click();
          }
      }

      function wrapPresentationFocusObjectMethod(object) {
          if (!object || typeof object.onFocusObject !== 'function') {
              return false;
          }

          if (object.onFocusObject.__sharekeyFocusObjectEditAttemptWrapped) {
              return true;
          }

          var originalMethod = object.onFocusObject;

          object.onFocusObject = function () {
              var result = originalMethod.apply(this, arguments);

              if (hasRecentPresentationPointerDown() && isPresentationObjectOrPlaceholderFocused()) {
                  showBlockedPresentationEditAttemptOnce('Main.onFocusObject');
              }

              return result;
          };

          object.onFocusObject.__sharekeyFocusObjectEditAttemptWrapped = true;

          return true;
      }

      function bindPresentationEditAttemptMethods() {
          if (type !== 'slide') {
              return;
          }

          var iframe = document.querySelector('iframe[name="frameEditor"]');
          var editorWindow = iframe && iframe.contentWindow;
          var documentHolder = getPresentationController('DocumentHolder', editorWindow);
          var toolbar = getPresentationController('Toolbar', editorWindow);
          var main = getPresentationController('Main', editorWindow);
          var wrappedCount = 0;

          [
              'onClickPlaceholder',
              'onClickPlaceholderChart',
              'onClickPlaceholderSmartArt',
              'onClickPlaceholderTable',
              'onEditObject',
              'onNewSlide',
              'onDuplicateSlide',
              'onDeleteSlide'
          ].forEach(function (methodName) {
              if (wrapPresentationEditAttemptMethod(documentHolder, methodName, 'DocumentHolder')) {
                  wrappedCount += 1;
              }
          });

          [
              'onAddSlide',
              'onDuplicateSlide',
              'onBtnInsertTextClick',
              'onMenuInsertTextClick',
              'onInsertImageClick',
              'onInsertShape',
              'onInsertTableClick',
              'onSelectChart',
              'onInsertEquationClick',
              'onInsertSymbolClick'
          ].forEach(function (methodName) {
              if (wrapPresentationEditAttemptMethod(toolbar, methodName, 'Toolbar')) {
                  wrappedCount += 1;
              }
          });

          if (wrapPresentationFocusObjectMethod(main)) {
              wrappedCount += 1;
          }
      }

      function rememberBlockedEditAttemptFocus() {
          var iframe = document.querySelector('iframe[name="frameEditor"]');
          var iframeDocument = iframe && iframe.contentDocument;

          lastBlockedEditAttemptFocus = {
              wrapperElement: document.activeElement || null,
              iframe: iframe || null,
              iframeElement: iframeDocument ? iframeDocument.activeElement : null
          };
      }

      function restoreBlockedEditAttemptFocus() {
          var focusState = lastBlockedEditAttemptFocus;

          if (!focusState) {
              return;
          }

          setTimeout(function () {
              var iframe = focusState.iframe;
              var iframeElement = focusState.iframeElement;

              try {
                  if (iframe && iframe.contentWindow) {
                      iframe.contentWindow.focus();
                  }
              } catch (e) {}

              try {
                  if (iframeElement && typeof iframeElement.focus === 'function') {
                      iframeElement.focus();
                  }
              } catch (e) {}

              try {
                  if (editorApi && typeof editorApi.asc_enableKeyEvents === 'function') {
                      editorApi.asc_enableKeyEvents(true, true);
                  }
              } catch (e) {}

              if (!iframe && focusState.wrapperElement && typeof focusState.wrapperElement.focus === 'function') {
                  try {
                      focusState.wrapperElement.focus();
                  } catch (e) {}
              }
          }, 0);
      }

      function isPresentationCanvasElement(element) {
          return type === 'slide' &&
              element &&
              element.tagName &&
              element.tagName.toLowerCase() === 'canvas' &&
              element.id === 'id_viewer_overlay';
      }

      function markPresentationPointerDown(e) {
          const isNotPresentationCanvasClick = type !== 'slide' ||
              !e ||
              e.button !== 0 ||
              e.metaKey ||
              e.ctrlKey ||
              e.altKey ||
              !isPresentationCanvasElement(getEditAttemptTarget(e));

          if (isNotPresentationCanvasClick) {
              return;
          }

          lastPresentationPointerDownAt = Date.now();
      }

      function hasRecentPresentationPointerDown() {
          return Date.now() - lastPresentationPointerDownAt < 1000;
      }

      function hasRecentPresentationEditModal() {
          return Date.now() - lastPresentationEditModalAt < 800;
      }

      function getPresentationSelectedElementsCount() {
          if (!editorApi || typeof editorApi.getSelectedElements !== 'function') {
              return 0;
          }

          try {
              var selectedElements = editorApi.getSelectedElements();

              if (!selectedElements || typeof selectedElements.length !== 'number') {
                  return 0;
              }

              return selectedElements.length;
          } catch (e) {
              return 0;
          }
      }

      function isPresentationObjectOrPlaceholderFocused() {
          return getPresentationSelectedElementsCount() > 1;
      }

      function isPresentationEmptySlidePlaceholderFocused() {
          return getPresentationSelectedElementsCount() === 0;
      }

      function showBlockedPresentationEditAttemptOnce(methodName) {
          if (hasRecentPresentationEditModal() || !showBlockedEditAttempt(methodName)) {
              return false;
          }

          lastPresentationEditModalAt = Date.now();

          return true;
      }

    // Returns true when the user action looks like an attempt to change document
    // content while the wrapper is still in view mode. We intentionally ignore
    // common navigation / system shortcuts so simple scrolling, copying, finding
    // or selecting text doesn't show a modal.
    function isEditAttemptEvent(e) {
        if (!e) {
            return false;
        }

        var target = getEditAttemptTarget(e);
        var isOnlyOfficeDocumentInput = !!(target && target.id === 'area_id');

        if (isNativeEditableElement(target) && !isOnlyOfficeDocumentInput) {
            return false;
        }

        if (e.type === 'paste' || e.type === 'cut' || e.type === 'drop') {
            return true;
        }

        if (e.type !== 'keydown') {
            return false;
        }

        var key = e.key;

        if (!key) {
            return false;
        }

        if (e.metaKey || e.ctrlKey) {
            return key.toLowerCase() === 'b' ||
                key.toLowerCase() === 'i' ||
                key.toLowerCase() === 'u';
        }

        if (e.altKey) {
            return false;
        }

      return key.length === 1 ||
          key === 'Backspace' ||
          key === 'Delete' ||
          key === 'Enter';
    }

      function isSaveShortcut(e) {
          if (!e || e.code !== 'KeyS') {
              return false;
          }

          return e.metaKey || e.ctrlKey;
      }

      function handleSaveShortcut(e) {
          if (!isSaveShortcut(e)) {
              return;
          }

          e.preventDefault();
          e.stopPropagation();

          log('save shortcut pressed → requestManualSave');

          onSaveButtonClick();
      }

      function bindSaveShortcutListeners() {
          document.addEventListener('keydown', handleSaveShortcut, true);

          log('save shortcut listener bound on wrapper document');
      }

    // Shows the turn-on-edit-mode modal only when the user tries to edit while
    // the document is in view mode and nobody else holds the edit lock.
    //
    // If another user already holds the edit lock, we do nothing here: the header
    // already explains that someone is editing, and the Edit button is disabled.
    function handleBlockedEditAttempt(e) {
        if (
            !isEditAttemptEvent(e) ||
            (currentMode === 'edit' && !isDesktopClosing && !isDesktopLoggingOut)
        ) {
            return;
        }

        e.preventDefault();
        e.stopPropagation();

        if (typeof e.stopImmediatePropagation === 'function') {
            e.stopImmediatePropagation();
        }

        if (isDesktopClosing || isDesktopLoggingOut) {
            showDesktopClosingModal(false);

            log('blocked edit attempt: desktop closing');

            return;
        }

        if (!canEdit) {
            showViewOnlyModeModal(false);

            return;
        }

      if (lockHolder && !lockHolder.isSelf) {
        showViewerModeModal(false);

        log('blocked edit attempt: lock held by ' + (lockHolder.userName || 'Someone'));

        return;
      }

      showNeedEditModeModal();

      log('blocked edit attempt: user is not in edit mode');
    }

    // Listen for edit-like actions inside the editor iframe while the document is
    // view-restricted. The restriction API still does the real blocking; this only
    // shows the correct outer-page modal explaining why editing is unavailable.
    //
    // The listener is bound to iframe.contentDocument, but the callback itself
    // belongs to the wrapper page, so showNeedEditModeModal / showLockedEditModal
    // can safely open modals that live outside the iframe.
    function bindBlockedEditAttemptListeners() {
      var iframe = document.querySelector('iframe[name="frameEditor"]');

      if (!iframe || !iframe.contentDocument) {
          return false;
      }

      if (iframe.contentDocument.__blockedEditAttemptListenersBound) {
          return true;
      }

      iframe.contentDocument.__blockedEditAttemptListenersBound = true;
      bindBlockedContentCopyListeners(iframe.contentDocument);

      ['keydown', 'paste', 'cut', 'drop'].forEach(function (eventName) {
        iframe.contentDocument.addEventListener(eventName, handleBlockedEditAttempt, true);
      });
      bindCellEditorEditAttemptListener();
      bindSheetTabEditAttemptListener();
      iframe.contentDocument.addEventListener('pointerdown', markPresentationPointerDown, true);
      iframe.contentDocument.addEventListener('mousedown', markPresentationPointerDown, true);
      iframe.contentDocument.addEventListener('click', handlePresentationCanvasClick, true);
      bindPresentationEditAttemptMethods();

      log('blocked edit attempt listeners bound');

      iframe.contentDocument.addEventListener('keydown', handleSaveShortcut, true);
      iframe.contentDocument.addEventListener('keydown', handleTurnOnEditModeModalKeyDown, true);

      log('save shortcut listener bound on editor iframe document');

      return true;
    }

    // Shows the outer-page modal for the case where the user tries to edit while
    // the document is in view mode and nobody else holds the edit lock. The modal
    // lives outside the editor iframe, so `document` here intentionally refers to
    // the wrapper page, not iframe.contentDocument.
    function showNeedEditModeModal() {
      var modal = document.getElementById('turn-on-edit-mode');

      rememberBlockedEditAttemptFocus();

      if (!modal) {
        log('showNeedEditModeModal: modal element not found');
        return;
      }

      modal.style.display = 'flex';

      focusTurnOnEditModeModal();

      log('showNeedEditModeModal');
    }

      function showDesktopClosingModal(shouldHideFirstLine) {
          var modal = document.getElementById('cannot-start-edit-mode');
          var descriptions = modal.querySelectorAll('.cm-description');

          if (descriptions[0] && shouldHideFirstLine) {
              descriptions[0].style.display = 'none';
          } else if (descriptions[0]) {
              descriptions[0].style.display = 'block';
          }

          if (descriptions[1] && isDesktopLoggingOut) {
              descriptions[1].innerHTML = 'The <strong>Log Out of Sharekey Main App</strong> window is open.';
          } else if (descriptions[1]) {
              descriptions[1].innerHTML = 'The <strong>Close Sharekey Main App</strong> window is open.';
          }

          if (descriptions[2] && isDesktopLoggingOut) {
              descriptions[2].innerText = 'To edit this document, go to the Main App and cancel logging out.';
          } else if (descriptions[2]) {
              descriptions[2].innerText = 'To edit this document, go to the Main App and cancel closing.';
          }

          if (!modal) {
              log('showDesktopClosingModal: modal element not found');
              return;
          }

          rememberBlockedEditAttemptFocus();

          modal.style.display = 'flex';

          log('showDesktopClosingModal');
      }

      function bindDesktopClosingModal() {
          var modal = document.getElementById('cannot-start-edit-mode');
          var closeButton = document.getElementById('csem-close-btn');
          var confirmButton = document.getElementById('csem-сonfirm-btn');

          if (!modal || !closeButton || !confirmButton) {
              log('bindDesktopClosingModal: modal or close button not found');

              return;
          }

          if (modal.__desktopClosingBound) {
              return;
          }

          modal.__desktopClosingBound = true;

          closeButton.onclick = function () {
              modal.style.display = 'none';
              restoreBlockedEditAttemptFocus();
          };

          confirmButton.onclick = function () {
              modal.style.display = 'none';
              restoreBlockedEditAttemptFocus();

              if (pm) {
                  pm.toHost({ type: 'focus' });
              }
          };
      }

      function showViewOnlyModeModal(hideFirstLine) {
          var modal = document.getElementById('view-only-mode');

          if (!modal) {
              log('showViewOnlyModeModal: modal element not found');

              return;
          }

          var descriptions = modal.querySelectorAll('.cm-description');
          var illustration = modal.querySelector('.cm-icon');

          if (illustration && isLargeFile) {
              illustration.innerHTML = '<rect width="250" height="150" fill="white"/>\n' +
                  '<path d="M138.284 13.7021C139.638 13.7022 140.941 14.2071 141.939 15.1123L142.136 15.2988L174.908 48.1064C175.928 49.127 176.5 50.5107 176.5 51.9531V64.0869C176.398 64.086 176.295 64.085 176.192 64.085C163.323 64.0851 152.12 71.1935 146.284 81.6768C144.751 84.4314 144.751 87.7824 146.284 90.5371C152.12 101.02 163.323 108.129 176.192 108.129C176.295 108.129 176.398 108.127 176.5 108.126V114.107C176.5 126.171 166.873 135.987 154.882 136.291C154.692 136.296 154.501 136.298 154.31 136.298H95.6914L95.1182 136.291C94.1665 136.267 93.2298 136.182 92.3115 136.042C91.9442 135.986 91.5799 135.921 91.2188 135.847C88.8714 135.366 86.6591 134.515 84.6455 133.357C84.0258 133.001 83.4249 132.616 82.8447 132.203C82.4744 131.94 82.1137 131.664 81.7607 131.379C81.2845 130.994 80.8234 130.592 80.3809 130.17C79.4816 129.313 78.6537 128.381 77.9082 127.384C77.7502 127.173 77.5978 126.957 77.4473 126.74C76.4817 125.349 75.6693 123.842 75.0391 122.244C74.774 121.572 74.5412 120.884 74.3418 120.182C74.292 120.006 74.2438 119.83 74.1982 119.653C74.149 119.462 74.1028 119.269 74.0586 119.076C73.6931 117.479 73.5001 115.816 73.5 114.107V35.8936C73.5 34.3614 73.6555 32.8656 73.9512 31.4209C74.099 30.6985 74.2819 29.9889 74.498 29.2939C74.7141 28.5993 74.9635 27.9194 75.2441 27.2559C75.8056 25.9284 76.4925 24.6669 77.29 23.4863C78.3868 21.8629 79.6925 20.3922 81.1689 19.1133C81.4373 18.8808 81.7108 18.6541 81.9902 18.4346C82.4094 18.1051 82.8414 17.7913 83.2842 17.4922C84.1695 16.8941 85.1003 16.3581 86.0703 15.8906C86.3939 15.7347 86.7216 15.5858 87.0537 15.4453C87.3855 15.305 87.7217 15.1732 88.0615 15.0488C88.4015 14.9243 88.7452 14.8073 89.0928 14.6992C89.4401 14.5912 89.7911 14.4916 90.1455 14.4004C91.0319 14.1723 91.9397 13.9975 92.8652 13.8799C93.4206 13.8093 93.9824 13.7592 94.5498 13.7305C94.7386 13.7209 94.9281 13.7138 95.1182 13.709C95.3086 13.7042 95.4998 13.7021 95.6914 13.7021H138.284ZM166.809 85.6221C166.816 85.4782 166.828 85.3354 166.842 85.1934C166.828 85.3354 166.816 85.4783 166.809 85.6221ZM167.074 83.8359C167.081 83.81 167.085 83.7836 167.092 83.7578L167.107 83.7021C167.096 83.7465 167.085 83.7913 167.074 83.8359ZM167.483 82.5781C167.497 82.5453 167.509 82.5121 167.522 82.4795C167.509 82.5121 167.497 82.5453 167.483 82.5781ZM172.208 77.5977C172.276 77.566 172.344 77.5349 172.412 77.5049C172.344 77.535 172.276 77.566 172.208 77.5977ZM138.206 35.2197C138.206 44.4691 145.705 51.9678 154.954 51.9678H169.298L138.206 20.8438V35.2197Z" fill="#D5ECEF"/>' +
                  '<path d="M176.193 68.3369C187.458 68.3371 197.27 74.5548 202.386 83.7451C203.204 85.214 203.204 86.9999 202.386 88.4688C197.27 97.6591 187.458 103.877 176.193 103.877C164.928 103.877 155.115 97.6592 149.999 88.4688C149.181 86.9999 149.181 85.2139 149.999 83.7451C155.115 74.5546 164.928 68.3369 176.193 68.3369ZM176.189 76.7119C171.001 76.7123 166.796 80.9177 166.796 86.1055C166.796 91.293 171.002 95.4996 176.189 95.5C181.377 95.5 185.582 91.2932 185.583 86.1055C185.583 80.9175 181.377 76.7119 176.189 76.7119ZM176.189 81.7119C178.616 81.7119 180.583 83.6789 180.583 86.1055C180.582 88.533 178.614 90.5 176.189 90.5C173.764 90.4996 171.796 88.5328 171.796 86.1055C171.796 83.6791 173.763 81.7123 176.189 81.7119Z" fill="url(#paint0_linear_18074_38991)"/>\n' +
                  '<path d="M8.6 91V84.019L6.455 84.89V82.797L10.524 81.211H10.68V91H8.6ZM16.7852 83.031C15.7452 83.031 14.9262 84.175 14.9262 86.164C14.9262 88.127 15.7452 89.271 16.7852 89.271C17.8252 89.271 18.6962 88.127 18.6962 86.164C18.6962 84.175 17.8252 83.031 16.7852 83.031ZM12.8202 86.164C12.8202 83.031 14.5752 81.081 16.7852 81.081C18.9822 81.081 20.7632 83.031 20.7632 86.164C20.7632 89.271 18.9822 91.221 16.7852 91.221C14.5752 91.221 12.8202 89.271 12.8202 86.164ZM28.0031 84.24L27.1581 91H25.1431L26.5341 81.302H29.0431L31.0841 87.789H31.1101L33.1121 81.302H35.5561L36.9471 91H34.9191L34.1001 84.214H34.0481L31.9421 91H30.1351L28.0421 84.24H28.0031ZM38.4993 91V81.302H42.1263C44.3233 81.302 45.4803 82.459 45.4803 83.902C45.4803 84.981 44.8693 85.67 44.0633 86.021C45.1553 86.385 45.7403 87.217 45.7403 88.283C45.7403 89.739 44.6223 91 42.2823 91H38.4993ZM40.6053 89.206H42.1263C43.1403 89.206 43.6733 88.699 43.6733 88.01C43.6733 87.295 43.0753 86.827 42.1653 86.827H40.6053V89.206ZM40.6053 85.332H42.0483C42.9973 85.332 43.4783 84.825 43.4783 84.201C43.4783 83.564 42.9843 83.096 42.0873 83.096H40.6053V85.332Z" fill="#2FA0AF"/>' +
                  '<path d="M62.592 84.344L57.964 86.164L62.592 87.971V89.947L55.559 87.126V85.202L62.592 82.368V84.344Z" fill="#2FA0AF"/>' +
                  '<path opacity="0.55" d="M91.4275 82V72.302H93.5465V80.076H97.5635V82H91.4275ZM101.404 82.182C99.7791 82.182 98.2061 80.674 98.2061 78.425C98.2061 76.124 99.7791 74.616 101.404 74.616C102.483 74.616 103.289 75.162 103.757 75.903H103.77L103.978 74.798H105.733V82H103.978L103.77 80.895H103.757C103.289 81.649 102.483 82.182 101.404 82.182ZM100.299 78.399C100.299 79.504 101.027 80.31 102.015 80.31C102.977 80.31 103.692 79.517 103.692 78.399C103.692 77.281 102.977 76.488 102.015 76.488C101.027 76.488 100.299 77.307 100.299 78.399ZM107.457 82V74.798H109.108L109.316 75.877H109.342C109.823 75.006 110.655 74.72 111.292 74.72C111.513 74.72 111.682 74.733 111.838 74.772V76.722C111.643 76.67 111.435 76.657 111.24 76.657C110.369 76.657 109.537 77.125 109.537 78.477V82H107.457ZM115.598 74.629C116.625 74.629 117.444 75.175 117.899 75.877L118.107 74.798H119.862V81.09C119.862 83.521 118.172 84.912 115.819 84.912C114.818 84.912 113.843 84.613 113.089 84.093V82.117C113.804 82.715 114.792 83.105 115.624 83.105C116.898 83.105 117.782 82.377 117.782 81.324V80.661C117.34 81.324 116.573 81.857 115.572 81.857C113.895 81.857 112.374 80.388 112.374 78.23C112.374 76.137 113.934 74.629 115.598 74.629ZM114.467 78.243C114.467 79.244 115.182 80.024 116.144 80.024C117.106 80.024 117.808 79.231 117.808 78.23C117.808 77.242 117.106 76.462 116.144 76.462C115.182 76.462 114.467 77.268 114.467 78.243ZM125.188 82.182C122.653 82.182 121.197 80.609 121.197 78.425C121.197 76.137 122.809 74.616 124.811 74.616C126.644 74.616 128.139 75.877 128.139 78.282C128.139 78.542 128.126 78.776 128.087 78.984H123.186C123.459 80.011 124.356 80.544 125.448 80.544C126.241 80.544 127.008 80.271 127.697 79.829V81.532C126.956 81.961 126.111 82.182 125.188 82.182ZM123.147 77.736H126.163C126.15 76.8 125.565 76.254 124.798 76.254C124.018 76.254 123.342 76.826 123.147 77.736ZM98.9685 100V90.302H104.91V92.239H101.088V94.67H104.676V96.49H101.088V100H98.9685ZM106.391 100V92.798H108.471V100H106.391ZM106.287 90.679C106.287 90.055 106.781 89.561 107.431 89.561C108.055 89.561 108.562 90.055 108.562 90.679C108.562 91.316 108.055 91.81 107.431 91.81C106.781 91.81 106.287 91.316 106.287 90.679ZM110.225 100V89.561H112.305V100H110.225ZM117.647 100.182C115.112 100.182 113.656 98.609 113.656 96.425C113.656 94.137 115.268 92.616 117.27 92.616C119.103 92.616 120.598 93.877 120.598 96.282C120.598 96.542 120.585 96.776 120.546 96.984H115.645C115.918 98.011 116.815 98.544 117.907 98.544C118.7 98.544 119.467 98.271 120.156 97.829V99.532C119.415 99.961 118.57 100.182 117.647 100.182ZM115.606 95.736H118.622C118.609 94.8 118.024 94.254 117.257 94.254C116.477 94.254 115.801 94.826 115.606 95.736Z" fill="#2FA0AF"/>' +
                  '<defs>' +
                  '<linearGradient id="paint0_linear_18074_38991" x1="206.808" y1="118.241" x2="184.771" y2="61.9463" gradientUnits="userSpaceOnUse">' +
                  '<stop stop-color="#125E99"/>' +
                  '<stop offset="1" stop-color="#64E1D8"/>' +
                  '</linearGradient>' +
                  '</defs>';
          }

          if (descriptions[0]) {
              descriptions[0].style.display = hideFirstLine ? 'none' : '';
          }

          if (descriptions[1] && isLargeFile) {
              descriptions[1].innerHTML = 'Currently, this document <strong>cannot be edited</strong> (over the 10 MB limit).';
          }

          if (descriptions[2] && isLargeFile) {
              descriptions[2].innerHTML = 'Large documents will be supported soon.';
              descriptions[2].style.color = '#2FA0AF';
          }

          modal.style.display = 'flex';
      }

      // Shows the outer-page modal for the case where the user tries to edit while
      // another user already holds the edit lock. The header already shows who is
      // editing, but this modal gives immediate feedback after a typing attempt.
      function showViewerModeModal(hideFirstDescription) {
          var modal = document.getElementById('viewer-mode');

          if (!modal) {
              log('showViewerModeModal: modal element not found');

              return;
          }

          var userName = (lockHolder && lockHolder.userName) || 'Someone';
          var userNameElement = document.getElementById('vm-username');
          var firstDescription = modal.querySelector('.cm-description');

          if (userNameElement) {
              userNameElement.textContent = userName;
          }

          if (firstDescription) {
              firstDescription.style.display = hideFirstDescription ? 'none' : '';
          }

          modal.style.display = 'flex';

          log('showViewerModeModal: lock held by ' + userName);
      }

    // Binds actions inside the outer turn-on-edit-mode modal.
    //   - Edit repeats the same behaviour as the iframe-header Edit button
    //   - Close only hides the modal and does not request edit mode
    function bindTurnOnEditModeModal() {
      var modal = document.getElementById('turn-on-edit-mode');
      var editButton = document.getElementById('toem-edit-btn');
      var closeButton = document.getElementById('toem-close-btn');

      if (!modal || !editButton || !closeButton) {
        log('bindTurnOnEditModeModal: modal, edit button or close button not found');
        return;
      }

      if (modal.__turnOnEditModeBound) {
          return;
      }

      modal.__turnOnEditModeBound = true;

      document.addEventListener('keydown', handleTurnOnEditModeModalKeyDown, true);

      editButton.onclick = function () {
        modal.style.display = 'none';

        restoreBlockedEditAttemptFocus();

        log('user clicked Edit in turn-on-edit-mode modal');

        onEditButtonClick();
      };

      closeButton.onclick = function () {
        modal.style.display = 'none';

        restoreBlockedEditAttemptFocus();

        log('user closed turn-on-edit-mode modal');
      };
    }

    function bindViewOnlyModeModal() {
        var modal = document.getElementById('view-only-mode');
        var closeButton = document.getElementById('vom-close-btn');
        var confirmButton = document.getElementById('vom-сonfirm-btn');

        if (!modal || !closeButton || !confirmButton) {
            log('bindViewOnlyModeModal: modal or close button not found');

            return;
        }

        if (modal.__viewOnlyModeBound) {
            return;
        }

        modal.__viewOnlyModeBound = true;

        closeButton.onclick = function () {
            modal.style.display = 'none';

            log('user closed view-only-mode modal');
        };

        confirmButton.onclick = function () {
            modal.style.display = 'none';

            log('user closed view-only-mode modal');
        };
    }

      function bindViewerModeModal() {
          var modal = document.getElementById('viewer-mode');
          var closeButton = document.getElementById('vm-close-btn');
          var confirmButton = document.getElementById('vm-сonfirm-btn');

          if (!modal || !closeButton || !confirmButton) {
              log('bindViewerModeModal: modal or close button not found');

              return;
          }

          if (modal.__viewerModeBound) {
              return;
          }

          modal.__viewerModeBound = true;

          closeButton.onclick = function () {
              modal.style.display = 'none';

              log('user closed viewer-mode modal');
          };

          confirmButton.onclick = function () {
              modal.style.display = 'none';

              log('user closed viewer-mode modal');
          };
      }

    function createDotLoader(doc) {
      var dotLoader = doc.createElement('span');

      dotLoader.className = 'sk-dot-loader';

      new Array(3)
          .fill('.')
          .forEach(function (dot, index) {
            var dotElement = doc.createElement('span');

            dotElement.className = 'sk-dot-loader__dot';
            dotElement.textContent = dot;

            dotLoader.appendChild(dotElement);
          });

      return dotLoader;
    }

      // Drive the in-header status label from conflictState + currentMode + lockHolder.
      // Priority:
      //   1. conflictState       → "<user> edited the document"
      //   2. editing here        → "You are editing the document..."
      //   3. someone holds lock  → "<userName> is editing the document..."
      //   4. no lock in view     → hidden
      //
      // Conflict wins over the normal editing label because once another user saved
      // a newer version, the important action is reload/refresh, not edit-lock status.
      // In this state the label explains who changed the document.
      function renderEditingLabel() {
          if (!headerEditingLabel) {
              return;
          }

          var li = headerEditingLabel.parentNode;
          var doc = headerEditingLabel.ownerDocument;

          // Conflict label is useful even for viewers without edit rights, because it
          // explains why the document may need a refresh. Normal edit-lock labels stay
          // hidden for viewers below.
          if (conflictState) {
              var updatedBy = conflictState.updatedBy || 'Someone';
              const b = doc.createElement('span');

              b.className = 'sk-editing-label__who';
              b.textContent = updatedBy;

              headerEditingLabel.textContent = '';
              headerEditingLabel.appendChild(b);
              headerEditingLabel.appendChild(doc.createTextNode('\u00A0edited the document'));

              headerEditingLabel.classList.remove('sk-editing-label--self', 'sk-editing-label--other');
              headerEditingLabel.classList.add('sk-editing-label--conflict');

              if (li) {
                  li.style.display = '';
              }

              return;
          }

          // No edit right ⇒ never surface normal "who is editing" status.
          if (!canEdit || isDesktopClosing || isDesktopLoggingOut) {
              headerEditingLabel.classList.remove('sk-editing-label--conflict', 'sk-editing-label--self', 'sk-editing-label--other');
              headerEditingLabel.textContent = '';

              if (li) {
                  li.style.display = 'none';
              }

              return;
          }

          var isSelf = (currentMode === 'edit') || !!(lockHolder && lockHolder.isSelf);
          var show = isSelf || !!lockHolder;

          if (show) {
              // Lead token is bold (.sk-editing-label__who): "You" / the user's name.
              // Built from DOM nodes (not innerHTML) so a hostile userName can't inject.
              var who  = isSelf ? 'You' : ((lockHolder && lockHolder.userName) || 'Someone');
              // \u00A0 (non-breaking space) joins the bold name to the verb — a normal
              // leading space collapses at the inline-flex item boundary, gluing them.
              var rest = isSelf ? '\u00A0are editing the document' : '\u00A0is editing the document';
              const b = doc.createElement('span');

              b.className = 'sk-editing-label__who';
              b.textContent = who;

              headerEditingLabel.textContent = '';
              headerEditingLabel.appendChild(b);
              headerEditingLabel.appendChild(doc.createTextNode(rest));
              headerEditingLabel.appendChild(createDotLoader(doc));
          }

          headerEditingLabel.classList.remove('sk-editing-label--conflict');
          headerEditingLabel.classList.toggle('sk-editing-label--self', isSelf);
          headerEditingLabel.classList.toggle('sk-editing-label--other', show && !isSelf);

          if (li) {
              li.style.display = show ? '' : 'none';
          }
      }

      // Reflect currentMode/lockHolder/conflictState onto the in-header Edit button.
      // Normal states:
      //   view + lock-free + canEdit → free    (solid marine, enabled)  → request-edit-mode
      //   edit                       → editing (white + marine border)  → release (mode-changed:view)
      //   view + lock-held           → locked  (translucent, disabled)  → no-op
      //   !canEdit                   → hidden
      //
      // Conflict state:
      //   conflictState + lock-free  → refresh (blue, enabled) → reload-request
      //   conflictState + lock-held  → locked until the edit lock is released
      //
      // We intentionally reuse the same button/slot instead of mounting a second
      // control: layout stays stable, and the only visual differences from Edit are
      // icon, text and background colour.
      function renderEditButton() {
          if (!headerEditBtn) {
              return;
          }

          var li = headerEditBtn.parentNode;   // the .sk-edit-tab <li>
          var icon = headerEditBtn.querySelector('.sk-edit-btn__icon');
          var label = headerEditBtn.querySelector('.sk-edit-btn__label');

          headerEditBtn.classList.remove('is-editing', 'is-locked', 'is-refresh', 'is-refreshing');

          hideEditTooltip();

          // Conflict refresh is allowed even when the user has no edit rights,
          // because it only reloads the document and does not request edit mode.
          // It becomes available only after the edit lock is released.
          if (conflictState) {
              if (li) {
                  li.style.display = '';
              }

              headerEditBtn.classList.add('is-refresh');
              headerEditBtn.disabled = false;

              if (icon) {
                  icon.innerHTML = SK_REFRESH_ICON_SVG;
              }

              if (label) {
                  label.textContent = 'Refresh';
              }

              return;
          }

          if (pm.isExternal) {
              if (li) {
                  li.style.display = 'none';
              }

              headerEditBtn.disabled = true;

              return;
          }

          if (!canEdit) {
              headerEditBtn.classList.add('is-locked');
              headerEditBtn.disabled = false;
          }

          if (isDesktopClosing || isDesktopLoggingOut) {
              if (li) {
                  li.style.display = '';
              }

              headerEditBtn.classList.add('is-locked');
              headerEditBtn.disabled = false;

              if (label) {
                  label.textContent = 'Edit';
              }

              return;
          }

          if (li) {
              li.style.display = '';
          }

          if (icon) {
              icon.innerHTML = SK_EDIT_ICON_SVG;
          }

          if (label) {
              label.textContent = 'Edit';
          }

          if (editModeTransition !== null) {
              headerEditBtn.classList.add('is-locked');
              headerEditBtn.disabled = true;
          } else if (currentMode === 'view' && lockHolder) {
              headerEditBtn.classList.add('is-locked');
              headerEditBtn.disabled = true;
          } else if (currentMode === 'edit') {
              headerEditBtn.classList.add('is-editing');
              headerEditBtn.disabled = false;
          } else {
              headerEditBtn.disabled = false;
          }

          if (headerEditTooltip) {
              headerEditTooltip.textContent = getEditTooltipText();
          }
      }

      // Switch the conflict Refresh button into an in-flight state after click.
      // The page should reload shortly after `reload-request`, but if the host is
      // slow, the user still gets immediate feedback in the header: the button stays
      // in a disabled visual state, the spinner icon rotates, and the text shows
      // animated dots.
      function renderRefreshingButton() {
          if (!headerEditBtn) {
              return;
          }

          var icon = headerEditBtn.querySelector('.sk-edit-btn__icon');
          var label = headerEditBtn.querySelector('.sk-edit-btn__label');
          var doc = headerEditBtn.ownerDocument;

          headerEditBtn.classList.remove('is-editing', 'is-locked', 'is-refresh');
          headerEditBtn.classList.add('is-refreshing');
          headerEditBtn.disabled = true;

          if (icon) {
              icon.innerHTML = SK_REFRESHING_ICON_SVG;
          }

          if (label) {
              label.textContent = '';
              label.appendChild(doc.createTextNode('Refreshing'));
              label.appendChild(createDotLoader(doc));
          }
      }

    // Reflect the save lifecycle onto the diskette button. States:
    //   idle    — nothing to save (initial)        (NOT clickable)
    //   dirty   — unsaved changes pending          (clickable → save now)
    //   saving  — capture/convert/send in flight   (NOT clickable)
    //   saved   — host ack'd the save (success)    (NOT clickable — nothing to save)
    //   error   — save failed                      (clickable → retry)
    // The button is only clickable when there's something to do: 'dirty' (save
    // now) or 'error' (retry). 'idle'/'saving'/'saved' are disabled — nothing to
    // save, or a save is already in flight. No-op until the button is injected.
    function renderSaveButton() {
      if (!headerSaveBtn) {
          return;
      }

      var s = currentSaveState;
      var clickable = (s === 'dirty' || s === 'error');
      headerSaveBtn.className = 'sk-save-btn sk-save-btn--' + s;
      headerSaveBtn.disabled = !clickable;
      headerSaveBtn.setAttribute('aria-busy', s === 'saving' ? 'true' : 'false');
      // saved/error are distinct icons (diskette + badge); the rest share the
      // bare diskette and differ only by currentColor (grey idle / black dirty).
      var iconEl = headerSaveBtn.querySelector('.sk-save-btn__icon');
      if (iconEl) {
        iconEl.innerHTML = s === 'saved' ? SK_SAVE_ICON_SAVED
                         : s === 'error' ? SK_SAVE_ICON_ERROR
                         : SK_SAVE_ICON_SVG;
      }
        if (headerSaveTooltip) {
            headerSaveTooltip.textContent = getSaveTooltipText();
        }
    }

    // Called by wrapper-postmessage.js (via window.skSetSaveState) whenever the
    // save lifecycle advances. Single funnel so the diskette always mirrors the
    // real save state — including saves triggered by edit-mode-off / page-leave.
    window.skSetSaveState = function (state) {
      currentSaveState = state || 'idle';
      renderSaveButton();
    };

    function renderDownloadButton() {
      if (!headerDownloadBtn) {
          return;
      }

      headerDownloadBtn.disabled = !canDownload || isDownloadStarting || isDownloading;

      if (headerDownloadTooltip) {
          headerDownloadTooltip.textContent = getDownloadTooltipText();
      }
    }

    // Hide the native OnlyOffice "Editing/Viewing" dropdown inside the
    // editor iframe — permanently. Same-origin, so contentDocument is
    // accessible. Idempotent — reuses a single <style id="hide-native-
    // dropdown"> element.
    function hideNativeDropdown() {
      var iframe = document.querySelector('iframe[name="frameEditor"]');

      if (!iframe || !iframe.contentDocument) {
          return;
      }
      var doc = iframe.contentDocument;
      var style = doc.getElementById('hide-native-dropdown');

      if (!style) {
        style = doc.createElement('style');
        style.id = 'hide-native-dropdown';

        if (doc.head) {
            doc.head.appendChild(style);
        }
      }
      // Hide both the dropdown AND the canRequestEditRights button (if any
      // legacy config flips it back on by accident).
      style.textContent =
        '.btn-header-pdf-mode { display: none !important; }' +
        '.btn-header-pdf-mode + * { display: none !important; }';
    }

    // Pencil icon — Figma "Icons 20px / Stroke set" (node 15802:85791),
    // exported as a 20×20 path and recoloured via currentColor so one markup
    // serves all three button states. Don't hand-edit the path; re-export
    // from Figma if the icon changes.
    var SK_EDIT_ICON_SVG =
      '<svg viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg" ' +
      'width="20" height="20" aria-hidden="true" focusable="false">' +
      '<path fill="currentColor" d="M3.96541 16.0566C3.83119 15.9234 3.76217 15.7378 3.77552 15.5491L3.96679 12.865L3.98544 12.7497C4.01415 12.6379 4.07279 12.5355 4.15531 12.4528L12.4369 4.17118C12.4778 4.13023 12.524 4.09422 12.5743 4.06553L12.8968 4.62969C12.6079 4.12332 12.5785 4.07104 12.5757 4.06553L12.575 4.06484L12.5764 4.06484C12.5772 4.06439 12.5782 4.06259 12.5791 4.06208C12.5812 4.06097 12.5842 4.06001 12.5867 4.05862C12.592 4.05574 12.5986 4.05155 12.6061 4.04757C12.6214 4.03947 12.6418 4.02961 12.6661 4.01788C12.715 3.99432 12.7821 3.96414 12.8643 3.93295C13.0273 3.87114 13.259 3.80141 13.5341 3.77136C14.0993 3.70964 14.8477 3.82007 15.5042 4.4764C16.1516 5.12386 16.2642 5.87621 16.2155 6.43751C16.1915 6.7131 16.1291 6.94758 16.0732 7.11285C16.0453 7.19554 16.0186 7.2637 15.9973 7.31311C15.9866 7.33772 15.9763 7.35907 15.969 7.37457C15.9654 7.38215 15.9626 7.38935 15.96 7.39459C15.9587 7.39721 15.9569 7.39948 15.9558 7.4015L15.9545 7.40426L15.9538 7.40495L15.9538 7.40633L15.9531 7.40702C15.9223 7.4668 15.8825 7.52174 15.835 7.5693L7.56171 15.8426C7.4524 15.9517 7.30693 16.0182 7.15292 16.0304L4.47503 16.2431C4.28645 16.2579 4.09971 16.1899 3.96541 16.0566ZM14.195 7.36973L14.8254 6.73928C14.8303 6.72626 14.8367 6.71225 14.842 6.69646C14.8734 6.60354 14.9079 6.47306 14.9207 6.32496C14.9456 6.03865 14.8926 5.70456 14.5851 5.39688C14.2663 5.07805 13.9394 5.0344 13.6743 5.06335C13.5349 5.07868 13.4128 5.11524 13.3256 5.14829C13.307 5.15534 13.2903 5.16286 13.2759 5.169L12.6351 5.80982L14.195 7.36973ZM13.2151 5.19662L13.2165 5.19662L13.2179 5.19524L13.2172 5.19455C13.2154 5.19549 13.2139 5.19681 13.213 5.19731L13.213 5.19869C13.2135 5.19846 13.2142 5.19714 13.2151 5.19662ZM6.81179 14.7529L13.2752 8.28952L11.7153 6.72961L5.24635 13.1985L5.12551 14.8869L6.81179 14.7529Z"/></svg>';

    // Refresh icon — used only in conflict state, when another user has saved a
    // newer version of the document. The button keeps the same DOM slot as Edit,
    // but its icon/text/background switch to Refresh.
    var SK_REFRESH_ICON_SVG =
        '<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg" ' +
        'aria-hidden="true" focusable="false">' +
        '<path d="M11.6651 11.2441C11.8633 11.4518 11.8551 11.7812 11.6476 11.9795C10.6796 12.903 9.3669 13.4706 7.92296 13.4707C4.97121 13.4706 2.5758 11.1011 2.53038 8.16016L1.75109 8.93848C1.54811 9.14146 1.21981 9.14127 1.01671 8.93848C0.81377 8.73539 0.812706 8.40615 1.01573 8.20312L2.67394 6.5459C2.88022 6.33961 3.21472 6.33963 3.42101 6.5459L5.07921 8.20312C5.28218 8.40609 5.28197 8.73538 5.07921 8.93848C4.87613 9.14147 4.54691 9.14152 4.34386 8.93848L3.56945 8.16309C3.61629 10.5283 5.54659 12.4315 7.92296 12.4316C9.08912 12.4316 10.1475 11.9731 10.9298 11.2266C11.1376 11.0286 11.467 11.0365 11.6651 11.2441ZM14.9835 7.06055C15.1865 7.26357 15.1874 7.59281 14.9845 7.7959L13.3263 9.4541C13.12 9.6603 12.7855 9.66034 12.5792 9.4541L10.921 7.7959C10.7185 7.59298 10.7186 7.26452 10.921 7.06152C11.124 6.85857 11.4533 6.85782 11.6564 7.06055L12.4318 7.83496C12.3841 5.47057 10.453 3.56764 8.07726 3.56738C6.91116 3.56738 5.85272 4.0261 5.07042 4.77246C4.86267 4.97063 4.53332 4.96256 4.33507 4.75488C4.13687 4.54711 4.1449 4.21778 4.35265 4.01953C5.32062 3.09613 6.63338 2.52832 8.07726 2.52832C11.0287 2.52858 13.4251 4.89827 13.4708 7.83887L14.2491 7.06055C14.452 6.8578 14.7804 6.85807 14.9835 7.06055Z" fill="white"/>' +
        '</svg>';
    // Refreshing spinner icon — used only while the Refresh request is already sent
    // and we are waiting for the host to reload the document.
    var SK_REFRESHING_ICON_SVG =
        '<svg class="sk-refreshing-icon" xmlns="http://www.w3.org/2000/svg" ' +
        'width="12" height="12" viewBox="0 0 12 12" fill="none" ' +
        'aria-hidden="true" focusable="false">' +
        '<circle cx="6" cy="6" r="5.35" stroke="#355069" stroke-opacity="0.1" stroke-width="1.3"/>' +
        '<path d="M6 0.65C7.4189 0.65 8.77968 1.21365 9.783 2.21696C10.7863 3.22027 11.35 4.58105 11.35 5.99995" ' +
        'stroke="#355069" stroke-opacity="0.55" stroke-width="1.3" stroke-linecap="round"/>' +
        '</svg>';

    var DOWNLOAD_ICON_SVG = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
        '<path d="M19.1494 13.9492C19.5084 13.9492 19.7998 14.2406 19.7998 14.5996V17.1064C19.7998 18.5937 18.5937 19.7998 17.1064 19.7998H6.89258C5.40535 19.7998 4.19922 18.5937 4.19922 17.1064V14.5996C4.19922 14.2406 4.49062 13.9492 4.84961 13.9492C5.20859 13.9492 5.5 14.2406 5.5 14.5996V17.1064C5.5 17.8757 6.12332 18.5 6.89258 18.5H17.1064C17.8757 18.5 18.5 17.8757 18.5 17.1064V14.5996C18.5 14.2407 18.7905 13.9493 19.1494 13.9492ZM12.002 4.19922C12.3608 4.19941 12.6514 4.49074 12.6514 4.84961V11.7217L14.79 9.57031C15.0429 9.31587 15.4545 9.31462 15.709 9.56738C15.9635 9.82028 15.9648 10.2318 15.7119 10.4863L12.4609 13.7568C12.339 13.8795 12.173 13.9492 12 13.9492C11.8273 13.949 11.6618 13.8794 11.54 13.7568L8.29004 10.4863C8.03734 10.2318 8.03863 9.82022 8.29297 9.56738C8.54745 9.31457 8.95903 9.31594 9.21191 9.57031L11.3516 11.7236V4.84961C11.3516 4.49065 11.643 4.19925 12.002 4.19922Z" fill="currentColor"/>' +
        '</svg>';

    // Save (diskette) icon — Figma "Frame 3205" (node 15904:12420), 20×20.
    // The diskette is ONE path recoloured via currentColor: per design it has
    // only TWO colours — grey (#A8A8A8, idle/saved/rest) and black (#363636,
    // unsaved/dirty) — set per-state by the CSS class on the button. saved &
    // error are DISTINCT icons (the same diskette + a coloured corner badge),
    // NOT a recolour of the whole glyph; renderSaveButton() swaps them in.
    // Don't hand-edit the diskette path; re-export from Figma.
    //   TODO(icons): the saved/error badges below are PLACEHOLDERS built from
    //   the diskette + a check / cross badge (Figma MCP rate-limited; real
    //   nodes 15904:12441 / :12453 pending). Drop in the exact SVGs when
    //   available — nothing else changes.
    var SK_DISKETTE_PATH =
      'M14.3 5.25H8.2M14.3 5.25V9.2H8.2V5.25M14.3 5.25L14.8 5.25L18.749 9.2V16.75C18.749 17.8546 17.8536 18.75 16.749 18.75H15.8M15.8 18.75H8.2M15.8 18.75V13.15H8.2V18.75M8.2 5.25H7.25098C6.14641 5.25 5.25098 6.14543 5.25098 7.25V16.75C5.25098 17.8546 6.14641 18.75 7.25098 18.75H8.2';

    function skSvg(inner) {
      return '<svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" ' +
        'width="24" height="24" aria-hidden="true" focusable="false">' + inner + '</svg>';
    }
    // idle / dirty / saving — bare diskette, colour from currentColor.
    var SK_SAVE_ICON_SVG = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
        '<path d="M15.0938 4.21191C15.2185 4.23694 15.3345 4.29836 15.4258 4.38965L19.6094 8.5752C19.7309 8.69701 19.7998 8.8621 19.7998 9.03418V17.1084C19.7995 18.5956 18.5937 19.8018 17.1064 19.8018H6.89258C5.40544 19.8017 4.19948 18.5955 4.19922 17.1084V6.89258C4.19932 5.40534 5.40534 4.19932 6.89258 4.19922H14.9658L15.0938 4.21191ZM6.89258 5.5C6.12331 5.5001 5.5001 6.12331 5.5 6.89258V17.1084C5.50026 17.8775 6.12341 18.5019 6.89258 18.502H7.32422V13.2188C7.32433 12.8599 7.61481 12.5694 7.97363 12.5693H16.0254C16.3843 12.5693 16.6757 12.8599 16.6758 13.2188V18.502H17.1064C17.8757 18.502 18.4997 17.8776 18.5 17.1084V9.30371L15.1074 5.91113V9.03418C15.1072 9.39278 14.8166 9.68431 14.458 9.68457H7.97363C7.6149 9.68446 7.32448 9.39287 7.32422 9.03418V5.5H6.89258ZM8.62402 18.502H15.375V13.8691H8.62402V18.502ZM8.62402 8.38477H13.8076V5.5H8.62402V8.38477Z" fill="currentColor"/>' +
        '</svg>';
    // saved — diskette + teal check badge (#3FC0C4, matches the Edit button;
    // bottom-right, white halo cuts it out of the diskette).
    var SK_SAVE_ICON_ERROR = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
        '<path d="M19.1504 14.9893C21.4476 14.9895 23.3096 16.8524 23.3096 19.1494C23.3094 21.4462 21.4475 23.3084 19.1504 23.3086C16.8531 23.3086 14.9904 21.4463 14.9902 19.1494C14.9902 16.8523 16.853 14.9893 19.1504 14.9893ZM20.9619 17.3389C20.7082 17.0851 20.2959 17.0843 20.042 17.3379L19.1494 18.2295L18.2588 17.3389C18.005 17.0851 17.5927 17.0843 17.3389 17.3379C17.0851 17.5917 17.0853 18.004 17.3389 18.2578L18.2305 19.1494L17.3389 20.042C17.0854 20.2958 17.0854 20.7072 17.3389 20.9609C17.5926 21.2147 18.0039 21.2145 18.2578 20.9609L19.1494 20.0684L20.043 20.9619C20.2968 21.2152 20.7082 21.2153 20.9619 20.9619C21.2156 20.7082 21.2152 20.2969 20.9619 20.043L20.0693 19.1504L20.9619 18.2588C21.2157 18.005 21.2155 17.5927 20.9619 17.3389Z" fill="#FF274B"/>' +
        '<path fill-rule="evenodd" clip-rule="evenodd" d="M15.0938 4.21191C15.2185 4.23694 15.3345 4.29836 15.4258 4.38965L19.6094 8.5752C19.7309 8.69701 19.7998 8.8621 19.7998 9.03418V13.0386C19.7998 13.3975 19.5088 13.6885 19.1499 13.6885C18.791 13.6885 18.5 13.3975 18.5 13.0386V9.30371L15.1074 5.91113V9.03418C15.1072 9.39278 14.8166 9.68431 14.458 9.68457H7.97363C7.6149 9.68446 7.32448 9.39287 7.32422 9.03418V5.5H6.84961C6.10402 5.5 5.5 6.10403 5.5 6.84961V17.1514C5.50026 17.8967 6.10419 18.502 6.84961 18.502H7.32422V13.2188C7.32433 12.8599 7.61481 12.5694 7.97363 12.5693H16.0254C16.3843 12.5693 16.6757 12.8599 16.6758 13.2188V13.8594C16.6758 14.2186 16.3797 14.5098 16.0205 14.5098C15.6667 14.5098 15.375 14.2229 15.375 13.8691H8.62402V18.502H13.0381C13.3968 18.502 13.6877 18.7926 13.688 19.1514C13.6883 19.5105 13.3972 19.8018 13.0381 19.8018H6.84961C5.38622 19.8018 4.19948 18.6147 4.19922 17.1514V6.84961C4.19922 5.38605 5.38606 4.19922 6.84961 4.19922H14.9658L15.0938 4.21191ZM8.62402 8.38477H13.8076V5.5H8.62402V8.38477Z" fill="currentColor"/>' +
        '</svg>';
    // error — diskette + red cross badge (bottom-right).
    var SK_SAVE_ICON_SAVED = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
        '<path fill-rule="evenodd" clip-rule="evenodd" d="M15.0938 4.21191C15.2185 4.23694 15.3345 4.29836 15.4258 4.38965L19.6094 8.5752C19.7309 8.69701 19.7998 8.8621 19.7998 9.03418V13.0386C19.7998 13.3975 19.5088 13.6885 19.1499 13.6885C18.791 13.6885 18.5 13.3975 18.5 13.0386V9.30371L15.1074 5.91113V9.03418C15.1072 9.39278 14.8166 9.68431 14.458 9.68457H7.97363C7.6149 9.68446 7.32448 9.39287 7.32422 9.03418V5.5H6.84961C6.10402 5.5 5.5 6.10403 5.5 6.84961V17.1514C5.50026 17.8967 6.10419 18.502 6.84961 18.502H7.32422V13.2188C7.32433 12.8599 7.61481 12.5694 7.97363 12.5693H16.0254C16.3843 12.5693 16.6757 12.8599 16.6758 13.2188V13.8594C16.6758 14.2186 16.3797 14.5098 16.0205 14.5098C15.6667 14.5098 15.375 14.2229 15.375 13.8691H8.62402V18.502H13.0381C13.3968 18.502 13.6877 18.7926 13.688 19.1514C13.6883 19.5105 13.3972 19.8018 13.0381 19.8018H6.84961C5.38622 19.8018 4.19948 18.6147 4.19922 17.1514V6.84961C4.19922 5.38605 5.38606 4.19922 6.84961 4.19922H14.9658L15.0938 4.21191ZM8.62402 8.38477H13.8076V5.5H8.62402V8.38477Z" fill="currentColor"/>' +
        '<path d="M19 14.8398C21.2972 14.84 23.1591 16.7025 23.1592 19C23.1592 21.2975 21.2972 23.16 19 23.1602C16.7026 23.1602 14.8398 21.2977 14.8398 19C14.8399 16.7024 16.7027 14.8398 19 14.8398ZM21.4873 17.1885C21.2336 16.9349 20.8222 16.935 20.5684 17.1885L18.3232 19.4326L17.4316 18.541C17.1778 18.2872 16.7656 18.2872 16.5117 18.541C16.2584 18.7948 16.2582 19.2063 16.5117 19.46L17.8643 20.8125C17.986 20.934 18.1512 21.0018 18.3232 21.002C18.4954 21.002 18.6613 20.9341 18.7832 20.8125L21.4873 18.1084C21.7411 17.8546 21.7411 17.4423 21.4873 17.1885Z" fill="#2FA0AF"/>' +
        '</svg>';

    // Main App — arrow-into-box icon (16×16), recoloured via currentColor.
    var SK_MAIN_APP_ICON_SVG =
      `<svg width="16" height="16" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M11.5967 2.15039C12.8387 2.15039 13.8457 3.15742 13.8457 4.39941V11.6006C13.8457 12.8426 12.8387 13.8496 11.5967 13.8496H4.39551C3.15361 13.8495 2.14648 12.8425 2.14648 11.6006V10.3389C2.14648 9.98081 2.43689 9.69054 2.79492 9.69043C3.15304 9.69043 3.44336 9.98074 3.44336 10.3389V11.6006C3.44336 12.1261 3.87 12.5526 4.39551 12.5527H11.5967C12.1223 12.5527 12.5488 12.1262 12.5488 11.6006V4.39941C12.5488 3.87382 12.1223 3.44727 11.5967 3.44727H4.39551C3.87003 3.44742 3.44341 3.87391 3.44336 4.39941V5.66504C3.44336 6.02316 3.15304 6.31348 2.79492 6.31348C2.43689 6.31337 2.14648 6.0231 2.14648 5.66504V4.39941C2.14653 3.15752 3.15364 2.15054 4.39551 2.15039H11.5967ZM6.52539 5.54297C6.77775 5.28919 7.18838 5.28795 7.44238 5.54004L9.45703 7.54199C9.57946 7.66372 9.64845 7.83028 9.64844 8.00293C9.64824 8.17539 9.57939 8.34133 9.45703 8.46289L7.44238 10.4639C7.18835 10.716 6.77774 10.7148 6.52539 10.4609C6.27316 10.2069 6.27454 9.79634 6.52832 9.54395L7.42773 8.65039H0.698242C0.340375 8.65016 0.0499313 8.35984 0.0498047 8.00195C0.0498204 7.64397 0.340307 7.35374 0.698242 7.35352H7.42773L6.52832 6.45996C6.2746 6.20756 6.27324 5.79694 6.52539 5.54297Z" fill="#2FA0AF"/>
      </svg>`;

    // Figma-derived styles for our header controls. Injected into the iframe
    // <head> (same mechanism as hideNativeDropdown). Unique sk-* class names
    // so the editor's own stylesheets never match our elements; injected late
    // so ours wins on equal specificity (no !important needed).
    var SK_HEADER_CONTROLS_CSS = [
      /* our <li> sits at the end of the toolbar tablist, right after "View".
         The toolbar-mask (view mode) starts at top:32px, so the tab row — and
         thus this button — stays visible AND clickable while read-only.
         inline-flex (not flex) keeps it flowing inline among the inline-block
         tabs; height matches the ~28px tab row so the button can't protrude. */
      '.box-tabs{height:30px !important;}',                         /* tab row height (was 28px) — per design; !important beats .toolbar .box-tabs / .top-title>.toolbar .box-tabs */
      /* Sharekey header logo (customization.logo): Header.setBranding writes an
         inline `max-height:20px` on the <img>, so !important is required to bump
         it to the logo's native 24px. */
      '#header-logo{cursor: pointer; padding-right: 17px !important;}',
      '#header-logo img{max-height:24px !important;}',
      /* Remove the "From Text/CSV" group from the xlsx Data tab. The button slot
         (#slot-btn-data-from-text) has no data-layout-name, so the layout config
         can't reach it — hide its whole .group plus the trailing separator so no
         empty gap is left. :has() is fine (editor runs in modern Chrome). Scoped
         to the spreadsheet's Data panel; the selector simply doesn't match in
         word/slide. */
      '.panel[data-tab="data"] .group:has(#slot-btn-data-from-text),',
      '.panel[data-tab="data"] .group:has(#slot-btn-data-from-text)+.separator{display:none !important;}',
      '.sk-edit-tab{display:inline-flex;align-items:center;height:28px;vertical-align:top;margin-left:8px;list-style:none;}',
      /* kill the native tab-hover chrome on OUR li: the inset bottom box-shadow
         + grey hover background (toolbar.less:160-163), and the ::after bottom
         underline the tabs draw (revealed on hover). Our li isn't a real tab,
         so it needs none of it. !important beats the non-important native rules. */
      '.sk-edit-tab:hover{box-shadow:none !important;background-color:transparent !important;}',
      '.sk-edit-tab::after,.sk-edit-tab::before{display:none !important;}',
        '.tooltip, .sk-tooltip{',
        '  box-sizing:border-box;',
        '  width:max-content;',
        '  max-width:none;',
        '  min-height:18px;',
        '  padding:2px 8px !important;',
        '  white-space:nowrap;',
        '  border-radius:3px;',
        '  background:#728596;',
        '  color:#FFFFFF;',
        "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;",
        '  font-size:10px;',
        '  line-height:14px;',
        '  font-weight:400;',
        '  text-align:center;',
        '  box-shadow:none;',
        '}',

        '.tooltip .tooltip-inner{',
        '  padding:0;',
        '  background:transparent;',
        '  color:inherit;',
        '  font:inherit;',
        '  max-width:none;',
        '}',

        '.sk-tooltip{',
        '  position:fixed;',
        '  display:none;',
        '  height:18px;',
        '  z-index:100000;',
        '  pointer-events:none;',
        '}',

        '.tooltip *{',
        '  border:none !important;',
        '  outline:none !important;',
        '  box-shadow:none !important;',
        "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif !important;",
        '  font-size:10px !important;',
        '  line-height:14px !important;',
        '  font-weight:400 !important;',
        '  text-align:center !important;',
        '}',
      '.sk-edit-btn{',
      '  box-sizing:border-box;',
      '  display:inline-flex;align-items:center;justify-content:center;',
      '  height:24px;padding:0 8px 0 4px;',                        /* fits inside the 28px tab row (was 32px → protruded) */
      '  border:1px solid transparent;border-radius:5px;',          /* keeps box size stable across bordered/unbordered states */
      '  background:#3fc0c4;color:#FFFFFF;',                         /* free state; color drives icon (currentColor) + label */
      "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;",
      '  font-size:12px;line-height:16px;font-weight:500;',
      '  white-space:nowrap;cursor:pointer;',
      '  -webkit-appearance:none;appearance:none;margin-bottom:5px;',  /* margin-bottom lifts the button within the row */
      '  transition:background .12s ease,border-color .12s ease,color .12s ease;',
      '}',
      '.sk-edit-btn__icon{display:flex;flex:0 0 auto;width:16px;height:16px;margin-right:4px}',
      '.sk-edit-btn__icon svg{display:block;width:16px;height:16px;}',
      '.sk-edit-btn__label{display:block;}',
      /* free (default) hover — not while pending/disabled */
      '.sk-edit-btn:not(.is-editing):not(.is-locked):not(.is-refresh):not(:disabled):hover{background:#41D1C9;}',
      /* editing — you hold the lock */
      '.sk-edit-btn.is-editing{background:#E2F5F6;color:#2FA0AF;border-color:#2FA0AF;}',
      '.sk-edit-btn.is-editing:hover{background:#CFEDEF;}',
      /* locked — someone else is editing (disabled) */
      '.sk-edit-btn.is-locked{background:rgba(53,80,105,0.15);color:#FFFFFF;border-color:transparent;cursor:default;}',
      '.sk-edit-btn:disabled{cursor:default;pointer-events:none;}',

      /* refresh — conflict state. Same button shape/spacing as Edit, but with
      Refresh icon/text and the blue background from design. */
      '.sk-edit-btn.is-refresh{background:#0F8CC9;color:#FFFFFF;border-color:transparent;}',
      '.sk-edit-btn.is-refresh:not(.is-refreshing):not(:disabled):hover{background:#3FA3D4;}',
      '.sk-edit-btn.is-refresh:not(.is-refreshing):not(:disabled):active{background:#3FA3D4;}',
      '.sk-edit-btn.is-refresh .sk-edit-btn__icon{margin-right:4px;}',

      /* refreshing — disabled in-flight state after the user clicks Refresh.
      Uses a separate spinner icon and muted disabled colours from design. */
      '@keyframes sk-refreshing-spin{from{transform:rotate(0deg);}to{transform:rotate(360deg);}}',
      '.sk-edit-btn.is-refreshing, .sk-edit-btn.is-refreshing:disabled{background:rgba(53,80,105,0.1);color:rgba(53,80,105,0.55);border-color:transparent;cursor:default;opacity:1;}',
      '.sk-edit-btn.is-refreshing:hover, .sk-edit-btn.is-refreshing:active, .sk-edit-btn.is-refreshing:disabled:hover, .sk-edit-btn.is-refreshing:disabled:active{background:rgba(53,80,105,0.1);color:rgba(53,80,105,0.55);}',
      '.sk-edit-btn.is-refreshing .sk-edit-btn__icon{width:12px;height:12px;margin-right:4px;}',
      '.sk-edit-btn.is-refreshing .sk-edit-btn__icon svg{width:12px;height:12px;}',
      '.sk-refreshing-icon{display:block;animation:sk-refreshing-spin 2s linear infinite;}',

      /* ── "<who> is editing the document…" label ───────────────────────────
         Sits in the tab row just right of the Edit button. renderEditingLabel
         toggles the <li> display + the --self/--other modifier classes.
         TODO(figma): placeholder styling — exact colours / font / spacing
         pending the Figma specs (nodes 15802:85984 self, 15802:86051 other;
         MCP rate-limited). Tweak the values below once available. */
      '.sk-editing-tab{display:inline-flex;align-items:center;height:28px;vertical-align:top;margin-left:8px;list-style:none;}',
      /* Win the cascade: OnlyOffice's `.toolbar .tabs li{align-items:end}` (0,2,1)
         beats a bare `.sk-editing-tab` (0,1,0), so the center above never applied.
         This selector (0,3,1) overrides it. NOTE: with centering now active, the
         label's `margin-bottom:8px` (below) nudges the centered label upward — keep
         or drop that margin depending on the look you want. */
      '.toolbar .tabs li.sk-editing-tab{align-items:center;}',
      '.sk-editing-tab:hover{box-shadow:none !important;background-color:transparent !important;}',
      '.sk-editing-tab::after,.sk-editing-tab::before{display:none !important;}',
      '.sk-editing-label{',
      '  display:inline-flex;',                                     /* centring is on .sk-editing-tab (the <li>) */
      "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;",
      '  font-size:12px;line-height:1;font-weight:400;white-space:nowrap;',
      '  color:#355069B2;',                                         /* #3: rgba(53,80,105,.7) */
      // The <li>.sk-editing-tab inherits `align-items:end` from OnlyOffice's
      // more-specific `.toolbar .tabs li` (our `.sk-editing-tab{align-items:center}`
      // loses on specificity), so the row is bottom-aligned. This bottom margin
      // lifts the label to baseline-align with the Edit button. Don't remove it.
      '  margin-bottom:5px;',
      '}',
      '.sk-editing-label__who{font-weight:600;}',                   /* #4: the editing user (or "You") — semibold */
      '.theme-type-dark .sk-editing-label{color:rgba(255,255,255,0.7);}', /* dark theme: #355069 would vanish */

      /* Conflict label: "Updated by <user>" reuses the same label slot as the
      normal editing-status text. Kept as a modifier class so the base spacing,
      font and dark-theme fallback stay shared. */

      '.sk-editing-label--conflict{color:#355069B2;}',
      '.theme-type-dark .sk-editing-label--conflict{color:rgba(255,255,255,0.7);}',

      '@keyframes sk-dot-loader-blink{0%{opacity:.2;}20%{opacity:1;}100%{opacity:.2;}}',
      '.sk-dot-loader{display:inline;}',
      '.sk-dot-loader__dot{color:inherit;animation-name:sk-dot-loader-blink;animation-duration:1.4s;animation-iteration-count:infinite;animation-fill-mode:both;}',
      '.sk-dot-loader__dot:nth-of-type(2){animation-delay:.2s;}',
      '.sk-dot-loader__dot:nth-of-type(3){animation-delay:.4s;}',

      /* ── Save (diskette) button — replaces the native quick-access Save ── */
      '#slot-btn-dt-save{display:none !important;}',                /* hide native; ours takes its place */
      '.sk-save-slot{display:inline-flex;align-items:center;}',
      '.sk-save-btn{',
      '  box-sizing:border-box; margin-top: 2px;',
      '  display:inline-flex;align-items:center;justify-content:center;',
      '  width:28px;height:28px;padding:0;',
      '  border:none;border-radius:5px;background:transparent;',
      '  color:#A8A8A8;',                                          /* idle (Figma) — drives the diskette via currentColor */
      '  cursor:pointer;-webkit-appearance:none;appearance:none;',
      '  transition:color .12s ease,background .12s ease,opacity .12s ease;',
      '}',
      '.sk-save-btn__icon{display:flex;width:24px;height:24px;}',
      '.sk-save-btn__icon svg{display:block;width:24px;height:24px;}',
      '.sk-save-btn:not(:disabled):hover{background:rgba(0,0,0,0.06);}',
      /* The diskette has only two colours per design: grey at rest, black when
         there are unsaved changes. saved/error swap to a DISTINCT badge icon
         (renderSaveButton) — the diskette under the badge keeps its base colour:
         grey for saved (back to rest), black for error (changes still pending). */
      '.sk-save-btn--idle{color rgba(168, 168, 168, 1);}',
      '.sk-save-btn--dirty{color:#363636;}',
      '.sk-save-btn--saving{color:#363636;opacity:.5;cursor:default;}',
      '.sk-save-btn--saved{color:rgba(168, 168, 168, 1);}',
      '.sk-save-btn--error{color:#363636;}',
      '.sk-save-btn:disabled{cursor:default;}',
      /* Dark theme (editor sets body.theme-type-dark): the "dark" dirty glyph
         would vanish on the dark toolbar, so flip the prominent states to a
         light glyph. Idle stays mid-grey (visible on both). The saved/error
         badges carry their own fills, so only the diskette base needs flipping.
         More specific than the base rules → wins; injected late so no !important. */
      '.theme-type-dark .sk-save-btn{color:#A8A8A8;}',
      '.theme-type-dark .sk-save-btn--idle{color:rgba(168, 168, 168, 1);}',
      '.theme-type-dark .sk-save-btn--dirty{color:#FFFFFF;}',
      '.theme-type-dark .sk-save-btn--saving{color:#FFFFFF;opacity:.5;}',
      '.theme-type-dark .sk-save-btn--saved{color:rgba(168, 168, 168, 1);}',
      '.theme-type-dark .sk-save-btn--error{color:#FFFFFF;}',
      '.theme-type-dark .sk-save-btn:not(:disabled):hover{background:rgba(255,255,255,0.1);}',
		/* This is the slot for the buttons in the header. We need to add a margin to the bottom of the slot to make the buttons align correctly. */
	  '#slot-btn-search{margin-bottom:5px}',
      /* ── Main App button — header-right, before the search slot. */
      '.sk-main-app-slot{display:inline-flex;align-items:center;margin-right:8px;vertical-align:middle;}',
      '.sk-main-app-btn{',
      '  box-sizing:border-box;',
      '  display:inline-flex;align-items:center;gap:4px;',
      '  height:24px;padding:0 8px 0 4px;',
      '  border:1px solid #2FA0AF;border-radius:5px;',
      '  background:white;color:#2FA0AF;',
      "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;",
      '  font-size:12px;line-height:16px;font-weight:500;white-space:nowrap;',
      '  cursor:pointer;-webkit-appearance:none;appearance:none;margin-bottom: 5px;',
      '  transition:background .12s ease;',
      '}',
      '.sk-main-app-btn__icon{display:flex;flex:0 0 auto;flex-shrink:0;width:16px;height:16px;}',
      '.sk-main-app-btn__icon svg{display:block;flex-shrink:0;width:16px;height:16px;}',
      '.sk-main-app-btn:not(:disabled):hover{background:#F0F8F9;}',
      '.theme-type-dark .sk-main-app-btn{background:rgba(255,255,255,0.1);color:rgba(255,255,255,0.8);}',
      '.theme-type-dark .sk-main-app-btn:not(:disabled):hover{background:rgba(255,255,255,0.2);}',
      '#slot-btn-dt-quick-access{display:none !important;}',
        '.download-slot{display:inline-flex;align-items:center;}',
        '.download-btn{',
        '  box-sizing:border-box;margin-top:2px;',
        '  display:inline-flex;align-items:center;justify-content:center;',
        '  width:28px;height:28px;padding:0;',
        '  border:none;border-radius:5px;background:transparent;',
        '  color:#363636;',
        '  cursor:pointer;-webkit-appearance:none;appearance:none;',
        '  transition:color .12s ease,background .12s ease;',
        '}',
        '.download-btn__icon{display:flex;width:24px;height:24px;}',
        '.download-btn__icon svg{display:block;width:24px;height:24px;}',
        '.download-btn:not(:disabled):hover{background:rgba(0,0,0,0.06);}',
        '.theme-type-dark .download-btn{color:#FFFFFF;}',
        '.theme-type-dark .download-btn:not(:disabled):hover{background:rgba(255,255,255,0.1);}',
        '.download-btn:disabled{opacity: var(--component-disabled-opacity, 0.4);}',
        '.btn-header:not(.dropdown-toggle):hover:not(:disabled){background:rgba(0,0,0,0.06) !important;}',
        '.btn-header:not(.dropdown-toggle) {',
        '  width:28px !important; height:28px !important; margin: 2px 4px 0 4px !important; display: flex !important; align-items: center !important; justify-content: center !important;',
        '}'
    ].join('\n');

    function injectHeaderControlStyles(doc) {
      var style = doc.getElementById('sk-header-controls');

      if (!style) {
        style = doc.createElement('style');
        style.id = 'sk-header-controls';

        if (doc.head) {
            doc.head.appendChild(style);
        }
      }

      style.textContent = SK_HEADER_CONTROLS_CSS;
      bindHeaderLogoClick(doc);
    }

      function bindHeaderLogoClick(doc) {
          var headerLogo = doc.getElementById('header-logo');

          if (!headerLogo) {
              return false;
          }

          if (headerLogo.__welcomeScreenClickBound) {
              return true;
          }

          headerLogo.__welcomeScreenClickBound = true;

          headerLogo.onclick = function () {
              showOnlyOfficeWelcomeScreen();

              log('user clicked header logo → show welcome screen');
          };

          return true;
      }

      function getEditTooltipText() {
          if (editModeTransition === 'opening') {
              return 'Opening Edit Mode...';
          }

          if (editModeTransition === 'exiting') {
              return 'Exiting Edit Mode...';
          }

          if (isLargeFile) {
              return 'Documents over 10 MB cannot currently be edited';
          }

          if (!canEdit) {
              return 'You need the Editor + role or higher to edit this document';
          }

          if (isDesktopClosing) {
              return 'Cancel closing the Main App to start editing';
          }

          if (isDesktopLoggingOut) {
              return 'Cancel logging out of the Main App to start editing';
          }

          return 'Only one Member can edit at a time';
      }

      function ensureEditTooltip(doc) {
          if (headerEditTooltip && headerEditTooltip.ownerDocument === doc) {
              headerEditTooltip.textContent = getEditTooltipText();

              return headerEditTooltip;
          }

          headerEditTooltip = doc.createElement('div');
          headerEditTooltip.className = 'sk-tooltip';
          headerEditTooltip.textContent = getEditTooltipText();

          if (doc.body) {
              doc.body.appendChild(headerEditTooltip);
          }

          return headerEditTooltip;
      }

      function bindOnlyOfficeWelcomeScreen() {
          var modal = document.getElementById('welcome-screen');
          var closeButton = document.getElementById('ws-close-btn');

          if (!modal || !closeButton) {
              log('bindOnlyOfficeWelcomeScreen: modal or close button not found');
              return;
          }

          if (modal.__onlyOfficeWelcomeScreenBound) {
              return;
          }

          modal.__onlyOfficeWelcomeScreenBound = true;

          closeButton.onclick = function () {
              modal.style.display = 'none';

              log('user closed onlyoffice welcome screen');
          };
      }

      function showOnlyOfficeWelcomeScreen() {
          var modal = document.getElementById('welcome-screen');

          if (!modal) {
              log('showOnlyOfficeWelcomeScreen: modal element not found');
              return;
          }

          modal.style.display = 'flex';

          log('showOnlyOfficeWelcomeScreen');
      }

      function hideEditTooltip() {
          if (!headerEditTooltip) {
              return;
          }

          headerEditTooltip.style.display = 'none';
      }

      function getDownloadTooltipText() {
          if (isDownloadStarting) {
              return 'Starting download...';
          }

          if (isDownloading) {
              return 'Please wait until the current download is finished';
          }

          if (!canDownload) {
              return 'Downloading is disabled in View-Only Mode';
          }

          return 'Download';
      }

      function ensureDownloadTooltip(doc) {
          if (headerDownloadTooltip && headerDownloadTooltip.ownerDocument === doc) {
              headerDownloadTooltip.textContent = getDownloadTooltipText();

              return headerDownloadTooltip;
          }

          headerDownloadTooltip = doc.createElement('div');
          headerDownloadTooltip.className = 'sk-tooltip';
          headerDownloadTooltip.textContent = getDownloadTooltipText();

          if (doc.body) {
              doc.body.appendChild(headerDownloadTooltip);
          }

          return headerDownloadTooltip;
      }

      function hideDownloadTooltip() {
          if (!headerDownloadTooltip) {
              return;
          }

          headerDownloadTooltip.style.display = 'none';
      }

      function getSaveTooltipText() {
          switch (currentSaveState) {
              case 'idle':
                  return 'No unsaved changes';
              case 'dirty':
                  return 'Save';
              case 'saving':
                  return 'Saving…';
              case 'saved':
                  return 'All changes saved';
              case 'error':
                  return 'Couldn’t save — click to retry';
              default:
                  return 'Save';
          }
      }

      function ensureSaveTooltip(doc) {
          if (headerSaveTooltip && headerSaveTooltip.ownerDocument === doc) {
              headerSaveTooltip.textContent = getSaveTooltipText();

              return headerSaveTooltip;
          }

          headerSaveTooltip = doc.createElement('div');
          headerSaveTooltip.className = 'sk-tooltip';
          headerSaveTooltip.textContent = getSaveTooltipText();

          if (doc.body) {
              doc.body.appendChild(headerSaveTooltip);
          }

          return headerSaveTooltip;
      }

      function showTooltip(type, button) {
          var doc = button.ownerDocument;
          var tooltip;

          if (type === 'save') {
              tooltip = ensureSaveTooltip(doc);
          } else if (type === 'edit') {
              tooltip = ensureEditTooltip(doc);
          } else if (type === 'download') {
              tooltip = ensureDownloadTooltip(doc);
          }

          if (!tooltip) {
              return;
          }

          var TOOLTIP_SCREEN_PADDING = 8;
          var TOOLTIP_OFFSET = 6;
          var rect = button.getBoundingClientRect();

          tooltip.style.display = 'block';

          var tooltipRect = tooltip.getBoundingClientRect();
          var left = rect.left + rect.width / 2 - tooltipRect.width / 2;
          var top = rect.bottom + TOOLTIP_OFFSET;
          var maxLeft = doc.documentElement.clientWidth - tooltipRect.width - TOOLTIP_SCREEN_PADDING;

          if (left < TOOLTIP_SCREEN_PADDING) {
              left = TOOLTIP_SCREEN_PADDING;
          }

          if (left > maxLeft) {
              left = maxLeft;
          }

          tooltip.style.left = left + 'px';
          tooltip.style.top = top + 'px';
      }

      function hideSaveTooltip() {
          if (!headerSaveTooltip) {
              return;
          }

          headerSaveTooltip.style.display = 'none';
      }

      function onEditTabClick() {
          if (!headerEditBtn) {
              return;
          }

          if (!canEdit) {
              showViewOnlyModeModal(true);

              return;
          }

          if (lockHolder && !lockHolder.isSelf) {
              showViewerModeModal(true);

              log('user clicked disabled Edit button: lock held by ' + (lockHolder.userName || 'Someone'));
          }
      }

      // Edit-button click — reuses the existing host postMessage flow that the
      // old outer-page #mode-button used (request-edit-mode / mode-changed). The
      // host drives the actual mode flip back via set-mode → handleSetMode, which
      // re-renders the button.
      function onEditButtonClick() {
          if (!pm || !headerEditBtn || headerEditBtn.disabled) {
              return;
          }

          // Refresh is not an edit-right action: it only asks the host to reload the
          // fresh document bytes. Allow it even when canEdit=false, but only once the
          // conflicting editor released the lock.
          if (conflictState) {
              log('user clicked Refresh after conflict → reload-request');
              pm.toHost({ type: 'reload-request' });
              renderRefreshingButton();
              return;
          }

          if (!canEdit) {
              showViewOnlyModeModal(true);

              return;   // role has no edit right — refuse (button shouldn't exist anyway)
          }

          if (isDesktopClosing || isDesktopLoggingOut) {
              showDesktopClosingModal(true);

              return;
          }

          if (currentMode === 'view') {
              log('user clicked Edit → request-edit-mode');
              editModeTransition = 'opening';
              renderEditButton();
              pm.toHost({ type: 'request-edit-mode' });
          } else if (currentMode === 'edit') {
              // Save pending changes BEFORE releasing the edit lock, so leaving edit
              // mode never drops unsaved work. triggerAutosave is a no-op if the doc
              // isn't dirty or a save is already in flight; when it does fire, the
              // diskette reflects saving→saved via the shared state funnel.
              log('user clicked Edit (editing) → save-on-exit + mode-changed: view');
              editModeTransition = 'exiting';
              renderEditButton();
              pm.triggerAutosave();
              pm.toHost({ type: 'mode-changed', mode: 'view' });
          }
      }

    // Diskette click → save now. Only meaningful when there's something to do:
    // 'dirty' (save) or 'error' (retry). 'idle'/'saving'/'saved' are no-ops (the
    // button is also `disabled` in those states, but guard defensively).
    // Delegates to wrapper-postmessage's requestManualSave, which runs the same
    // capture → x2t → `saved` path as autosave and drives the state back via
    // skSetSaveState.
    function onSaveButtonClick() {
      if (!pm || (currentSaveState !== 'dirty' && currentSaveState !== 'error')) {
          return;
      }

      log('user clicked Save (' + currentSaveState + ') → requestManualSave');
      pm.requestManualSave();
    }

    function focusMainAppWindow() {
      // Tab focus MUST run synchronously inside the click turn — browsers do
      // not grant cross-tab focus from async postMessage handlers (so the
      // host calling window.focus() on `focus-request` is a no-op).
      // window.open('', name) is the reliable primitive when the host tab was
      // given window.name before opening the editor (PROTOCOL.md).
      //
      // Never read cross-origin opener properties (name, closed, location …)
      // — the editor (office.origin) and host (app.origin) are different
      // origins; use the agreed window name directly.
      var MAIN_APP_WINDOW_NAME = 'main-app';

        if (!window.opener || window.opener.closed) {
            window.open(window.HOST_ORIGIN, '_blank');

            return;
        }

      try {
        var target = window.open('', MAIN_APP_WINDOW_NAME);

        if (target && target !== window) {
          log('Main App → window.open("", "' + MAIN_APP_WINDOW_NAME + '")');
        }
      } catch (e) {
        log('Main App window.open failed: ' + e.message);
      }
    }

    function skGoToMainApp() {
      log('user clicked Main App');
      focusMainAppWindow();
    }

    window.skGoToMainApp = skGoToMainApp;

      function mountDownloadButton(doc) {
          var existing = doc.getElementById('download-btn');

          if (existing) {
              headerDownloadBtn = existing;

              if (existing.parentNode) {
                  existing.parentNode.onmouseenter = function () {
                      if (headerDownloadBtn) {
                          showTooltip('download', headerDownloadBtn);
                      }
                  };
                  existing.parentNode.onmouseleave = hideDownloadTooltip;
              }

              renderDownloadButton();

              return true;
          }

          var nativeSaveSlot = doc.getElementById('slot-btn-dt-save');
          var parent = nativeSaveSlot && nativeSaveSlot.parentNode;

          if (!parent) {
              return false;
          }

          var slot = doc.createElement('div');
          slot.className = 'download-slot';
          slot.onmouseenter = function () {
              if (headerDownloadBtn) {
                  showTooltip('download', headerDownloadBtn);
              }
          };
          slot.onmouseleave = hideDownloadTooltip;

          var btn = doc.createElement('button');

          btn.id = 'download-btn';
          btn.className = 'download-btn';
          btn.type = 'button';
          btn.disabled = !canDownload;
          btn.innerHTML =
              '<span class="download-btn__icon">' +
              DOWNLOAD_ICON_SVG +
              '</span>';

          btn.onclick = function () {
              if (!pm || !canDownload || isDownloadStarting || isDownloading) {
                  return;
              }

              log('user clicked Download');

              isDownloadStarting = true;
              renderDownloadButton();

              pm.downloadCurrentFile(function () {
                  isDownloadStarting = false;
                  isDownloading = true;
                  renderDownloadButton();
              })
                  .catch(function (e) {
                      log('download failed: ' + (e && e.message ? e.message : e));
                  })
                  .then(function () {
                      isDownloadStarting = false;
                      isDownloading = false;
                      renderDownloadButton();
                  });
          };

          slot.appendChild(btn);

          // Put Download at the end of the quick-access controls
          parent.appendChild(slot);

          headerDownloadBtn = btn;

          renderDownloadButton();

          log('mountHeaderControls: Download button injected into quick-access toolbar');

          return true;
      }

    // Edit button → trailing <li> in the toolbar tab strip
    // (section.tabs > ul[role="tablist"], Mixtbar.js:113), right after "View".
    function mountEditButton(doc) {
      // Always mount; renderEditButton hides it (display:none) when canEdit is
      // false — visibility-based, not DOM removal, so a permissions/iframe
      // timing race can never leave a "visible-but-dead" button.
      var existing = doc.getElementById('sk-edit-btn');

      if (existing) {
          headerEditBtn = existing;

          if (existing.parentNode) {
            existing.parentNode.onclick = onEditTabClick;
            existing.parentNode.onmouseenter = function() {
                if (headerEditBtn && headerEditBtn.classList.contains('is-locked')) {
                    showTooltip('edit', headerEditBtn);
                }
            };
            existing.parentNode.onmouseleave = hideEditTooltip;
          }

          renderEditButton();

          return true;
        }
      var anchor = doc.querySelector('.tabs ul[role="tablist"]') ||
                   doc.querySelector('ul[role="tablist"]');

      if (!anchor) {
          return false;
      }

      var slot = doc.createElement('li');
      slot.className = 'sk-edit-tab';
      slot.onclick = onEditTabClick;
      slot.onmouseenter = function() {
          if (headerEditBtn && headerEditBtn.classList.contains('is-locked')) {
              showTooltip('edit', headerEditBtn);
          }
      };
      slot.onmouseleave = hideEditTooltip;
      var btn = doc.createElement('button');
      btn.id = 'sk-edit-btn';
      btn.className = 'sk-edit-btn';
      btn.type = 'button';
      btn.innerHTML =
        '<span class="sk-edit-btn__icon">' + SK_EDIT_ICON_SVG + '</span>' +
        '<span class="sk-edit-btn__label">Edit</span>';
      btn.onclick = function (e) {
          e.stopPropagation();

          onEditButtonClick();
      };
      slot.appendChild(btn);
      anchor.appendChild(slot);

      headerEditBtn = btn;
      renderEditButton();
      log('mountHeaderControls: Edit button injected into toolbar tab strip');

      return true;
    }

    // "<who> is editing the document…" label → a trailing <li> in the tab strip,
    // immediately right of the Edit button. Visibility + text driven by
    // renderEditingLabel (hidden when the lock is free in view mode). Anchored
    // off the Edit <li>, so it returns false until that exists.
    function mountEditingLabel(doc) {
      // Always mount; renderEditingLabel keeps it hidden when canEdit is false.
      var existing = doc.getElementById('sk-editing-label');

      if (existing) {
          headerEditingLabel = existing;
          renderEditingLabel();
          return true;
      }

      if (!headerEditBtn) {
          return false;                 // mount the Edit button first
      }

      var editLi = headerEditBtn.parentNode;            // the .sk-edit-tab <li>

      if (!editLi || !editLi.parentNode) {
          return false;
      }

      var slot = doc.createElement('li');
      slot.className = 'sk-editing-tab';
      var span = doc.createElement('span');
      span.id = 'sk-editing-label';
      span.className = 'sk-editing-label';
      slot.appendChild(span);
      // Insert right after the Edit <li> so it flows to its right in the tab row.
      editLi.parentNode.insertBefore(slot, editLi.nextSibling);

      headerEditingLabel = span;
      renderEditingLabel();
      log('mountHeaderControls: editing label injected into toolbar tab strip');

      return true;
    }

    // Save (diskette) button → injected in place of the native quick-access
    // Save slot (#slot-btn-dt-save, Header.js:149), which we hide via CSS.
    function mountSaveButton(doc) {
      var existing = doc.getElementById('sk-save-btn');

        if (existing) {
            headerSaveBtn = existing;
            existing.onmouseenter = function() {
                if (headerSaveBtn) {
                    showTooltip('save', headerSaveBtn);
                }
            };
            existing.onmouseleave = hideSaveTooltip;

            renderSaveButton();

            return true;
        }

      var nativeSlot = doc.getElementById('slot-btn-dt-save');

      if (!nativeSlot || !nativeSlot.parentNode) {
          return false;
      }

      var slot = doc.createElement('div');
      slot.className = 'sk-save-slot';
      var btn = doc.createElement('button');
      btn.id = 'sk-save-btn';
      btn.className = 'sk-save-btn sk-save-btn--idle';
      btn.type = 'button';
      btn.innerHTML = '<span class="sk-save-btn__icon">' + SK_SAVE_ICON_SVG + '</span>';
      btn.onclick = onSaveButtonClick;
      btn.onmouseenter = function() {
          if (headerSaveBtn) {
              showTooltip('save', headerSaveBtn);
          }
      };
      btn.onmouseleave = hideSaveTooltip;
      slot.appendChild(btn);
      // Insert where the diskette was (before the now-hidden native slot).
      nativeSlot.parentNode.insertBefore(slot, nativeSlot);

      headerSaveBtn = btn;
      renderSaveButton();
      log('mountHeaderControls: Save button injected into quick-access toolbar');

      return true;
    }

    // Main App button → header-right, immediately before the search slot
    // (#slot-btn-search, Header.js). Skipped in standalone mode (no opener).
    function mountMainAppButton(doc) {
      var existing = doc.getElementById('sk-main-app-btn');

      if (existing) {
        headerMainAppBtn = existing;
        return true;
      }

      var searchSlot = doc.getElementById('slot-btn-search') ||
        doc.querySelector('[data-layout-name="header-search"]');
      var parent = searchSlot && searchSlot.parentNode;

      if (!parent) {
        parent = doc.querySelector('#box-tools') ||
          doc.querySelector('.extra-right') ||
          doc.querySelector('.box-tools');
      }

      if (!parent) {
          return false;
      }

      var slot = doc.createElement('div');
      slot.className = 'sk-main-app-slot';
      var btn = doc.createElement('button');
      btn.id = 'sk-main-app-btn';
      btn.className = 'sk-main-app-btn';
      btn.type = 'button';
      var mainAppTitle = window.SK_DESKTOP_TRANSPORT ? 'Main App' : 'Main Tab';
      btn.innerHTML =
        '<span class="sk-main-app-btn__icon">' + SK_MAIN_APP_ICON_SVG + '</span>' +
        '<span class="sk-main-app-btn__label">' + mainAppTitle + '</span>';

      btn.onclick = function (e) {
        e.stopPropagation();

          if (window.SK_DESKTOP_TRANSPORT) {
              if (pm) {
                  pm.toHost({ type: 'focus' });
              }

              return;
          }

        // Click lands in the OnlyOffice iframe — call the editor root
        // (edit.html) synchronously so user activation reaches window.open().
        var topWin = doc.defaultView && doc.defaultView.top;

        if (topWin && typeof topWin.skGoToMainApp === 'function') {
          topWin.skGoToMainApp();
        }
      };

      slot.appendChild(btn);

      if (searchSlot && searchSlot.parentNode) {
          searchSlot.parentNode.insertBefore(slot, searchSlot);
      } else {
          parent.appendChild(slot);
      }

      headerMainAppBtn = btn;
      log('mountHeaderControls: Main App button injected into header-right');

      return true;
    }

      // Approach B: inject OUR OWN Edit control into the editor's toolbar tab
      // strip (inside the iframe), right after the "View" tab, styled from Figma
      // — instead of the brittle outer-page overlay. Same iframe-DOM reach +
      // <style> injection idiom as hideNativeDropdown. Idempotent. Returns true
      // once the button is in place. The toolbar renders late (after onAppReady),
      // so callers poll on a false return — mirrors the tryCacheApi poller.
      function mountHeaderControls() {
          var iframe = document.querySelector('iframe[name="frameEditor"]');

          if (!iframe || !iframe.contentDocument) {
              return false;
          }
          var doc = iframe.contentDocument;

          // Inject (or refresh) our styles whenever the iframe doc is reachable.
          injectHeaderControlStyles(doc);
          bindBlockedEditAttemptListeners();

          // Mount both controls; each is idempotent and anchored independently
          // (tab strip vs quick-access toolbar render at different times), so we
          // only report "done" once BOTH are in place — the poller keeps trying
          // until then.

          var shouldMountMainAppButton = window.SK_DESKTOP_TRANSPORT || hasOpener;
          var undoButton = doc.getElementById('slot-btn-dt-undo');
          var redoButton = doc.getElementById('slot-btn-dt-redo');
          var slideshowButton = doc.getElementById('slot-btn-dt-start-over');
          var slideshowInnerButton = slideshowButton.querySelector('button');

          if (pm.isExternal) {
              undoButton.style.display = 'none';
              redoButton.style.display = 'none';
              slideshowButton.style.display = 'none';

              return mountDownloadButton(doc) && (!shouldMountMainAppButton || mountMainAppButton(doc));
          }

          undoButton.querySelector('button').innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
              '<path d="M8.31934 4.38822C8.57315 4.13655 8.98338 4.1385 9.23535 4.39213C9.48698 4.64595 9.48506 5.05619 9.23145 5.30815L6.42188 8.09916H13.9492C17.1801 8.09916 19.7998 10.7189 19.7998 13.9497C19.7997 17.1805 17.18 19.7994 13.9492 19.7994H12C11.6412 19.7994 11.3499 19.5087 11.3496 19.1499C11.3496 18.791 11.641 18.4996 12 18.4996H13.9492C16.462 18.4996 18.4999 16.4625 18.5 13.9497C18.5 11.4369 16.4621 9.39994 13.9492 9.39994H6.42578L9.23242 12.189C9.48577 12.4412 9.4864 12.8514 9.23438 13.105C8.98219 13.3584 8.57197 13.3599 8.31836 13.108L4.42773 9.24076C4.28896 9.12147 4.19924 8.94674 4.19922 8.74955C4.19922 8.56487 4.27661 8.39826 4.40039 8.27983L8.31934 4.38822Z" fill="currentColor"/>' +
              '</svg>';
          redoButton.querySelector('button').innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
              '<path d="M14.7637 4.3921C15.0156 4.13849 15.4259 4.13658 15.6797 4.38819L19.5977 8.2798C19.7218 8.39827 19.7998 8.56453 19.7998 8.74952C19.7998 8.94703 19.7095 9.12143 19.5703 9.24073L15.6807 13.1079C15.427 13.3598 15.0168 13.3583 14.7646 13.105C14.5127 12.8514 14.5134 12.4412 14.7666 12.189L17.5732 9.39991H10.0498C7.537 9.40002 5.49902 11.4369 5.49902 13.9497C5.49917 16.4624 7.53709 18.4994 10.0498 18.4995H11.999C12.358 18.4995 12.6494 18.7909 12.6494 19.1499C12.6492 19.5087 12.3579 19.7993 11.999 19.7993H10.0498C6.81912 19.7992 4.19937 17.1804 4.19922 13.9497C4.19922 10.7189 6.81903 8.09924 10.0498 8.09913H17.5771L14.7676 5.30812C14.5141 5.05618 14.5122 4.64591 14.7637 4.3921Z" fill="currentColor"/>' +
              '</svg>';
          if (slideshowInnerButton) {
              slideshowInnerButton.innerHTML = '<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">' +
                  '<path d="M17.1494 5.5C18.613 5.5 19.7998 6.68684 19.7998 8.15039V15.8506C19.7997 17.3141 18.6129 18.5 17.1494 18.5H6.84961C5.38612 18.5 4.19932 17.314 4.19922 15.8506V8.15039C4.19922 6.68684 5.38606 5.5 6.84961 5.5H17.1494ZM6.84961 6.80078C6.10402 6.80078 5.5 7.40481 5.5 8.15039V15.8506C5.50011 16.5961 6.10409 17.2002 6.84961 17.2002H17.1494C17.8949 17.2002 18.4999 16.5961 18.5 15.8506V8.15039C18.5 7.4048 17.895 6.80078 17.1494 6.80078H6.84961Z" fill="currentColor"/>' +
                  '<path d="M14.6581 11.5708C14.9838 11.7647 14.9838 12.2363 14.6581 12.4301L10.6756 14.8006C10.3423 14.999 9.91988 14.7589 9.91988 14.371L9.91988 9.62998C9.91988 9.24212 10.3423 9.00195 10.6756 9.20033L14.6581 11.5708Z" fill="#0F8CC9"/>' +
                  '</svg>';
          }
          var editDone = mountEditButton(doc);
          mountEditingLabel(doc);              // anchored off the Edit <li>; no-op if edit not mounted yet
          var saveDone = mountSaveButton(doc);
          var downloadDone = mountDownloadButton(doc);
          var mainAppDone = !shouldMountMainAppButton || mountMainAppButton(doc);

          return editDone && saveDone && mainAppDone && downloadDone;
      }

    function autoLoadFixture() {
      log('standalone mode — auto-loading ' + fixtureUrl);
      fetch(fixtureUrl).then(function (r) {
        if (!r.ok) {
            throw new Error('fixture fetch ' + r.status);
        }

        return r.arrayBuffer();
      }).then(function (buf) {
        if (!window.X2TBridge) {
            throw new Error('X2TBridge not available');
        }

        return window.X2TBridge.convertToBin(new Uint8Array(buf), 'sample.' + fixtureExt);
      }).then(function (bin) {
        log('fixture converted, bin=' + bin.length + ' bytes — calling editor.openDocument');
        // Register x2t-extracted media into the editor iframe before
        // openDocument so image references resolve. Same logic as
        // wrapper-postmessage.js — duplicated rather than shared because
        // standalone mode skips the postmessage path entirely.
        if (pm && typeof pm.registerExtractedMedia === 'function') {
          pm.registerExtractedMedia();
        }
        editorInstance.openDocument(bin);
      }).catch(function (err) {
        log('auto-load failed: ' + err.message);
      });
    }

    // ── Construct the DocsAPI editor (called exactly once at boot) ──────
    // No mode parameter — we always build with edit mode + edit permissions;
    // view-only is achieved later via asc_setRestriction.
    function constructEditor() {
      log('constructing editor type=' + type);
      var built = window.buildEditorConfig(type, window.__skHostUser || {});

      var docsApiConfig = {
        documentType: type,           // 'word' | 'cell' | 'slide'
        width:  '100%',
        height: '100%',
        type:   'desktop',
        // Document is omitted intentionally: the editor sits at splash until
        // the host sends `{type:'load', bytes, fileName}` over postMessage,
        // at which point WrapperPostMessage drives the load via DocsAPI's
        // `editor.openDocument(buffer)` method.
        editorConfig: built.editorConfig,
        events: events
      };

      editorInstance = new window.DocsAPI.DocEditor('placeholder', docsApiConfig);
      window.editor  = editorInstance;
      updateOverlayUI();
    }

    // ── Cache the iframe-internal api ref so we can hot-toggle restrictions
    // without destroying the editor. Returns true on success. Safe to call
    // repeatedly — internal state shapes up asynchronously after onAppReady.
    function cacheEditorApi() {
      try {
        var iframe = document.querySelector('iframe[name="frameEditor"]');

        if (!iframe || !iframe.contentWindow) {
          log('cacheEditorApi: iframe not yet present');
          return false;
        }

        var w = iframe.contentWindow;
        var ns = w.DE || w.SSE || w.PE;

        if (!ns || typeof ns.getController !== 'function') {
          log('cacheEditorApi: editor namespace not yet exposed');
          return false;
        }

        var vp = ns.getController('Viewport');

        if (!vp || typeof vp.getApi !== 'function') {
          log('cacheEditorApi: Viewport controller not yet available');
          return false;
        }

        editorApi   = vp.getApi();
        editorApiNs = w.Asc;
        log('cacheEditorApi: editor api cached');
        installMacControlClickContextMenu(w);
        // Hook the user-initiated restriction-change callback so dropdown
        // clicks (Editing↔Viewing) get bridged to the host.
        if (editorApi && typeof editorApi.asc_registerCallback === 'function') {
          editorApi.asc_registerCallback('asc_onChangeRestrictions', onRestrictionsChanged);
        }
        // ── Word-only: listen for the iframe-internal `document:ready` NC
        // notification. Word's Main controller fires this from INSIDE a
        // setInterval(50ms)-polled block (controller/Main.js ~L1518) that
        // also runs `toolbarController.createDelayedElements()`,
        // `activateControls()`, and `api.UpdateInterfaceState()`. That whole
        // block is async — it fires AFTER `Common.Gateway.documentReady()`
        // (which is what reaches our DocsAPI `onDocumentReady` event). So our
        // `onDocumentReady`-time `editing:disable` can land BEFORE word's
        // toolbar is fully wired (controllers attached, state activated),
        // and some piece of the view-mode chrome (mask/dropdown/header) ends
        // up not painted on initial .docx boot.
        //
        // Cell/slide have the same setInterval pattern but happen not to need
        // this — observed empirically that their initial-boot mask sticks on
        // the first `editing:disable`. Word does not; hence the word-only
        // gate. Re-applying the restriction inside `document:ready` (which
        // fires AFTER all the polled setup) lets DisableToolbar run against
        // a fully-prepared toolbar and the mask is painted correctly. The
        // user-initiated dropdown toggle has always worked because by then
        // the polled block long since finished.
        try {
          var nc = w.Common && w.Common.NotificationCenter;

          if (type === 'word' && nc && typeof nc.on === 'function' && !nc.__wrapperDocReadyBound) {
            nc.__wrapperDocReadyBound = true;
            nc.on('document:ready', function () {
              log('iframe NC document:ready (word) — re-applying full restriction post-polled-setup');
              // Re-run the full apply with whatever the latest pendingRestrict is.
              // applyRestriction is idempotent for the disable=true→true case
              // (DisableToolbar's early-return on existing mask makes the editing:disable
              // a no-op visually, but it refreshes the stackDisableActions entry).
              applyRestriction(pendingRestrict);
              initialRestrictionApplied = true;
            });
          }
        } catch (e) {
          log('cacheEditorApi: document:ready bind error: ' + (e && e.message ? e.message : e));
        }

        return true;
      } catch (e) {
        log('cacheEditorApi error:', e && e.message ? e.message : e);

        return false;
      }
    }

    // ── Flag objects for editing:disable dispatched on cell/slide ──────
    // Modeled on the editor controllers' `disableEditing` methods
    // (disconnect path), but constructed *dynamically* per call because
    // some sub-handlers ignore the outer `disable` arg and key off the
    // flag-object instead. Most notably, Toolbar.DisableToolbar reads
    // `options.viewMode` to set `this.editMode = !viewMode`. With a
    // statically-hardcoded `viewMode: true`, the re-enable call leaves
    // `editMode=false`, and the function then computes
    // `disable = false || !editMode = true` → the toolbar mask stays
    // forever. The fix: have `viewMode` (and the conceptually similar
    // `clear` sub-flags) mirror the `disable` arg, just like the editor's
    // own `disableEditing` does (`viewMode: disable`).
    //   CELL  ← spreadsheeteditor/main/app/controller/Main.js (~L1313,
    //           Toolbar.DisableToolbar at Toolbar.js:4896 is the culprit)
    //   SLIDE ← presentationeditor/main/app/controller/Main.js (~L1223,
    //           Toolbar.DisableToolbar at Toolbar.js:2722 is the culprit)
    // Caveat: a few flags (e.g. `chat`) come from the disconnect path
    // and are irrelevant to a plain view/edit toggle but harmless.
    function buildCellDisableFlags(disable) {
      return {
        viewMode:        disable,                            // ← was hardcoded true; must mirror `disable`
        allowSignature:  false,
        allowProtect:    false,
        rightMenu:       { clear: disable, disable: true },  // clear selection only on disable side
        statusBar:       true,
        leftMenu:        { disable: true, previewMode: true },
        fileMenu:        { protect: true, history: false },
        comments:        { disable: true, previewMode: true },
        chat:            true,
        review:          true,
        viewport:        true,
        documentHolder:  { clear: disable, disable: true },
        toolbar:         true,
        celleditor:      { previewMode: true },
        header:          { search: false },
        shortcuts:       false
      };
    }

    function buildSlideDisableFlags(disable) {
      return {
        viewMode:        disable,                            // ← was hardcoded true; must mirror `disable`
        allowSignature:  false,
        rightMenu:       { clear: disable, disable: true },
        statusBar:       true,
        leftMenu:        { disable: true, previewMode: true },
        fileMenu:        { protect: true, history: false },
        comments:        { disable: true, previewMode: true },
        chat:            true,
        review:          true,
        viewport:        true,
        documentHolder:  { clear: disable, disable: true },
        toolbar:         true,
        header:          { search: false },
        shortcuts:       false,
        documentPreview: { draw: false }
      };
    }

      function ensurePresentationEditContextMenus(iframeWin) {
          if (!iframeWin.PE || typeof iframeWin.PE.getController !== 'function') {
              return;
          }

          var documentHolderController = iframeWin.PE.getController('DocumentHolder');
          var documentHolderView = documentHolderController && documentHolderController.getView();

          if (!documentHolderView || typeof documentHolderView.createDelayedElements !== 'function') {
              return;
          }

          documentHolderView.createDelayedElements();

          log('presentation edit context menus initialized');
      }

      function installMacControlClickContextMenu(iframeWin) {
          var isUnsupportedEditorType = type === 'cell';
          var isMacOS = !!(
              iframeWin &&
              iframeWin.navigator &&
              iframeWin.navigator.platform.indexOf('Mac') === 0
          );

          if (isUnsupportedEditorType || !isMacOS) {
              return;
          }

          var ascCommon = iframeWin.AscCommon;
          var hasMouseEventHandlers = !!(
              ascCommon &&
              typeof ascCommon.check_MouseDownEvent === 'function' &&
              typeof ascCommon.check_MouseUpEvent === 'function'
          );

          if (!hasMouseEventHandlers) {
              return;
          }

          if (ascCommon.__customControlClickContextMenuInstalled) {
              return;
          }

          ascCommon.__customControlClickContextMenuInstalled = true;

          var originalCheckMouseDownEvent = ascCommon.check_MouseDownEvent;
          var originalCheckMouseUpEvent = ascCommon.check_MouseUpEvent;

          function normalizeControlClick(e) {
              var isControlClick = !!(
                  e &&
                  e.ctrlKey &&
                  e.button === 0 &&
                  ascCommon.global_mouseEvent
              );

              if (!isControlClick) {
                  return;
              }

              ascCommon.global_mouseEvent.Button = 2;
              ascCommon.global_mouseEvent.CtrlKey = false;
              ascCommon.global_mouseEvent.ctrlKey = false;
          }

          ascCommon.check_MouseDownEvent = function () {
              var result = originalCheckMouseDownEvent.apply(this, arguments);

              normalizeControlClick(arguments[0]);

              return result;
          };

          ascCommon.check_MouseUpEvent = function () {
              var result = originalCheckMouseUpEvent.apply(this, arguments);

              normalizeControlClick(arguments[0]);

              return result;
          };

          log('macOS Control + click context menu normalization installed');
      }

    // Modelled on word's `disableEditing` (documenteditor/main/app/controller/
    // Main.js ~L865), with the same dynamic-per-call rationale as the cell/
    // slide builders: `viewMode` and `clear` sub-flags MUST mirror `disable`
    // so the Toolbar/RightMenu/DocumentHolder masks lift correctly on the
    // re-enable side. We omit `plugins` and `protect` (both `false` in word's
    // disconnect path, so onEditingDisable's `if (options.plugins)` /
    // `if (options.protect)` branches no-op anyway — and we don't ship those
    // controllers). `temp` (the reconnect/refresh-file branch in word) is
    // collapsed to its default-false value because a view/edit toggle is
    // neither a reconnect nor a refresh.
    function buildWordDisableFlags(disable) {
      return {
        viewMode:        disable,                            // ← must mirror `disable`
        reviewMode:      false,
        fillFormMode:    false,
        viewDocMode:     false,
        allowMerge:      false,
        allowSignature:  false,
        allowProtect:    false,
        rightMenu:       { clear: disable, disable: true },  // clear selection only on disable side
        statusBar:       true,
        leftMenu:        { disable: true, previewMode: true },
        fileMenu:        { protect: true, history: false },
        navigation:      { disable: true, previewMode: true },
        comments:        { disable: true, previewMode: true },
        chat:            true,
        review:          true,
        viewport:        true,
        documentHolder:  { clear: disable, disable: true },
        toolbar:         true,
        plugins:         false,
        protect:         false,
        header:          { docmode: true, search: false, startfill: false },
        shortcuts:       false
      };
    }

    // ── Hot-apply the given mode via the proper per-editor API path. ────
    // All three editor types (word/cell/slide) use the same shape: fire
    // `editing:disable` with a type-specific flag-object, then call
    // asc_setRestriction directly. The stack-key (third positional arg to
    // editing:disable) MUST match between disable and re-enable; we use
    // 'view' consistently. If the api isn't cached yet (we're in the
    // window between mount and onAppReady), queue the request and the
    // tryCacheApi poll will apply it once ready.
    //
    // Word originally went through the docmode-apply NotificationCenter
    // handler instead, which worked for the user-toggle case but not the
    // initial-boot case (the .docx briefly showed edit-mode chrome before
    // the mask appeared). Switching to the editing:disable + asc_setRestriction
    // path makes initial boot uniform with cell/slide, which were already
    // robust on both code paths.
    function applyRestriction(mode) {
      // mode is 'view' | 'edit'
      // Keep pendingRestrict in sync with the latest requested mode so that
      // the deferred phase-2 application in onDocumentReady picks up any
      // host-driven set-mode that arrived during the appReady→documentReady
      // window.
      pendingRestrict = mode;

      if (!editorApi || !editorApiNs) {
        log('applyRestriction: api not yet cached, queuing ' + mode);

        return;
      }

      var iframe = document.querySelector('iframe[name="frameEditor"]');

      if (!iframe || !iframe.contentWindow) {
        log('applyRestriction: iframe gone, queuing ' + mode);
        pendingRestrict = mode;

        return;
      }

      var iframeWin = iframe.contentWindow;
      var nc        = iframeWin.Common && iframeWin.Common.NotificationCenter;

      if (!nc || typeof nc.trigger !== 'function') {
        log('applyRestriction: NotificationCenter not ready, falling back to bare asc_setRestriction');
        bareSetRestriction(mode);

        return;
      }

      var R = editorApiNs.c_oAscRestrictionType;
      var disable = (mode === 'view');

      // Suppress the asc_onChangeRestrictions callback that will fire as a
      // side-effect of asc_setRestriction. We know the target restriction
      // value and stash it; the callback compares and bails on match.
      lastAppliedRestriction = disable ? R.View : R.None;

      try {
        if (type === 'word') {
          // Build flags per-call so `viewMode` (and `clear` sub-flags) track
          // `disable`. Required for Toolbar.DisableToolbar to actually un-mask
          // on re-enable. See buildWordDisableFlags comment above.
          nc.trigger('editing:disable', disable, buildWordDisableFlags(disable), 'view');
          // Mirror onDocModeApply: when entering view mode (disable=true), the
          // native dropdown also fires `reviewchanges:turn` so the review-mode
          // chrome aligns with view state; on re-enable (disable=false) it
          // would have been the 'edit' branch which fires `reviewchanges:turn`
          // false. Without this, word's review controller can hold stale
          // review-mode state across our toggle and leave subtle toolbar
          // chrome in the wrong place. Safe to fire for all transitions —
          // it only affects review-related buttons.
          nc.trigger('reviewchanges:turn', false);
          editorApi.asc_setRestriction(disable ? R.View : R.None);

          if (!disable) {
              var documentHolderController = iframeWin.DE && iframeWin.DE.getController('DocumentHolder');
              var documentHolderView = documentHolderController && documentHolderController.getView();
              var shouldCreateDelayedElements = documentHolderView &&
                  !documentHolderView.tableMenu &&
                  typeof documentHolderView.createDelayedElements === 'function';

              if (shouldCreateDelayedElements) {
                  documentHolderView.createDelayedElements();
              }
            }

            nc.trigger('doc:mode-changed', mode);
          log('applyRestriction(word): editing:disable ' + disable + ' + reviewchanges:turn false + asc_setRestriction(' + (disable ? 'View' : 'None') + ')');
        } else if (type === 'cell') {
          // Build flags per-call so `viewMode` (and `clear` sub-flags) track
          // `disable`. Required for Toolbar.DisableToolbar to actually un-mask
          // on re-enable — it ignores its `disable` arg when viewMode forces
          // editMode=false. See buildCellDisableFlags comment above.
          nc.trigger('editing:disable', disable, buildCellDisableFlags(disable), 'view');
          editorApi.asc_setRestriction(disable ? R.View : R.None);
          log('applyRestriction(cell): editing:disable ' + disable + ' + asc_setRestriction(' + (disable ? 'View' : 'None') + ')');
        } else if (type === 'slide') {
            nc.trigger('editing:disable', disable, buildSlideDisableFlags(disable), 'view');
            editorApi.asc_setRestriction(disable ? R.View : R.None);

            if (!disable) {
                ensurePresentationEditContextMenus(iframeWin);
            }

            log('applyRestriction(slide): editing:disable ' + disable + ' + asc_setRestriction(' + (disable ? 'View' : 'None') + ')');
        } else {
          // Unknown type — fall back to the bare path.
          log('applyRestriction: unknown editor type "' + type + '" — falling back to bare asc_setRestriction');
          bareSetRestriction(mode);
        }
      } catch (e) {
        log('applyRestriction error: ' + (e && e.message ? e.message : e));
        // On any failure, fall back to the bare path so editing-block is at least correct
        bareSetRestriction(mode);
      }
    }

    function bareSetRestriction(mode) {
      if (!editorApi || !editorApiNs) {
          return;
      }

      var R = editorApiNs.c_oAscRestrictionType;
      var target = (mode === 'view') ? R.View : R.None;
      lastAppliedRestriction = target;

      try {
        editorApi.asc_setRestriction(target);
      } catch (e) {
        log('bareSetRestriction error: ' + (e && e.message ? e.message : e));
      }
    }

    // ── asc_onChangeRestrictions handler ────────────────────────────────
    // Two sources fire this callback:
    //   1. Our own programmatic asc_setRestriction (from bareSetRestriction /
    //      applyRestriction). Matched by lastAppliedRestriction; clear + bail.
    //   2. SDK-internal asc_setRestriction. For word specifically, the
    //      DocProtection controller's onAppReady (web-apps/.../controller/
    //      DocProtection.js:175) fires a Promise.resolve().then microtask
    //      after the `app:ready` NC notification, and that microtask calls
    //      applyRestrictions(None) → asc_setRestriction(R.None) for any
    //      unprotected .docx (DocProtection.js:301). Cell/slide have no
    //      equivalent path (WBProtection.onAppReady doesn't touch
    //      restrictions), which is why this bug was word-specific.
    //
    // We can treat all unmatched callbacks as SDK-internal intrusions
    // because the native Editing/Viewing dropdown — the only user-facing
    // affordance that would call asc_setRestriction directly — is hidden
    // permanently via CSS injection (hideNativeDropdown). The user's only
    // mode-toggle path is our injected header Edit button, which routes through
    // the host postMessage protocol (request-edit-mode / mode-changed) and
    // bypasses asc_setRestriction entirely. So: silently re-assert our
    // intended restriction; never post mode-changed to the host (that path
    // is what was causing the initial-boot edit-chrome flash on .docx —
    // DocProtection's R.None looked like a user dropdown click, the
    // harness echoed set-mode: edit, the wrapper applied edit, and the
    // view-mode mask we'd just painted was torn back down).
    function onRestrictionsChanged(r) {
      if (r === lastAppliedRestriction) {
        lastAppliedRestriction = null;

        return;
      }

      if (!editorApi || !editorApiNs) {
          return;
      }

      log('asc_onChangeRestrictions: SDK-internal change (raw=' + r + ') — re-asserting ' + pendingRestrict);
      bareSetRestriction(pendingRestrict);
    }

    // ── Mode change handler — dispatched from wrapper-postmessage.js ────
    // No more destroy+reconstruct. We just call asc_setRestriction.
    // Scroll position, cursor, undo history all survive the toggle.
    function handleSetMode(newMode, newLockHolder) {
      if (newMode !== 'view' && newMode !== 'edit') {
        log('handleSetMode: invalid mode "' + newMode + '" — ignoring');

        return;
      }

      editModeTransition = null;
      // NOTE: we intentionally do NOT gate edit on canEdit here. set-mode is only
      // sent by the origin-pinned main app AFTER a rights-checked acquireEditLock,
      // so a set-mode:edit is authoritative; gating it here risked blocking a
      // legitimate editor if canEdit were briefly stale. Viewer UX is handled by
      // hiding the Edit button (renderEditButton); security is enforced server-side
      // (acquireEditLock + appendDiffChunk EDIT_CONTENT).
      lockHolder = newLockHolder || null;
      // Remember the OTHER user's name while we have it, so a conflict that
      // arrives right after the lock is released (cleared above) can still name
      // who edited. Not cleared on release — it's always overwritten by the next
      // real holder before another conflict can occur.
      if (newLockHolder && newLockHolder.userName) {
          lastLockHolderName = newLockHolder.isSelf ? 'You' : newLockHolder.userName;
      }
      // Remember the OTHER user's id too (not our own — isSelf never needs a
      // lookup). Survives the lock release the same way lastLockHolderName does,
      // so a post-release conflict can still ask the host to name who edited by
      // id — even after the host's editLock (and its reactive watcher) forgot them.
      if (newLockHolder && newLockHolder.userId) {
          lastLockHolderId = newLockHolder.userId;
      }
      // Live "Someone is editing…" (lock held, name unresolved) — ask the host to
      // resolve it by id too, same path as the conflict banner below.
      if (newLockHolder && (!newLockHolder.userName || newLockHolder.userName === 'Someone')) {
          maybeRequestEditorName('Someone', newLockHolder.userId);
      }

      currentMode = newMode;
      updateOverlayUI();
      applyRestriction(newMode);
      log('handleSetMode: applied ' + newMode +
          (lockHolder ? ' (lockHolder=' + lockHolder.userName + ')' : ''));
    }
    window.handleSetMode = handleSetMode;

    // ── Conflict handler — dispatched from wrapper-postmessage.js ───────────
    // The conflict payload itself does not carry `updatedBy`. The user who saved
    // the newer version is the edit-lock holder, so we name them from
    // `lockHolder.userName`. BUT a save-and-exit releases the lock in the same
    // beat as its save: the set-mode:view that clears `lockHolder` can be
    // processed just before this conflict, so we fall back to the last holder
    // name we remembered (lastLockHolderName) before finally landing on "Someone".
      function handleConflict(userId, userName) {
          conflictState = {
              updatedBy: userName || 'Someone',
              userId: userId || null
          };

          updateOverlayUI();

          log('handleConflict: document updated by ' + conflictState.updatedBy);

          maybeRequestEditorName(conflictState.updatedBy, conflictState.userId);
      }
    window.handleConflict = handleConflict;

    function maybeRequestEditorName(displayName, userId) {
      if (displayName !== 'Someone' || !pm) {
          return;
      }

      if (editorNameRequestedFor === userId) {
          return;
      }

      editorNameRequestedFor = userId;
      log('holder name unresolved → request-editor-name for ' + userId);
      pm.toHost({ type: 'request-editor-name', userId: userId });
    }

    function handleEditorName(userId, userName) {
      log('Handle editor name: ');
      log(userId);
      log(userName);

      if (!userId || !userName || userName === 'Someone') {
          return;
      }

      let changed = false;

      if (lastLockHolderId === userId) {
          lastLockHolderName = userName;
      }

      if (conflictState && conflictState.userId === userId && conflictState.updatedBy === 'Someone') {
        conflictState.updatedBy = userName;
        changed = true;
      }

      if (lockHolder && lockHolder.userId === userId && lockHolder.userName === 'Someone') {
        lockHolder.userName = userName;
        changed = true;
      }

      if (changed) {
        updateOverlayUI();
        log('handleEditorName: resolved ' + userId + ' → ' + userName);
      }
    }
    window.handleEditorName = handleEditorName;

    // ── Conflict-clear handler — dispatched from wrapper-postmessage.js ─────
    // Reserved for the future flow where the host decides the conflict is no
    // longer relevant without forcing a full page reload. We restore the normal
    // header state: Refresh turns back into Edit, and the "Updated by <user>"
    // label goes back to the usual editing-status label or hides completely.
    function handleConflictCleared() {
      conflictState = null;
      editorNameRequestedFor = null;   // allow a fresh lookup on the next conflict

      updateOverlayUI();

      log('handleConflictCleared: conflict state cleared');
    }
    window.handleConflictCleared = handleConflictCleared;

    // Role-gated edit capability — dispatched from wrapper-postmessage.js on the
    // `permissions` postMessage (the main app derives canEdit from the user's
    // Role / EDIT_CONTENT right). Sent before `load`, so it lands before the
    // toolbar tab-strip exists → no Edit-button flash for viewers.
    function handlePermissions(perms) {
        var nextCanEdit = !(perms && perms.canEdit === false);
        var nextCanDownload = !(perms && perms.canDownload === false);
        var nextIsLargeFile = !(perms && perms.isLargeFile === false);
        var nextDesktopClosing = !!(perms && perms.isDesktopClosing);
        var nextDesktopLoggingOut = !!(perms && perms.isDesktopLoggingOut);

        if (nextCanEdit === canEdit && nextDesktopClosing === isDesktopClosing && nextDesktopLoggingOut === isDesktopLoggingOut && nextCanDownload === canDownload && nextIsLargeFile === isLargeFile) {
            return;
        }

        canEdit = nextCanEdit;
        isDesktopClosing = nextDesktopClosing;
        isDesktopLoggingOut = nextDesktopLoggingOut;
        canDownload = nextCanDownload;
        isLargeFile = nextIsLargeFile;
        log('handlePermissions: canEdit = ' + canEdit);
        log('handlePermissions: isDesktopClosing = ' + isDesktopClosing);
        log('handlePermissions: isDesktopLoggingOut = ' + isDesktopLoggingOut);
        log('handlePermissions: canDownload = ' + canDownload);
        log('handlePermissions: isLargeFile = ' + isLargeFile);
        // Re-render the Edit button + editing label to reflect the new capability
        // (edit button hides when pm.isExternal === true).
        // The controls are mounted pm.isExternal === false, otherwise only download button.
        updateOverlayUI();
    }
    window.handlePermissions = handlePermissions;

    // ── Events block (closed over by constructEditor) ──────────────────
    events = {
      onAppReady: function () {
        log('onAppReady — initialising postmessage bridge mode=' + currentMode);
        pm = new window.WrapperPostMessage({ editor: editorInstance, editorType: type });
        pm.signalReady();

        if (isStandalone) {
            autoLoadFixture();
        }

        updateOverlayUI();
        bindTurnOnEditModeModal();
        bindDesktopClosingModal();
        bindViewerModeModal();
        bindViewOnlyModeModal();
        bindOnlyOfficeWelcomeScreen();
        bindSaveShortcutListeners();
        bindBlockedContentCopyListeners(document);

        // Cache the iframe's internal api so we can hot-toggle restriction
        // without destroy/reconstruct. The Viewport controller may not be
        // wired up the instant onAppReady fires (especially for cell/slide),
        // so we try immediately and poll briefly if it isn't ready yet.
        var attempts = 0;
        function tryCacheApi() {
          if (cacheEditorApi()) {
            // ── Two-phase initial mode application (all editor types) ────
            // Phase 1 (here, at onAppReady): apply ONLY the bare
            // asc_setRestriction so editing is blocked from the instant the
            // document opens. We deliberately skip the full per-type
            // editing:disable dispatch because:
            //   • SSE/PE Toolbar controller's setApi runs inside
            //     onDocumentContentReady (Main.js ~L1717 SSE / ~L1567 PE),
            //     which fires AFTER Common.Gateway.appReady().
            //   • The toolbar view's DOM (the `.toolbar` element that
            //     DisableToolbar appends the `.toolbar-mask` child to) is
            //     not rendered until `app:face` triggers `toolbar.render()`
            //     in onAppShowed (~L4943 SSE / ~L2754 PE).
            //   • Triggering editing:disable before the toolbar renders
            //     means the mask gets appended to an empty selector — no
            //     visible mask, even though restriction blocks editing.
            //   • Word has the same problem: dispatching too early on word
            //     also misses the toolbar render, producing the edit-chrome
            //     flash on initial .docx boot. Word is now handled identically
            //     to cell/slide here (full apply deferred to onDocumentReady).
            // Phase 2 (in onDocumentReady): fire the full applyRestriction
            // path so the per-type UI gets fully painted (toolbar mask, etc.).
            bareSetRestriction(pendingRestrict);
            log('onAppReady: phase-1 bare restriction applied (' + pendingRestrict + '); full applyRestriction deferred to onDocumentReady');
            // Permanently hide the native dropdown now that the iframe DOM
            // is confirmed ready. Our injected header Edit button is the
            // affordance (mountHeaderControls).
            hideNativeDropdown();
            return;
          }
          if (++attempts < 20) {
              setTimeout(tryCacheApi, 50);   // up to ~1 s of polling
          } else {
              log('tryCacheApi: gave up after ' + attempts + ' attempts');
          }
        }
        tryCacheApi();

        // Inject our Edit button into the header (approach B). The header
        // renders late and independently of the Viewport api, so poll on its
        // own clock (~3 s). Idempotent, so a later onDocumentReady call is safe.
        var mountAttempts = 0;

        (function tryMountHeader() {
          if (mountHeaderControls()) {
              return;
          }

          if (++mountAttempts < 60) {
              setTimeout(tryMountHeader, 50);
          } else {
              log('tryMountHeader: header anchor never appeared');
          }
        })();
      },
      onDocumentReady: function () {
        log('onDocumentReady');
        // Phase 2 of the initial-mode setup — uniform across word/cell/slide.
        // By the time the SDK fires onDocumentReady, the editor's Main
        // controller has finished onDocumentContentReady — toolbar.setApi
        // has run, app:face has triggered toolbar.render(), and the
        // `.toolbar` DOM exists to accept the mask child. Now the full
        // applyRestriction dispatch (editing:disable + asc_setRestriction,
        // with the type-specific flag-object) actually repaints the toolbar
        // with a usable mask / mode chrome.
        if (!initialRestrictionApplied) {
          // Bounded retry: in case onDocumentReady fires before the
          // iframe's internal NotificationCenter is wired (unlikely here
          // because we're firing AFTER documentReady from the editor's
          // own Main, but defensive). Cap at ~2 s.
          var startedAt = Date.now();
          (function tryFullApply() {
            if (editorApi && editorApiNs) {
              applyRestriction(pendingRestrict);
              initialRestrictionApplied = true;
              log('onDocumentReady: phase-2 full applyRestriction applied (' + pendingRestrict + ')');
              return;
            }
            if (Date.now() - startedAt > 2000) {
              log('onDocumentReady: phase-2 gave up after 2 s — api never cached');
              return;
            }
            setTimeout(tryFullApply, 50);
          })();
        }
        // Header is fully rendered by now — mount our Edit button if the
        // onAppReady poller hasn't already (idempotent).
        mountHeaderControls();
        if (pm) {
          pm.toHost({ type: 'opened', requestId: pm.requestId });
          if (pm.shouldShowOnlyOfficeWelcomeScreen) {
            showOnlyOfficeWelcomeScreen();
            pm.shouldShowOnlyOfficeWelcomeScreen = false;
          }
          // Ask the host for the CURRENT lock/mode state. On a fresh boot this
          // is mostly redundant with the host's reactive push, but after a
          // freshness reload the editor has no in-memory mode and the host's
          // lock watcher won't re-fire (editLock unchanged) — so without this
          // the Edit button would wrongly show enabled while another user is
          // still editing. The host replies with `set-mode` → handleSetMode.
          // Controls are mounted (above), so the response can render.
          pm.toHost({ type: 'request-edit-state' });
        }
      },
      onError: function (e) {
        log('onError', e && e.data);
        if (pm) {
            pm.error('SDK_ERROR', JSON.stringify(e && e.data || {}), pm.requestId);
        }
      },
      onWarning: function (e) {
        log('onWarning', e && e.data);
      },
      onDocumentStateChange: function (e) {
        // DocsAPI surfaces this when the editor's Main controller calls
        // `Common.Gateway.setDocumentModified(bool)`. Word's Main fires it
        // from `onDocumentModifiedChanged`, which the SDK fires via the
        // `asc_onDocumentModifiedChanged` callback whenever the model's
        // modified flag flips (post-edit, post-undo-to-clean, post-save).
        var dirty = !!(e && e.data);
        log('onDocumentStateChange dirty=' + dirty);
        // Track so beforeunload can decide whether to warn.
        window.__editorDirty = dirty;
        // pm.onDirtyChanged relays to host (`dirty` postMessage) AND drives
        // the editor-side autosave debounce.
        if (pm) {
            pm.onDirtyChanged(dirty);
        }
      },
      onRequestClose: function () {
        log('onRequestClose');

        if (pm) {
            pm.toHost({type: 'close-request'});
        }
      }
    };

      function notifyHostAboutPageClosing() {
          if (!pm || window.__skPageClosingNotified) {
              return;
          }

          window.__skPageClosingNotified = true;

          if (window.__editorDirty) {
              pm.triggerAutosave();
          }

          pm.toHost({
              type: 'page-closing',
              dirty: !!window.__editorDirty,
          });
      }

    // Browser-level guard: if the user closes the editor tab with unsaved
    // edits, prompt before discarding. The host main app should ALSO show
    // a confirmation in its own UI on `close-request`, but this is the
    // last line of defence for direct tab-close (Cmd-W, X button) where no
    // host event ever fires.
      window.addEventListener('beforeunload', function (e) {
          if (window.SK_DESKTOP_TRANSPORT && pm) {
              if (!window.__editorDirty) {
                  return;
              }

              e.preventDefault();
              log('desktop beforeunload + dirty → saveAndClose');
              pm.saveAndClose();

              return;
          }

          var cannotReconnectModal = document.getElementById("cannot-reconnect-modal");
          var closedHostModal = document.getElementById("main-app-closed-modal");
          var loggedOutModal = document.getElementById("main-app-logged-out-modal");
          var isNativeCloseConfirmationNeeded = window.__editorDirty &&
              (!cannotReconnectModal || getComputedStyle(cannotReconnectModal).display === 'none') &&
              (!loggedOutModal || getComputedStyle(loggedOutModal).display === 'none') &&
              (!closedHostModal || getComputedStyle(closedHostModal).display === 'none');

          if (isNativeCloseConfirmationNeeded) {
              e.preventDefault();
          }
      });

    // Autosave: also fire when the tab is backgrounded. visibilitychange
    // fires while the page is still alive (unlike beforeunload), so the
    // async x2t + postMessage chain has time to complete. This catches the
    // common "user Cmd-Tab'd away mid-edit" case.
    document.addEventListener('visibilitychange', function () {
      if (document.visibilityState === 'hidden' && window.__editorDirty && pm) {
        log('tab hidden + dirty → fire immediate autosave');
        pm.triggerAutosave();
      }
    });

    window.addEventListener('pagehide', notifyHostAboutPageClosing);

    // The Edit button is injected into the iframe header after onAppReady
    // (mountHeaderControls); its click handler is wired there. Nothing to
    // wire on the outer page anymore.

    // Ensure overlay state is consistent before the editor mounts (the banner
    // is an outer-page element; renderEditButton no-ops until the header
    // button exists).
    updateOverlayUI();

    // ── Boot: construct the editor in edit mode. We immediately queue a
    // pending view restriction so the document opens read-only by default;
    // the main app will send `set-mode: edit` after acquiring the lock to
    // promote it. ───────────────────────────────────────────────────────
    constructEditor();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }
    /* jshint +W003 */
    /* jshint +W106 */
    /* jshint +W104 */
    /* jshint +W119 */
})();
