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
    var headerEditBtn  = null;     // our Edit button injected into the iframe header
                                   // (approach B); null until mountHeaderControls runs
    var headerSaveBtn  = null;     // our Save (diskette) button, same approach
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
          if (currentMode === 'edit' || !canEdit) {
              return false;
          }

          if (lockHolder && !lockHolder.isSelf) {
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
      if ((currentMode === 'edit' && !isDesktopClosing) || !canEdit || !isEditAttemptEvent(e)) {
          return;
      }

        if (isDesktopClosing) {
            e.preventDefault();
            e.stopPropagation();

            if (typeof e.stopImmediatePropagation === 'function') {
                e.stopImmediatePropagation();
            }

            showDesktopClosingModal();

            log('blocked edit attempt: desktop closing');

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

      function showDesktopClosingModal() {
          var modal = document.getElementById('cannot-start-edit-mode');

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
          var firstDescription = modal.querySelector('.vm-dialog-description');

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
              var b = doc.createElement('span');

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
          if (!canEdit || isDesktopClosing) {
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
              var b = doc.createElement('span');

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

          // No edit right ⇒ hide the Edit button entirely (display:none on the <li>).
          // Visibility-based so there's never a "visible-but-dead" button.
          if (!canEdit) {
              if (li) {
                  li.style.display = 'none';
              }

              headerEditBtn.disabled = true;

              return;
          }

          if (isDesktopClosing) {
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

          if (currentMode === 'view' && lockHolder) {
              headerEditBtn.classList.add('is-locked');
              headerEditBtn.disabled = true;
          } else if (currentMode === 'edit') {
              headerEditBtn.classList.add('is-editing');
              headerEditBtn.disabled = false;
          } else {
              headerEditBtn.disabled = false;   // free — base class only
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
      headerSaveBtn.title = ({
        idle:'No unsaved changes', dirty:'Save (unsaved changes)', saving:'Saving…',
        saved:'All changes saved', error:'Couldn’t save — click to retry'
      })[s] || 'Save';
    }

    // Called by wrapper-postmessage.js (via window.skSetSaveState) whenever the
    // save lifecycle advances. Single funnel so the diskette always mirrors the
    // real save state — including saves triggered by edit-mode-off / page-leave.
    window.skSetSaveState = function (state) {
      currentSaveState = state || 'idle';
      renderSaveButton();
    };

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
    var SK_SAVE_ICON_SVG = `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M14.3 5.25H8.2M14.3 5.25V9.2H8.2V5.25M14.3 5.25L14.8 5.25L18.749 9.2V16.75C18.749 17.8546 17.8536 18.75 16.749 18.75H15.8M15.8 18.75H8.2M15.8 18.75V13.15H8.2V18.75M8.2 5.25H7.25098C6.14641 5.25 5.25098 6.14543 5.25098 7.25V16.75C5.25098 17.8546 6.14641 18.75 7.25098 18.75H8.2" stroke="currentColor" stroke-linejoin="round"/>
    </svg>`;
    // saved — diskette + teal check badge (#3FC0C4, matches the Edit button;
    // bottom-right, white halo cuts it out of the diskette).
    var SK_SAVE_ICON_ERROR = `<svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path fill-rule="evenodd" clip-rule="evenodd" d="M14.7998 4.75C14.9324 4.75 15.0596 4.80278 15.1533 4.89648L19.1025 8.84668C19.1963 8.94044 19.249 9.06763 19.249 9.2002V13.25C19.249 13.5261 19.0252 13.75 18.749 13.75C18.4729 13.75 18.249 13.5261 18.249 13.25V9.40625L14.7998 5.95703V9.2002C14.7997 9.47625 14.5759 9.7002 14.2998 9.7002H8.2002C7.92412 9.7002 7.7003 9.47625 7.7002 9.2002V5.75H7.25098C6.42255 5.75 5.75098 6.42157 5.75098 7.25V16.75C5.75098 17.5784 6.42255 18.25 7.25098 18.25H7.7002V13.1504C7.7002 12.8742 7.92405 12.6504 8.2002 12.6504H15.7998C16.0759 12.6504 16.2998 12.8742 16.2998 13.1504V13.9209C16.2998 14.197 16.0759 14.4209 15.7998 14.4209C15.5237 14.4209 15.2998 14.197 15.2998 13.9209V13.6504H8.7002V18.25H13.252C13.5281 18.25 13.752 18.4739 13.752 18.75C13.752 19.0261 13.5281 19.25 13.252 19.25H7.25098C5.87027 19.25 4.75098 18.1307 4.75098 16.75V7.25C4.75098 5.86929 5.87027 4.75 7.25098 4.75H14.7998ZM8.7002 8.7002H13.7998V5.75H8.7002V8.7002Z" fill="currentColor" />
    <path d="M18.75 14.7495C20.9596 14.7495 22.751 16.5402 22.751 18.7495C22.7509 20.9587 20.9595 22.7495 18.75 22.7495C16.5407 22.7493 14.7501 20.9586 14.75 18.7495C14.75 16.5403 16.5406 14.7497 18.75 14.7495ZM20.4043 17.0962C20.209 16.9009 19.8925 16.9009 19.6973 17.0962L18.751 18.0425L17.8047 17.0962C17.6094 16.9009 17.2929 16.9009 17.0977 17.0962C16.9025 17.2915 16.9024 17.608 17.0977 17.8032L18.0439 18.7495L17.0977 19.6958C16.9025 19.8911 16.9024 20.2076 17.0977 20.4028C17.2929 20.5979 17.6095 20.5979 17.8047 20.4028L18.751 19.4565L19.6973 20.4028C19.8925 20.5979 20.2091 20.5979 20.4043 20.4028C20.5995 20.2076 20.5994 19.8911 20.4043 19.6958L19.458 18.7495L20.4043 17.8032C20.5995 17.608 20.5994 17.2915 20.4043 17.0962Z" fill="#FF274B"/>
    </svg>
      `;
    // error — diskette + red cross badge (bottom-right).
    var SK_SAVE_ICON_SAVED = `
      <svg width="24" height="24" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path fill-rule="evenodd" clip-rule="evenodd" d="M14.7998 4.75C14.9324 4.75 15.0596 4.80278 15.1533 4.89648L19.1025 8.84668C19.1963 8.94044 19.249 9.06763 19.249 9.2002V13.25C19.249 13.5261 19.0252 13.75 18.749 13.75C18.4729 13.75 18.249 13.5261 18.249 13.25V9.40625L14.7998 5.95703V9.2002C14.7997 9.47625 14.5759 9.7002 14.2998 9.7002H8.2002C7.92412 9.7002 7.7003 9.47625 7.7002 9.2002V5.75H7.25098C6.42255 5.75 5.75098 6.42157 5.75098 7.25V16.75C5.75098 17.5784 6.42255 18.25 7.25098 18.25H7.7002V13.1504C7.7002 12.8742 7.92405 12.6504 8.2002 12.6504H15.7998C16.0759 12.6504 16.2998 12.8742 16.2998 13.1504V13.9209C16.2998 14.197 16.0759 14.4209 15.7998 14.4209C15.5237 14.4209 15.2998 14.197 15.2998 13.9209V13.6504H8.7002V18.25H13.252C13.5281 18.25 13.752 18.4739 13.752 18.75C13.752 19.0261 13.5281 19.25 13.252 19.25H7.25098C5.87027 19.25 4.75098 18.1307 4.75098 16.75V7.25C4.75098 5.86929 5.87027 4.75 7.25098 4.75H14.7998ZM8.7002 8.7002H13.7998V5.75H8.7002V8.7002Z" fill="currentColor"/>
      <path d="M18.75 14.749C20.9594 14.749 22.7508 16.5399 22.751 18.749C22.751 20.9583 20.9596 22.75 18.75 22.75C16.5406 22.7498 14.75 20.9582 14.75 18.749C14.7502 16.54 16.5408 14.7493 18.75 14.749ZM21.0547 17.0957C20.8595 16.9005 20.5429 16.9007 20.3477 17.0957L18.1006 19.3418L17.1543 18.3955C16.9591 18.2004 16.6425 18.2004 16.4473 18.3955C16.252 18.5907 16.2521 18.9073 16.4473 19.1025L17.7471 20.4023C17.8408 20.496 17.9681 20.5488 18.1006 20.5488C18.2331 20.5488 18.3604 20.4961 18.4541 20.4023L21.0547 17.8027C21.2496 17.6075 21.2496 17.2909 21.0547 17.0957Z" fill="#2FA0AF"/>
      </svg>
      `;

    // Main App — arrow-into-box icon (16×16), recoloured via currentColor.
    var SK_MAIN_APP_ICON_SVG =
      `<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16" fill="none">
       <path d="M11.5977 2.14941C12.8401 2.14952 13.8474 3.15702 13.8477 4.39941V11.5996C13.8475 12.8421 12.8401 13.8495 11.5977 13.8496H4.39746C3.1551 13.8494 2.14762 12.842 2.14746 11.5996V10.6006C2.14746 10.2414 2.43865 9.9502 2.79785 9.9502C3.15705 9.9502 3.44824 10.2414 3.44824 10.6006V11.5996C3.4484 12.124 3.87307 12.5496 4.39746 12.5498H11.5977C12.1221 12.5497 12.5477 12.1241 12.5479 11.5996V4.39941C12.5476 3.87499 12.1221 3.4503 11.5977 3.4502H4.39746C3.8731 3.45041 3.44845 3.87505 3.44824 4.39941V5.40039C3.44819 5.75955 3.15702 6.05078 2.79785 6.05078C2.43868 6.05078 2.14751 5.75955 2.14746 5.40039V4.39941C2.14767 3.15708 3.15513 2.14963 4.39746 2.14941H11.5977ZM6.22852 5.59082C6.48084 5.33719 6.89165 5.33674 7.14551 5.58887L9.11035 7.54102C9.23239 7.66259 9.30073 7.8287 9.30078 8.00098C9.30057 8.17329 9.2316 8.33848 9.10938 8.45996L7.14551 10.4121C6.8916 10.6642 6.48081 10.6629 6.22852 10.4092C5.97639 10.1552 5.97766 9.74446 6.23145 9.49219L7.08008 8.64941H0.84668C0.488848 8.64941 0.198511 8.35875 0.198242 8.00098C0.198242 7.64298 0.488683 7.35254 0.84668 7.35254H7.08105L6.23145 6.50781C5.97783 6.25547 5.97636 5.84467 6.22852 5.59082Z" fill="currentColor"/>
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
        '.sk-edit-tooltip{',
        '  position:fixed;',
        '  display:none;',
        '  box-sizing:border-box;',
        '  width: max-content;',
        '  max-width: none;',
        '  height:18px;',
        '  padding:2px 8px;',
        '  white-space: nowrap;',
        '  border-radius:3px;',
        '  background:#728596;',
        '  color:#FFFFFF;',
        "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;",
        '  font-size:10px;',
        '  line-height:14px;',
        '  font-weight:400;',
        '  text-align:center;',
        '  z-index:100000;',
        '  pointer-events:none;',
        '}',
      '.sk-edit-btn{',
      '  box-sizing:border-box;',
      '  display:inline-flex;align-items:center;justify-content:center;',
      '  height:24px;padding:0 8px 0 4px;',                        /* fits inside the 28px tab row (was 32px → protruded) */
      '  border:1px solid transparent;border-radius:5px;',          /* keeps box size stable across bordered/unbordered states */
      '  background:#3fc0c4;color:#FFFFFF;',                         /* free state; color drives icon (currentColor) + label */
      "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;",
      '  font-size:12px;line-height:1;font-weight:500;',
      '  white-space:nowrap;cursor:pointer;',
      '  -webkit-appearance:none;appearance:none;margin:0 0 6px;',  /* margin-bottom lifts the button within the row */
      '  transition:background .12s ease,border-color .12s ease,color .12s ease;',
      '}',
      '.sk-edit-btn__icon{display:flex;flex:0 0 auto;width:16px;height:16px;margin-right:4px}',
      '.sk-edit-btn__icon svg{display:block;width:16px;height:16px;}',
      '.sk-edit-btn__label{display:block;}',
      /* free (default) hover — not while pending/disabled */
      '.sk-edit-btn:not(.is-editing):not(.is-locked):not(.is-refresh):not(:disabled):hover{background:#41D1C9;}',
      /* editing — you hold the lock */
      '.sk-edit-btn.is-editing{background:#FFFFFF;color:#2FA0AF;border-color:#2FA0AF;}',
      '.sk-edit-btn.is-editing:hover{background:#F0F8F9;}',
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
      '  margin-bottom:6px;',
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
	  '.btn-slot{margin-bottom:6px}',
      /* ── Main App button — header-right, before the search slot. */
      '.sk-main-app-slot{display:inline-flex;align-items:center;margin-right:8px;vertical-align:middle;}',
      '.sk-main-app-btn{',
      '  box-sizing:border-box;',
      '  display:inline-flex;align-items:center;justify-content:center;gap:5px;',
      '  min-width:86px;height:24px;padding:2px 9px;',
      '  border:none;border-radius:6px;',
      '  background:rgba(53,80,105,0.1);color:rgba(53,80,105,0.8);',
      "  font-family:'New Hero',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;",
      '  font-size:12px;line-height:16px;font-weight:500;white-space:nowrap;',
      '  cursor:pointer;-webkit-appearance:none;appearance:none;margin:0 0 6px;',
      '  transition:background .12s ease;',
      '}',
      '.sk-main-app-btn__icon{display:flex;flex:0 0 auto;flex-shrink:0;width:16px;height:16px;}',
      '.sk-main-app-btn__icon svg{display:block;flex-shrink:0;width:16px;height:16px;}',
      '.sk-main-app-btn:not(:disabled):hover{background:rgba(53,80,105,0.2);}',
      '.theme-type-dark .sk-main-app-btn{background:rgba(255,255,255,0.1);color:rgba(255,255,255,0.8);}',
      '.theme-type-dark .sk-main-app-btn:not(:disabled):hover{background:rgba(255,255,255,0.2);}'
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

      function ensureEditTooltip(doc) {
          if (headerEditTooltip && headerEditTooltip.ownerDocument === doc) {
              headerEditTooltip.textContent = isDesktopClosing ? 'Cancel closing the Main App to start editing' : 'Only one Member can edit at a time';

              return headerEditTooltip;
          }

          headerEditTooltip = doc.createElement('div');
          headerEditTooltip.className = 'sk-edit-tooltip';

          headerEditTooltip.textContent = isDesktopClosing ? 'Cancel closing the Main App to start editing' : 'Only one Member can edit at a time';

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

      function showEditTooltip() {
          if (!headerEditBtn) {
              return;
          }

          if (!headerEditBtn.classList.contains('is-locked')) {
              return;
          }

          var doc = headerEditBtn.ownerDocument;
          var tooltip = ensureEditTooltip(doc);

          if (!tooltip) {
              return;
          }

          var TOOLTIP_SCREEN_PADDING = 8;
          var TOOLTIP_OFFSET = 6;
          var rect = headerEditBtn.getBoundingClientRect();

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

      function hideEditTooltip() {
          if (!headerEditTooltip) {
              return;
          }

          headerEditTooltip.style.display = 'none';
      }

      function onEditTabClick() {
          if (!headerEditBtn) {
              return;
          }

          if (!headerEditBtn.classList.contains('is-locked')) {
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
              return;   // role has no edit right — refuse (button shouldn't exist anyway)
          }

          if (isDesktopClosing) {
              showDesktopClosingModal();

              return;
          }

          if (currentMode === 'view') {
              log('user clicked Edit → request-edit-mode');
              pm.toHost({ type: 'request-edit-mode' });
              headerEditBtn.disabled = true;   // pending; re-enabled on set-mode
          } else if (currentMode === 'edit') {
              // Save pending changes BEFORE releasing the edit lock, so leaving edit
              // mode never drops unsaved work. triggerAutosave is a no-op if the doc
              // isn't dirty or a save is already in flight; when it does fire, the
              // diskette reflects saving→saved via the shared state funnel.
              log('user clicked Edit (editing) → save-on-exit + mode-changed: view');
              pm.triggerAutosave();
              pm.toHost({ type: 'mode-changed', mode: 'view' });
              headerEditBtn.disabled = true;
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
      var MAIN_APP_WINDOW_NAME = 'sharekey-main';

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
      var editDone = mountEditButton(doc);
      mountEditingLabel(doc);              // anchored off the Edit <li>; no-op if edit not mounted yet
      var saveDone = mountSaveButton(doc);
      var shouldMountMainAppButton = window.SK_DESKTOP_TRANSPORT || hasOpener;
      var mainAppDone = !shouldMountMainAppButton || mountMainAppButton(doc);

      return editDone && saveDone && mainAppDone;
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
            existing.parentNode.onmouseenter = showEditTooltip;
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
      slot.onmouseenter = showEditTooltip;
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
      btn.innerHTML =
        '<span class="sk-main-app-btn__icon">' + SK_MAIN_APP_ICON_SVG + '</span>' +
        '<span class="sk-main-app-btn__label">Main App</span>';

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
    var nextDesktopClosing = !!(perms && perms.isDesktopClosing);

    if (nextCanEdit === canEdit && nextDesktopClosing === isDesktopClosing) {
        return;
    }

    canEdit = nextCanEdit;
    isDesktopClosing = nextDesktopClosing;
      log('handlePermissions: canEdit=' + canEdit);
      // Re-render the Edit button + editing label to reflect the new capability
      // (renderEditButton hides the button when !canEdit). The controls are
      // mounted unconditionally by the poller, so this just flips visibility —
      // no DOM removal/race.
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
          if (!window.__editorDirty) {
              return;
          }

          if (window.SK_DESKTOP_TRANSPORT && pm) {
              e.preventDefault();

              log('desktop beforeunload + dirty → saveAndClose');

              pm.saveAndClose();

              return;
          }

          e.preventDefault();
          e.returnValue = 'You have unsaved edits. Close anyway?';

          return e.returnValue;
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
})();
