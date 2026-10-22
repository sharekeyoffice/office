(function () {
    'use strict';

    var FORMULA_BAR_INPUT_ID = 'ce-cell-content';
    var editorTypes = {
        word: 'word',
        slide: 'slide',
        cell: 'cell'
    };

    var editorApi = null;
    var editorWindow = null;
    var editorType = null;
    var logicDocument = null;
    var history = null;
    var isRepeatAllowed = function () {
        return false;
    };

    var currentRepeatCandidate = null;
    var lastRepeatCandidate = null;
    var lastTypedCharacter = null;
    var isBackspaceLastAction = false;
    var repeatApiCallDepth = 0;
    var isRepeatingAction = false;
    var isInstalled = false;

    var originalStartAction = null;
    var originalCreateNewPoint = null;
    var originalOnKeyDown = null;

    function log() {
        if (window.console) {
            console.log.apply(console, ['[repeat-last-action]'].concat([].slice.call(arguments)));
        }
    }

    function getLogicDocument() {
        if (editorApi && editorApi.WordControl) {
            return editorApi.WordControl.m_oLogicDocument || null;
        }
    }

    function getHistory() {
        if (editorType === editorTypes.cell) {
            return editorApi && editorApi.wb && editorApi.wb.model ?
                editorApi.wb.model.History || null :
                null;
        }

        var currentLogicDocument = getLogicDocument();

        if (currentLogicDocument) {
            return currentLogicDocument.History || null;
        }
    }

    function isRedoShortcut(event) {
        if (!event) {
            return false;
        }

        var key = event.key && event.key.toLowerCase();
        var keyCode = event.keyCode || event.which || (typeof event.GetKeyCode === 'function' ? event.GetKeyCode() : 0);
        var hasPrimaryModifier = event.metaKey || event.ctrlKey || event.MacCmdKey || event.CtrlKey;
        var hasShift = event.shiftKey || event.ShiftKey;
        var hasAlt = event.altKey || event.AltKey;

        return (key === 'y' || event.code === 'KeyY' || keyCode === 89) &&
            !!hasPrimaryModifier &&
            !hasShift &&
            !hasAlt;
    }

    function hasRedoAvailable() {
        if (editorApi && typeof editorApi.asc_getCanRedo === 'function') {
            try {
                return !!editorApi.asc_getCanRedo();
            } catch (e) {}
        }

        if (history && typeof history.Can_Redo === 'function') {
            try {
                return !!history.Can_Redo();
            } catch (e) {}
        }

        return false;
    }

    function markCurrentRepeatCandidate() {
        if (currentRepeatCandidate && !isRepeatingAction) {
            currentRepeatCandidate.hasAction = true;
        }
    }

    function wrapApiMethod(methodName) {
        var originalMethod;

        try {
            originalMethod = editorApi[methodName];
        } catch (e) {
            return;
        }

        var isMethodExcluded = typeof originalMethod !== 'function' ||
            originalMethod.__isRepeatLastActionWrapped ||
            (editorType === editorTypes.cell && (methodName === 'onKeyDown' || methodName === 'asc_closeCellEditor'));

        if (isMethodExcluded) {
            return;
        }

        var wrappedMethod = function () {
            var isTopLevelCall = repeatApiCallDepth === 0;
            var isTrackingCall = isTopLevelCall && !isRepeatingAction && isRepeatAllowed();

            if (isTrackingCall) {
                currentRepeatCandidate = {
                    method: methodName,
                    arguments: Array.prototype.slice.call(arguments),
                    hasAction: false
                };
            }

            repeatApiCallDepth += 1;

            try {
                return originalMethod.apply(this, arguments);
            } finally {
                repeatApiCallDepth -= 1;

                if (isTrackingCall) {
                    if (currentRepeatCandidate && currentRepeatCandidate.hasAction) {
                        lastRepeatCandidate = {
                            method: currentRepeatCandidate.method,
                            arguments: currentRepeatCandidate.arguments
                        };

                        log('repeat candidate:', lastRepeatCandidate.method);
                    }

                    currentRepeatCandidate = null;
                }
            }
        };

        wrappedMethod.__isRepeatLastActionWrapped = true;

        try {
            editorApi[methodName] = wrappedMethod;
        } catch (e) {}
    }

    function wrapApiMethods() {
        var current = editorApi;

        var visited = new Set();
        var wrappedMethodNames = new Set();

        while (current && !visited.has(current)) {
            visited.add(current);

            Object.getOwnPropertyNames(current).forEach(function (methodName) {
                if (methodName === 'constructor' || wrappedMethodNames.has(methodName)) {
                    return;
                }

                wrappedMethodNames.add(methodName);
                wrapApiMethod(methodName);
            });

            current = Object.getPrototypeOf(current);
        }
    }

    function installActionMarkers() {
        logicDocument = getLogicDocument();
        history = getHistory();

        if (editorType === editorTypes.cell) {
            if (!history || typeof history.Create_NewPoint !== 'function') {
                return false;
            }
        } else {
            if (!logicDocument || typeof logicDocument.StartAction !== 'function') {
                return false;
            }

            originalStartAction = logicDocument.StartAction;

            logicDocument.StartAction = function () {
                markCurrentRepeatCandidate();

                return originalStartAction.apply(this, arguments);
            };
        }

        if (history && typeof history.Create_NewPoint === 'function') {
            originalCreateNewPoint = history.Create_NewPoint;

            history.Create_NewPoint = function () {
                markCurrentRepeatCandidate();

                return originalCreateNewPoint.apply(this, arguments);
            };
        }

        return true;
    }

    function repeatLastAction() {
        if (!lastRepeatCandidate || isRepeatingAction || !isRepeatAllowed()) {
            return false;
        }

        var method;

        try {
            method = editorApi[lastRepeatCandidate.method];
        } catch (e) {
            return false;
        }

        if (typeof method !== 'function') {
            return false;
        }

        isRepeatingAction = true;

        try {
            log('repeat:', lastRepeatCandidate.method);

            method.apply(editorApi, lastRepeatCandidate.arguments);
        } finally {
            isRepeatingAction = false;
        }

        return true;
    }

    function installFormulaBarInputTracker() {
        if (editorType !== editorTypes.cell) {
            return;
        }

        ['input', 'beforeinput'].forEach(function (eventName) {
            editorWindow.document.addEventListener(eventName, function (event) {
                var isFormulaBarInput = eventName === 'input' && event.target.id === FORMULA_BAR_INPUT_ID;
                var isCellInput = eventName === 'beforeinput' && event.target.id === 'area_id';
                var isFormulaBarBackspace = eventName === 'beforeinput' &&
                    event.target.id === FORMULA_BAR_INPUT_ID &&
                    event.inputType === 'deleteContentBackward';

                if (isFormulaBarBackspace) {
                    isBackspaceLastAction = true;
                    lastTypedCharacter = null;
                    lastRepeatCandidate = null;

                    log('last action: backspace');

                    return;
                }

                if ((isFormulaBarInput || isCellInput) && event.inputType === 'insertText' && event.data) {
                    isBackspaceLastAction = false;
                    lastTypedCharacter = Array.from(event.data).pop();
                    lastRepeatCandidate = null;

                    log('last typed character:', lastTypedCharacter);
                }
            }, true);
        });
    }

    function installFormattingShortcutHandler() {
        if (editorType !== editorTypes.cell || !editorApi.wb || typeof editorApi.wb.setFontAttributes !== 'function') {
            return;
        }

        var originalSetFontAttributes = editorApi.wb.setFontAttributes;

        var formattingMethods = {
            b: 'asc_setCellBold',
            i: 'asc_setCellItalic',
            u: 'asc_setCellUnderline',
            s: 'asc_setCellStrikeout'
        };

        editorApi.wb.setFontAttributes = function (attribute, value) {
            var result = originalSetFontAttributes.apply(this, arguments);

            var methodName = formattingMethods[attribute];

            var isFormattingAction = methodName &&
                (typeof value === 'boolean' || (attribute === 'u' && typeof value === 'number')) &&
                repeatApiCallDepth === 0 &&
                !isRepeatingAction &&
                isRepeatAllowed();

            if (isFormattingAction) {
                lastRepeatCandidate = {
                    method: methodName,
                    arguments: [value]
                };

                lastTypedCharacter = null;
                isBackspaceLastAction = false;

                log('repeat candidate:', lastRepeatCandidate.method);
            }

            return result;
        };
    }

    function installEditorKeyDownHandler() {
        editorApi.onKeyDown = function (event) {
            var isRedo = isRedoShortcut(event);
            var cellEditor = editorType === editorTypes.cell && editorApi.wb && editorApi.wb.cellEditor;
            var keyCode = event.keyCode || event.which || (
                typeof event.GetKeyCode === 'function' ? event.GetKeyCode() : 0
            );
            var isBackspace = event.key === 'Backspace' || keyCode === 8;
            var isCellEditing = cellEditor && cellEditor.isOpened && !cellEditor.isTopLineActive;

            if (isBackspace && isCellEditing && isRepeatAllowed()) {
                isBackspaceLastAction = true;
                lastTypedCharacter = null;
                lastRepeatCandidate = null;

                log('last action: backspace');
            }

            var isCellRepeat = isRedo && isCellEditing && !isRepeatingAction && isRepeatAllowed() &&
                (cellEditor.redoList.length > 0 || lastTypedCharacter || isBackspaceLastAction);

            if (isCellRepeat) {
                event.preventDefault();

                if (cellEditor.redoList.length > 0) {
                    cellEditor.redo();

                    log('cell editor redo');
                } else if (isBackspaceLastAction) {
                    cellEditor._removeChars(-8);

                    log('repeat backspace');
                } else {
                    cellEditor._addChars(lastTypedCharacter);

                    log('repeat character:', lastTypedCharacter);
                }

                return false;
            }

            var hasRedo = isRedo && hasRedoAvailable();
            var result = originalOnKeyDown.apply(this, arguments);

            if (isRedo && !hasRedo && lastRepeatCandidate && !isRepeatingAction && isRepeatAllowed()) {
                repeatLastAction();
            }

            return result;
        };
    }

    function installFormulaBarShortcutHandler() {
        if (editorType !== editorTypes.cell || typeof editorApi.asc_registerCallback !== 'function') {
            return;
        }

        editorApi.asc_registerCallback('asc_onInputKeyDown', function (event) {
            var isFormulaBarShortcut = event && event.target && event.target.id === FORMULA_BAR_INPUT_ID &&
                isRedoShortcut(event) && !isRepeatingAction && isRepeatAllowed();

            if (!isFormulaBarShortcut) {
                return;
            }

            var cellEditor = editorApi.wb.cellEditor;
            var isFormulaBarEditing = cellEditor.isOpened && cellEditor.isTopLineActive;
            var hasCellRedo = isFormulaBarEditing && cellEditor.redoList.length > 0;
            var hasRedo = hasRedoAvailable();
            var isBackspaceRepeat = isFormulaBarEditing && !hasRedo &&
                isBackspaceLastAction && cellEditor.input.selectionStart === cellEditor.input.selectionEnd &&
                cellEditor.input.selectionStart > 0;
            var isCharacterRepeat = isFormulaBarEditing && !hasRedo && !!lastTypedCharacter;
            var isTextUpdated = false;

            if (hasCellRedo) {
                event.preventDefault();

                cellEditor.redo();

                isTextUpdated = true;
            } else if (isBackspaceRepeat) {
                event.preventDefault();

                cellEditor._updateCursorByTopLine();
                cellEditor._removeChars(-8);

                isTextUpdated = true;

                log('repeat backspace');
            } else if (isCharacterRepeat) {
                event.preventDefault();

                cellEditor._updateCursorByTopLine();
                cellEditor._addChars(lastTypedCharacter);

                isTextUpdated = true;

                log('repeat character:', lastTypedCharacter);
            } else if (!hasRedo && lastRepeatCandidate) {
                event.preventDefault();

                repeatLastAction();
            }

            if (isTextUpdated) {
                cellEditor.input.value = cellEditor.getText();
                cellEditor._updateTopLineCurPos();
            }
        });
    }

    function installShortcutHandler() {
        if (!editorApi || typeof editorApi.onKeyDown !== 'function') {
            return false;
        }

        originalOnKeyDown = editorApi.onKeyDown;

        installFormattingShortcutHandler();
        installEditorKeyDownHandler();
        installFormulaBarShortcutHandler();

        return true;
    }

    function install(options) {
        if (isInstalled) {
            return true;
        }

        editorApi = options && options.editorApi;
        editorWindow = options && options.editorWindow;
        editorType = options && options.type;

        if (options && typeof options.isRepeatAllowed === 'function') {
            isRepeatAllowed = options.isRepeatAllowed;
        }

        var isSupportedEditor = !!editorApi && !!editorWindow &&
            (((editorType === editorTypes.word || editorType === editorTypes.slide) && !!getLogicDocument()) ||
                (editorType === editorTypes.cell && !!getHistory()));

        if (!isSupportedEditor) {
            return false;
        }

        wrapApiMethods();

        var hasActionMarkers = installActionMarkers();
        var hasShortcutHandler = installShortcutHandler();

        isInstalled = hasActionMarkers && hasShortcutHandler;

        if (isInstalled) {
            installFormulaBarInputTracker();

            log('installed for ' + editorType);
        }

        return isInstalled;
    }

    window.installRepeatLastActionManager = install;
})();
