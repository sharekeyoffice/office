(function (global) {
    'use strict';

    var priorities = {
        'main-app-logged-out-modal': 100,
        'main-app-closed-modal': 90,
        'deleted-file-modal': 85,
        'access-lost-modal': 85,
        'cannot-reconnect-modal': 80,
        'reconnecting-modal': 70,
        'document-password-modal': 65,
        'cannot-start-edit-mode': 60,
        'viewer-mode': 50,
        'view-only-mode': 50,
        'turn-on-edit-mode': 40,
        'welcome-screen': 1
    };

    var closeButtonIds = {
        'document-password-modal': 'dmp-close-btn',
        'cannot-start-edit-mode': 'csem-close-btn',
        'viewer-mode': 'vm-close-btn',
        'view-only-mode': 'vom-close-btn',
        'turn-on-edit-mode': 'toem-close-btn',
        'welcome-screen': 'ws-close-btn',
        'deleted-file-modal': 'dfm-close-btn',
        'access-lost-modal': 'alm-close-btn',
    };

    var activeModals = {};
    var keyboardDocuments = new WeakSet();

    function handleModalKeyDown(event) {
        const isEscape = event.key === 'Escape' && !event.repeat && !event.isComposing;

        if (!isEscape) {
            return;
        }

        const topModalId = getTopModalId();
        const closeButtonId = closeButtonIds[topModalId];
        const closeButton = closeButtonId && document.getElementById(closeButtonId);

        if (!closeButton) {
            return;
        }

        event.preventDefault();
        event.stopImmediatePropagation();

        closeButton.dispatchEvent(new global.MouseEvent('click', { bubbles: true }));
    }

    function bindKeyboardDocument(targetDocument) {
        if (!targetDocument || keyboardDocuments.has(targetDocument)) {
            return;
        }

        targetDocument.addEventListener('keydown', handleModalKeyDown, true);
        keyboardDocuments.add(targetDocument);
    }

    function getTopModalId() {
        var modalIds = Object.keys(activeModals);
        var topModalId = null;
        var topPriority = -1;

        modalIds.forEach(function (modalId) {
            var priority = priorities[modalId] || 0;

            if (priority > topPriority) {
                topPriority = priority;
                topModalId = modalId;
            }
        });

        return topModalId;
    }

    function render() {
        var topModalId = getTopModalId();

        Object.keys(priorities).forEach(function (modalId) {
            var modal = document.getElementById(modalId);

            if (!modal) {
                return;
            }

            modal.style.display = modalId === topModalId ? 'flex' : 'none';
        });
    }

    function show(modalId) {
        const editorIframe = document.querySelector('iframe[name="frameEditor"]');

        if (editorIframe) {
            bindKeyboardDocument(editorIframe.contentDocument);
        }

        activeModals[modalId] = true;

        render();
    }

    function hide(modalId) {
        delete activeModals[modalId];

        render();
    }

    function isActive(modalId) {
        return !!activeModals[modalId];
    }

    function requestDocumentPassword(verifyPassword, abortSignal) {
        if (abortSignal.aborted || isActive('document-password-modal')) {
            return Promise.resolve(null);
        }

        const passwordModal = document.getElementById('document-password-modal');
        const passwordForm = passwordModal.querySelector('form');
        const passwordInput = passwordForm.elements.namedItem('password');
        const passwordErrorMessage = document.getElementById('dpm-error');
        const passwordVisibilityButton = passwordForm.querySelector('.document-password-toggle');
        const openDocumentButton = passwordForm.querySelector('button[type="submit"]');
        const closeButton = document.getElementById('dmp-close-btn');

        function setPasswordVisibility(isPasswordVisible) {
            passwordInput.type = isPasswordVisible ? 'text' : 'password';
            passwordVisibilityButton.setAttribute('aria-pressed', String(isPasswordVisible));
        }

        function updatePasswordErrorState(isPasswordIncorrect) {
            passwordInput.setAttribute('aria-invalid', String(isPasswordIncorrect));
            passwordErrorMessage.style.visibility = isPasswordIncorrect ? 'visible' : 'hidden';
            openDocumentButton.disabled = isPasswordIncorrect || !passwordInput.value;
        }

        function resetPasswordForm() {
            passwordForm.reset();
            setPasswordVisibility(false);
            updatePasswordErrorState(false);
            passwordInput.readOnly = false;
            passwordVisibilityButton.disabled = false;
        }

        resetPasswordForm();
        show('document-password-modal');
        passwordInput.focus();

        return new Promise((resolve, reject) => {
            let isPasswordPromptClosed = false;

            function closePasswordPrompt() {
                isPasswordPromptClosed = true;
                resetPasswordForm();
                passwordForm.onsubmit = null;
                passwordInput.oninput = null;
                passwordVisibilityButton.onmousedown = null;
                passwordVisibilityButton.onclick = null;

                if (closeButton) {
                    closeButton.onclick = null;
                }

                abortSignal.removeEventListener('abort', cancelPasswordPrompt);
                hide('document-password-modal');
            }

            function cancelPasswordPrompt() {
                closePasswordPrompt();
                resolve(null);
            }

            passwordInput.oninput = () => updatePasswordErrorState(false);
            passwordVisibilityButton.onmousedown = event => event.preventDefault();
            passwordVisibilityButton.onclick = () => {
                const selectionStart = passwordInput.selectionStart;
                const selectionEnd = passwordInput.selectionEnd;
                setPasswordVisibility(passwordInput.type === 'password');
                passwordInput.setSelectionRange(selectionStart, selectionEnd);
            };

            if (closeButton) {
                closeButton.onclick = cancelPasswordPrompt;
            }

            abortSignal.addEventListener('abort', cancelPasswordPrompt, { once: true });
            passwordForm.onsubmit = async event => {
                event.preventDefault();

                if (openDocumentButton.disabled) {
                    return;
                }

                openDocumentButton.disabled = true;
                passwordVisibilityButton.disabled = true;
                passwordInput.readOnly = true;

                try {
                    const decryptedDocumentBytes = await verifyPassword(passwordInput.value);

                    if (isPasswordPromptClosed) {
                        return;
                    }

                    if (decryptedDocumentBytes) {
                        closePasswordPrompt();
                        resolve(decryptedDocumentBytes);

                        return;
                    }

                    passwordInput.readOnly = false;
                    passwordVisibilityButton.disabled = false;
                    updatePasswordErrorState(true);
                    passwordInput.focus();
                } catch (verificationError) {
                    if (isPasswordPromptClosed) {
                        return;
                    }

                    closePasswordPrompt();
                    reject(verificationError);
                }
            };
        });
    }

    bindKeyboardDocument(document);

    global.modalManager = {
        requestDocumentPassword: requestDocumentPassword,
        show: show,
        hide: hide,
        isActive: isActive
    };
})(window);
