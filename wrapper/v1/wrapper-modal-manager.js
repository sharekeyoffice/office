(function (global) {
    'use strict';

    var priorities = {
        'main-app-logged-out-modal': 100,
        'main-app-closed-modal': 90,
        'cannot-reconnect-modal': 80,
        'reconnecting-modal': 70,
        'cannot-start-edit-mode': 60,
        'viewer-mode': 50,
        'view-only-mode': 50,
        'turn-on-edit-mode': 40,
        'welcome-screen': 1
    };

    var activeModals = {};

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

    global.modalManager = {
        show: show,
        hide: hide,
        isActive: isActive
    };
})(window);
