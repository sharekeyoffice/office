(function (global) {
    'use strict';

    const PROPERTY_NAME = 'OfficeEditorPreferences';
    let disposeCurrent = null;
    let getSnapshot = function () { return '{}'; };
    let savedSnapshot = '{}';

    function install(options) {
        if (disposeCurrent) {
            disposeCurrent();
        }

        const api = options.editorApi;
        const win = options.editorWindow;
        const type = options.type;
        let restoring = false;
        let borderListeners = [];
        const themeColorUpdates = [];
        const toolbarController = type === 'cell' && win.SSE.controllers.Toolbar;
        const rightMenuController = type === 'cell' && win.SSE.controllers.RightMenu;
        const toolbar = toolbarController && toolbarController.toolbar;
        const rightMenu = rightMenuController && rightMenuController.rightmenu;
        const notifications = toolbar && win.Common.NotificationCenter;
        const zoomEvent = type === 'cell' ? 'asc_onZoomChanged' : 'asc_onZoomChange';

        function read() {
            const property = api.asc_getAllCustomProperties().find(function (item) {
                return item.asc_getName() === PROPERTY_NAME;
            });

            if (!property) {
                return {};
            }

            try {
                const value = JSON.parse(property.asc_getValue());

                return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
            } catch (error) {
                return {};
            }
        }

        function update(changes) {
            if (restoring || !options.canSave()) {
                return;
            }

            const previous = read();
            const serialized = JSON.stringify(Object.assign({}, previous, changes));

            if (JSON.stringify(previous) === serialized) {
                return;
            }

            const history = win.AscCommon.History;
            history.TurnOff();

            try {
                api.getCustomProperties().AddProperty(PROPERTY_NAME, win.AscCommon.c_oVariantTypes.vtLpwstr, serialized);
            } finally {
                history.TurnOn();
            }

            options.onChange();
        }

        function onBorderStyle(menu, item, checked) {
            if (checked) {
                update({ borderStyle: item.value });
                restoreBorders({ borderStyle: item.value });
            }
        }

        function onSidebarBorderStyle(combo, record) {
            update({ borderStyle: record.value });
            restoreBorders({ borderStyle: record.value });
        }

        function getBorderColor(color) {
            if (color === 'auto') {
                return color;
            }

            const value = typeof color === 'string' ? color : color && color.color;

            if (typeof value !== 'string' || !/^[a-f\d]{3}([a-f\d]{3})?$/i.test(value)) {
                return null;
            }

            if (typeof color === 'string') {
                return color;
            }

            const themeColors = win.Common.Utils.ThemeColor.getEffectColors();
            const themeColor = themeColors && themeColors.find(function (item) {
                return item.effectId === color.effectId;
            });

            return {
                color: themeColor ? themeColor.color : value,
                effectId: color.effectId,
                effectValue: color.effectValue
            };
        }

        function onBorderColor(control, color) {
            color = getBorderColor(color);

            if (color) {
                update({ borderColor: color });
                restoreBorders({ borderColor: color });
            }
        }

        function onAutoBorderColor() {
            onBorderColor(null, 'auto');
        }

        function getAutoBorderColorItem() {
            const menu = toolbar.mnuBorderColor && toolbar.mnuBorderColor.menu;

            return menu && menu.items.find(function (item) {
                return item.id === 'id-toolbar-menu-auto-bordercolor';
            });
        }

        function restoreBorders(preferences) {
            const menu = toolbar.mnuBorderWidth;
            const settings = rightMenu && rightMenu.cellSettings;
            const style = preferences.borderStyle;
            const isSupportedStyle = menu && menu.items.some(function (item) {
                return item.value === style;
            });
            const color = getBorderColor(preferences.borderColor);
            restoring = true;

            try {
                if (isSupportedStyle) {
                    toolbar.btnBorders.options.borderswidth = style;
                    menu.items.forEach(function (item) {
                        item.setChecked(item.value === style, true);
                    });
                }

                if (settings && settings.cmbBorderType && settings.cmbBorderType.store.findWhere({ value: style })) {
                    settings.BorderType = style;
                    settings.cmbBorderType.setValue(style);
                }

                if (!color) {
                    return;
                }

                const isAutoColor = color === 'auto';
                const picker = toolbar.mnuBorderColorPicker;

                if (picker) {
                    let sdkColor;

                    if (isAutoColor) {
                        sdkColor = new win.Asc.asc_CColor();
                        sdkColor.put_auto(true);
                        picker.clearSelection();
                    } else {
                        sdkColor = win.Common.Utils.ThemeColor.getRgbColor(color);
                        picker.select(color, true);
                    }

                    picker.currentColor = isAutoColor ? { color: '000000', isAuto: true } : color;
                    toolbar.btnBorders.options.borderscolor = sdkColor;
                    toolbar.mnuBorderColor.cmpEl.find('.menu-item-icon').css('border-color', '#' + (isAutoColor ? '000000' : color.color || color));
                    getAutoBorderColorItem().cmpEl.find('> a').toggleClass('selected', isAutoColor);
                }

                if (settings && settings.btnBorderColor) {
                    const button = settings.btnBorderColor;
                    button.options.color = color;
                    button.setColor(color);
                    button.setAutoColor(isAutoColor);

                    if (button.colorPicker) {
                        if (isAutoColor) {
                            button.colorPicker.clearSelection();
                        } else {
                            button.colorPicker.select(color, true);
                        }
                    }
                }
            } finally {
                restoring = false;
            }
        }

        function addBorderListener(control, event, callback) {
            if (control) {
                control.on(event, callback);
                borderListeners.push({ control: control, event: event, callback: callback });
            }
        }

        function removeBorderListeners() {
            borderListeners.forEach(function (listener) {
                listener.control.off(listener.event, listener.callback);
            });
            borderListeners = [];
        }

        function bindBorderControls() {
            removeBorderListeners();
            addBorderListener(toolbar.mnuBorderWidth, 'item:toggle', onBorderStyle);
            addBorderListener(toolbar.mnuBorderColorPicker, 'select', onBorderColor);
            addBorderListener(getAutoBorderColorItem(), 'click', onAutoBorderColor);

            const settings = rightMenu && rightMenu.cellSettings;

            if (settings) {
                addBorderListener(settings.cmbBorderType, 'selected', onSidebarBorderStyle);
                addBorderListener(settings.btnBorderColor, 'color:select', onBorderColor);
                addBorderListener(settings.btnBorderColor, 'auto:select', onAutoBorderColor);
            }

            restoreBorders(read());
        }

        function preserveBordersAfterThemeUpdate(controller, method) {
            const original = controller[method];

            const updateThemeColors = function () {
                const result = original.apply(this, arguments);
                restoreBorders(read());
                return result;
            };

            controller[method] = updateThemeColors;
            themeColorUpdates.push(function () {
                if (controller[method] === updateThemeColors) {
                    controller[method] = original;
                }
            });
        }

        function getZoom(percent, zoomType) {
            if (type !== 'cell' && zoomType === 1) {
                return { mode: 'fit-width' };
            }

            if (type !== 'cell' && zoomType === 2) {
                return { mode: 'fit-page' };
            }

            const value = type === 'cell' ? Math.round(percent * 100) : percent;
            const maxZoom = type === 'cell' ? 400 : 500;

            return Number.isFinite(value) && value >= 10 && value <= maxZoom ?
                { mode: 'percent', value: value } :
                null;
        }

        function restoreZoom(zoom) {
            if (!zoom || typeof zoom !== 'object') {
                return;
            }

            if (type !== 'cell' && zoom.mode === 'fit-page') {
                api.zoomFitToPage();
            } else if (type !== 'cell' && zoom.mode === 'fit-width') {
                api.zoomFitToWidth();
            } else if (zoom.mode === 'percent' && getZoom(type === 'cell' ? zoom.value / 100 : zoom.value, 0)) {
                if (type === 'cell') {
                    api.asc_setZoom(zoom.value / 100);
                } else {
                    api.zoom(zoom.value);
                }
            }
        }

        getSnapshot = function () {
            return JSON.stringify(read());
        };
        savedSnapshot = getSnapshot();

        restoring = true;
        restoreZoom(read().zoom);
        restoring = false;

        let lastZoom = type === 'cell' ?
            getZoom(api.asc_getZoom(), 0) :
            getZoom(api.WordControl.m_nZoomValue, api.WordControl.m_nZoomType);

        function onZoomChange(percent, zoomType) {
            const zoom = getZoom(percent, zoomType);

            if (restoring || !zoom || JSON.stringify(zoom) === JSON.stringify(lastZoom)) {
                return;
            }

            lastZoom = zoom;
            update({ zoom: zoom });
        }

        api.asc_registerCallback(zoomEvent, onZoomChange);

        if (toolbar) {
            preserveBordersAfterThemeUpdate(toolbarController, 'updateThemeColors');
            if (rightMenuController) {
                preserveBordersAfterThemeUpdate(rightMenuController, 'UpdateThemeColors');
            }

            toolbar.on('render:after', bindBorderControls);
            notifications.on('document:ready', bindBorderControls);

            if (rightMenu) {
                rightMenu.on('render:after', bindBorderControls);
            }

            bindBorderControls();
        }

        disposeCurrent = function () {
            api.asc_unregisterCallback(zoomEvent, onZoomChange);

            if (toolbar) {
                toolbar.off('render:after', bindBorderControls);
                notifications.off('document:ready', bindBorderControls);

                if (rightMenu) {
                    rightMenu.off('render:after', bindBorderControls);
                }
            }

            removeBorderListeners();
            themeColorUpdates.forEach(function (restore) { restore(); });
        };
    }

    global.DocumentPreferences = {
        install: install,
        getSnapshot: function () {
            return getSnapshot();
        },
        markSaved: function (snapshot) {
            savedSnapshot = snapshot;
        },
        hasUnsavedChanges: function () {
            return getSnapshot() !== savedSnapshot;
        }
    };
})(window);
