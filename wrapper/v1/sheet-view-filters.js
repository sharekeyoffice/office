(function (global) {
  'use strict';

  // View filters change the displayed model only. Keep its export and undo history
  // at the original state, and restore that state before granting edit access.
  function install(editorWindow) {
    const controllers = editorWindow.SSE && editorWindow.SSE.controllers;
    const api = controllers && controllers.Main && controllers.Main.api;

    if (!api || !api.wb) {
      return;
    }

    if (api.__skViewFilters) {
      api.__skViewFilters.refresh();
      return;
    }

    const history = editorWindow.AscCommon.History;
    const restrictions = editorWindow.Asc.c_oAscRestrictionType;
    const excel = editorWindow.AscCommonExcel;
    const captureSave = editorWindow.__captureSave;
    const document = editorWindow.document;
    const openFilterDialogs = new Set();
    let session = null;
    let applyingFilter = false;
    let hookedToolbar = null;

    const worksheetPrototype = editorWindow.AscCommonExcel.WorksheetView.prototype;
    const getCursor = worksheetPrototype.getCursorTypeFromXY;

    worksheetPrototype.getCursorTypeFromXY = function (x, y) {
      const cursor = getCursor.apply(this, arguments);
      const isCell = cursor && cursor.target === excel.c_oTargetType.Cells;

      if (!isViewer() || !isCell) {
        return cursor;
      }

      let filterCursor = null;
      const sheetX = this.getRightToLeft() ? this.getCtxWidth() - x : x;

      this._drawElements(function (range, offsetX, offsetY) {
        if (range.contains(cursor.col, cursor.row)) {
          filterCursor = this.af_checkCursor(sheetX + offsetX, y + offsetY, cursor.row, cursor.col);
        }

        return filterCursor === null;
      });

      return filterCursor || cursor;
    };

    function isViewer() {
      return api.restrictions === restrictions.View;
    }

    function currentSession() {
      if (session && session.model !== api.wb.model) {
        session = null;
      }

      return session;
    }

    function cloneFilter(filter) {
      return filter ? filter.clone() : filter;
    }

    function markFilterDialog(dialog) {
      dialog.$window.attr('data-sk-view-filter', '');
      openFilterDialogs.add(dialog);
      dialog.on('close', function () { openFilterDialogs.delete(dialog); });
    }

    const hookedDialogPrototypes = new Set();

    function hookFilterDialogs() {
      ['DigitalFilterDialog', 'Top10FilterDialog'].forEach(function (dialogName) {
        const definition = editorWindow.SSE.Views[dialogName];

        // Before the lazy view module loads, these names contain translations.
        if (typeof definition !== 'function') {
          return;
        }

        const prototype = definition.prototype;

        if (hookedDialogPrototypes.has(prototype)) {
          return;
        }

        const show = prototype.show;
        prototype.show = function () {
          const result = show.apply(this, arguments);

          if (isViewer()) {
            markFilterDialog(this);
          }

          return result;
        };
        hookedDialogPrototypes.add(prototype);
      });
    }

    function rememberWorksheet(worksheet) {
      if (!currentSession()) {
        const savedHistoryIndex = editorWindow.__skSavedHistoryIndex;
        let originalBytes;

        try {
          originalBytes = captureSave();
        } finally {
          editorWindow.__skSavedHistoryIndex = savedHistoryIndex;
        }

        session = {
          model: api.wb.model,
          originalBytes: originalBytes,
          originalPreview: global.OfficePreview.capture(editorWindow, 'xlsx'),
          worksheets: new Map()
        };
      }

      const model = worksheet.model;

      if (session.worksheets.has(model)) {
        return;
      }

      const hiddenRows = new Uint8Array(model.getRowsCount());

      for (let row = 0; row < hiddenRows.length; row++) {
        hiddenRows[row] = model.getRowHidden(row) ? 1 : 0;
      }

      session.worksheets.set(model, {
        worksheet: worksheet,
        autoFilter: cloneFilter(model.AutoFilter),
        tables: (model.TableParts || []).map(function (table) {
          return { table: table, autoFilter: cloneFilter(table.AutoFilter) };
        }),
        namedViews: (model.aNamedSheetViews || []).map(function (view) {
          return {
            view: view,
            filters: view.nsvFilters.map(function (filter) { return filter.clone(); })
          };
        }),
        hiddenRows: hiddenRows
      });
    }

    function withoutHistory(callback) {
      const removeLastPoint = history.Remove_LastPoint;
      const wasApplying = applyingFilter;
      applyingFilter = true;
      history.TurnOff();
      // Native filter transactions otherwise remove an existing empty undo point.
      history.Remove_LastPoint = function () {};

      try {
        return callback();
      } finally {
        history.Remove_LastPoint = removeLastPoint;
        history.TurnOn();
        applyingFilter = wasApplying;
      }
    }

    function restoreRows(model, hiddenRows) {
      const rowCount = Math.max(hiddenRows.length, model.getRowsCount());
      let row = 0;

      while (row < rowCount) {
        const wasHidden = hiddenRows[row] === 1;

        if (!!model.getRowHidden(row) === wasHidden) {
          row++;

          continue;
        }

        const firstRow = row++;

        while (row < rowCount && (hiddenRows[row] === 1) === wasHidden) {
          row++;
        }

        model.setRowHidden(wasHidden, firstRow, row - 1);
      }
    }

    function reset() {
      Array.from(openFilterDialogs).forEach(function (dialog) {
        dialog.close();
      });
      const localSession = currentSession();

      if (!localSession) {
        return;
      }

      withoutHistory(function () {
        localSession.worksheets.forEach(function (state, model) {
          model.AutoFilter = cloneFilter(state.autoFilter);
          state.tables.forEach(function (entry) {
            entry.table.AutoFilter = cloneFilter(entry.autoFilter);
          });
          state.namedViews.forEach(function (entry) {
            entry.view.nsvFilters = entry.filters.map(function (filter) {
              return filter.clone();
            });
          });
          restoreRows(model, state.hiddenRows);

          const range = new editorWindow.Asc.Range(
            0, 0, Math.max(0, model.getColsCount() - 1), Math.max(0, model.getRowsCount() - 1)
          );
          const isInactiveSheet = model !== api.wb.getWorksheet().model;
          state.worksheet._onUpdateFormatTable(range, isInactiveSheet);
          state.worksheet._updateSlicers(range);
          state.worksheet.handlers.trigger('onFilterInfo');
        });
      });

      session = null;
      refreshControls();
    }

    function runLocalFilter(worksheet, callback) {
      const selection = worksheet.model.selectionRange.getLast();
      const isPivotSelection = worksheet.model.inPivotTable(selection);
      const isBusy = history.Transaction !== 0 || applyingFilter;

      if (isPivotSelection || isBusy) {
        return false;
      }

      rememberWorksheet(worksheet);
      const canEdit = api.canEdit;
      const filteringMode = excel.filteringMode;
      const lockAll = worksheet._isLockedAll;
      const lockCells = worksheet._isLockedCells;
      const lockNames = worksheet._isLockedDefNames;

      // These native operations are synchronous for plain AutoFilters. Allow only
      // this call, without changing the editor restriction or acquiring edit locks.
      api.canEdit = function () {
        return true;
      };
      excel.filteringMode = true;
      worksheet._isLockedAll = function (callback) {
        callback(true);
      };
      worksheet._isLockedCells = function (range, subType, callback) {
        callback(true);
      };
      worksheet._isLockedDefNames = function () {};

      try {
        return withoutHistory(callback);
      } finally {
        api.canEdit = canEdit;
        excel.filteringMode = filteringMode;
        worksheet._isLockedAll = lockAll;
        worksheet._isLockedCells = lockCells;
        worksheet._isLockedDefNames = lockNames;
        refreshControls();
      }
    }

    ['asc_addAutoFilter', 'asc_changeAutoFilter', 'asc_applyAutoFilter', 'asc_clearFilter',
      'asc_clearFilterColumn'].forEach(function (methodName) {
      const original = api[methodName];

      api[methodName] = function () {
        const args = arguments;

        if (!isViewer()) {
          return original.apply(this, args);
        }

        const isTableCreation = methodName === 'asc_addAutoFilter' && (args[0] != null || args[1] != null);
        const isPivotFilter = methodName === 'asc_applyAutoFilter' && args[0] && args[0].pivotObj;
        const isTableStyleChange = methodName === 'asc_changeAutoFilter' &&
          (args[1] !== editorWindow.Asc.c_oAscChangeFilterOptions.filter || typeof args[2] !== 'boolean');

        if (isTableCreation || isPivotFilter || isTableStyleChange) {
          return false;
        }

        const worksheet = methodName === 'asc_clearFilter' && args[0] ? api.wb.getWorksheetById(args[0]) : api.wb.getWorksheet();

        if (!worksheet) {
          return false;
        }

        return runLocalFilter(worksheet, function () {
          return original.apply(api, args);
        });
      };
    });

    const setRestriction = api.asc_setRestriction;
    api.asc_setRestriction = function (restriction) {
      if (restriction === restrictions.None) {
        reset();
      }

      const result = setRestriction.apply(this, arguments);
      refreshControls();

      return result;
    };

    editorWindow.__captureSave = function () {
      const localSession = currentSession();

      if (!localSession) {
        return captureSave.apply(this, arguments);
      }

      editorWindow.__skSavedHistoryIndex = history.Index;

      return localSession.originalBytes.slice();
    };

    const handlers = api.wb.handlers;
    const trigger = handlers.trigger;
    handlers.trigger = function (eventName, config) {
      if (applyingFilter && eventName === 'asc_onDocumentModifiedChanged') {
        return;
      }

      if (isViewer() && eventName === 'asc_onSetAFDialog') {
        if (config.asc_getPivotObj()) {
          return;
        }

        const holder = controllers.DocumentHolder;
        const wasEdit = holder.permissions.isEdit;

        try {
          holder.permissions.isEdit = true;
          holder.onApiAutofilter(config);
        } finally {
          holder.permissions.isEdit = wasEdit;
        }

        const dialog = holder.dlgFilter;

        if (dialog) {
          hookFilterDialogs();
          markFilterDialog(dialog);
          // Sorting changes cell data; only filtering is available in view mode.
          ['miSortLow2High', 'miSortHigh2Low', 'miSortOptions', 'miSortCellColor', 'miSortFontColor', 'miReapply', 'miReapplySeparator']
              .forEach(function (name) {
                dialog[name].setVisible(false);
              });
        }

        return;
      }

      const result = trigger.apply(this, arguments);

      if (eventName === 'asc_onSelectionChanged') {
        refreshControls();
      }

      return result;
    };

    function refreshControls() {
      const controller = controllers.Toolbar;
      const toolbar = controller && controller.toolbar;

      if (!toolbar) {
        return;
      }

      if (toolbar !== hookedToolbar) {
        toolbar.on('tab:active change:compact', refreshControls);
        hookedToolbar = toolbar;
      }

      if (!document.getElementById('sk-view-filter-style')) {
        const style = document.createElement('style');
        style.id = 'sk-view-filter-style';
        style.textContent = [
          '.toolbar.masked .group.sk-view-filter-group:not(.no-mask){opacity:1;}',
          '.toolbar.masked .sk-view-filter-group .btn:not(.sk-view-filter-control){opacity:.4;}',
          '.toolbar.masked .sk-view-filter-group .sk-view-filter-control:disabled{opacity:.4;}',
          '.toolbar.masked .sk-view-filter-control{position:relative;z-index:1;}'
        ].join('\n');
        document.head.appendChild(style);
      }

      document.querySelectorAll('.sk-view-filter-group').forEach(function (group) {
        group.classList.remove('sk-view-filter-group');
      });

      const viewer = isViewer();
      const info = api.asc_getCellInfo();
      const selectionTypes = editorWindow.Asc.c_oAscSelectionType;
      const isCellSelection = info && [
          selectionTypes.RangeCells,
          selectionTypes.RangeCol,
          selectionTypes.RangeRow,
          selectionTypes.RangeMax
      ].includes(info.asc_getSelectionType());
      const filterInfo = info && info.asc_getAutoFilterInfo();
      const worksheet = api.wb.getWorksheet();
      const isProtected = worksheet.model.getSheetProtection(editorWindow.Asc.c_oAscSheetProtectType.autoFilter);
      const canFilter = isCellSelection && filterInfo && filterInfo.asc_getIsAutoFilter() !== null && !isProtected && !history.Transaction;
      const hasAppliedFilter = filterInfo && filterInfo.asc_getIsApplyAutoFilter() === true;

      function updateButtons(buttons, action) {
        (buttons || []).forEach(function (button) {
          const element = button.cmpEl && button.cmpEl[0];

          if (!element) {
            return;
          }

          element.classList.toggle('sk-view-filter-control', viewer);

          if (!viewer) {
            element.removeAttribute('data-sk-view-filter-action');
            element.removeAttribute('data-sk-view-filter');

            return;
          }

          element.setAttribute('data-sk-view-filter-action', action);
          element.setAttribute('data-sk-view-filter', '');
          const group = element.closest('.group');

          if (group) {
            group.classList.add('sk-view-filter-group');
          }

          const disabled = !canFilter || (action === 'clear' && !hasAppliedFilter);
          button.setDisabled(disabled);

          if (action === 'toggle') {
            button.toggle(filterInfo && filterInfo.asc_getIsAutoFilter() === true, true);
          }
        });
      }

      updateButtons(toolbar.btnsSetAutofilter, 'toggle');
      updateButtons(toolbar.btnsClearAutofilter, 'clear');
    }

    document.addEventListener('click', function (event) {
      const button = event.target.closest('[data-sk-view-filter-action]');

      if (!isViewer() || !button || button.disabled) {
        return;
      }

      event.preventDefault();
      event.stopImmediatePropagation();

      if (button.getAttribute('data-sk-view-filter-action') === 'clear') {
        api.asc_clearFilter();

        return;
      }

      const info = api.asc_getCellInfo();
      const filterInfo = info.asc_getAutoFilterInfo();
      const tableInfo = info.asc_getFormatTableInfo();
      const hasFilter = filterInfo && filterInfo.asc_getIsAutoFilter() === true;

      if (tableInfo || hasFilter) {
        api.asc_changeAutoFilter(tableInfo && tableInfo.asc_getTableName(), editorWindow.Asc.c_oAscChangeFilterOptions.filter, !hasFilter);
      } else {
        api.asc_addAutoFilter();
      }
    }, true);

    api.__skViewFilters = {
      reset: reset,
      refresh: refreshControls,
      capturePreview: function () {
        const localSession = currentSession();

        return localSession ? localSession.originalPreview : global.OfficePreview.capture(editorWindow, 'xlsx');
      }
    };

    refreshControls();
  }

  global.OfficeSheetFilters = { install: install };
})(window);
