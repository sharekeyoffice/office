/**
 * stubs.js — no-op replacements for SDK classes the viewer doesn't need.
 *
 * The SDK's `_onEndLoadSdk` lifecycle and various constructors eagerly
 * instantiate ~15 classes that exist only to power editing / collaboration /
 * macros / plugin host / builder API. Round 1 of slimming proved we can't
 * just drop those files — every constructor reference must resolve to *some*
 * class, even if its methods do nothing.
 *
 * This file defines the minimum surface each class needs so the viewer's
 * boot + open path runs to completion. Loaded BEFORE the SDK manifest, so
 * SDK files that try to add real implementations later just overwrite our
 * stubs (no-op for files we then exclude from the manifest).
 *
 * See analysis/plan/exclusions/README.md for the audit trail of which file
 * each stub corresponds to and why.
 */
(function (window) {
  'use strict';

  window.AscCommon = window.AscCommon || {};
  window.AscFonts  = window.AscFonts  || {};
  window.Asc       = window.Asc       || {};

  function noop() {}
  function returnFalse() { return false; }
  function returnEmptyArr() { return []; }
  function returnEmptyObj() { return {}; }

  // -----------------------------------------------------------------
  // AscBuilder.{Word, Cell, Slide} — 2.5 MB combined when un-stubbed
  // Used: `_onEndLoadSdk` → `AscBuilder.<Editor>.init()`; getters return `.Api`.
  // -----------------------------------------------------------------
  window.AscBuilder = window.AscBuilder || {};
  ['Word', 'Cell', 'Slide'].forEach(function (E) {
    window.AscBuilder[E] = window.AscBuilder[E] || { init: noop, Api: {} };
  });

  // -----------------------------------------------------------------
  // MacroRecorder — apiBase.js:210 instantiates one per editor.
  // Methods called: stop(), cancel(), addStepData(type, data).
  // -----------------------------------------------------------------
  window.AscCommon.MacroRecorder = window.AscCommon.MacroRecorder || function () {
    this.stop = noop;
    this.cancel = noop;
    this.addStepData = noop;
  };

  // -----------------------------------------------------------------
  // CDocumentMacros — apiBase.js:3316 (during _onEndLoadSdk).
  // Methods checked: isExistAuto(), CheckLock(), SetData(), GetData(),
  //                  Get_Id(), runAuto(), run(), getAllNames(), getGuidByName(),
  //                  getNameByGuid().
  // For viewer: isExistAuto must return false so auto-run gates skip.
  // -----------------------------------------------------------------
  window.AscCommon.CDocumentMacros = window.AscCommon.CDocumentMacros || function () {
    this.Data = '';
    this.isExistAuto = returnFalse;
    this.CheckLock = noop;
    this.SetData = noop;
    this.GetData = function () { return this.Data; };
    this.Get_Id = function () { return 'macros-stub'; };
    this.Write_ToBinary2 = noop;
    this.Read_FromBinary2 = noop;
    this.Refresh_RecalcData = noop;
    this.runAuto = noop;
    this.run = noop;
    this.getAllNames = returnEmptyArr;
    this.getGuidByName = function () { return null; };
    this.getNameByGuid = function () { return null; };
  };

  // -----------------------------------------------------------------
  // Plugin host — apiBase_plugins.js (95 KB) + per-editor api_plugins.js
  // (~75 KB combined). `apiBase.js:3314` calls `Asc.createPluginsManager(this)`
  // unconditionally during _onEndLoadSdk; the factory lives in
  // apiBase_plugins.js. Methods called on the returned manager: register, run,
  // close, runResize, buttonClick, onEnableMouseEvents, isWorked,
  // onPluginWindowDockChanged, pluginsMap[guid] property reads.
  // -----------------------------------------------------------------
  function makePluginsManagerStub() {
    return {
      pluginsMap: {},
      register: noop,
      run: noop,
      close: noop,
      runResize: noop,
      buttonClick: noop,
      onEnableMouseEvents: noop,
      onPluginWindowDockChanged: noop,
      isWorked: returnFalse,
      isRunned: returnFalse,
      getRunningPlugins: returnEmptyArr,
      load: noop,
      unload: noop,
      executeMethod: noop,
      tryToShowPlugin: returnFalse,
      // also a direct apiBase ref expected by a few helpers
      onPluginEvent: noop,
      onSendThemeColors: noop
    };
  }
  window.Asc.createPluginsManager = window.Asc.createPluginsManager || function () {
    return makePluginsManagerStub();
  };

  // CPluginCtxMenuInfo is instantiated at apiBase.js:5317 (context-menu paths).
  window.AscCommon.CPluginCtxMenuInfo = window.AscCommon.CPluginCtxMenuInfo || function () {};

  // pluginMethod_* — apiBase_plugins.js mixes these into
  // baseEditorsApi.prototype. Outside the plugin files, only two are called
  // from the open lifecycle:
  //   - pluginMethod_SetProperties     (apiBase.js:1513, onDocumentContentReady)
  //   - pluginMethod_GetCustomFunctions (cell formula path)
  // Stubs.js runs BEFORE the SDK so baseEditorsApi doesn't exist yet — install
  // the methods lazily once AscCommon.baseEditorsApi appears.
  window.__viewerInstallPluginMethodStubs = function () {
    var P = window.AscCommon && window.AscCommon.baseEditorsApi && window.AscCommon.baseEditorsApi.prototype;
    if (!P) return false;
    if (P.__viewerPluginMethodStubsInstalled) return true;
    P['pluginMethod_SetProperties']      = noop;
    P['pluginMethod_GetCustomFunctions'] = function () { return '[]'; };
    // cell/api_plugins.js mixin: cell/api.js:3421 calls this on
    // _openDocumentEndCallback. Without the plugin file, the method is missing.
    P['registerCustomFunctionsLibrary'] = noop;
    // Defensive catch-all: any other pluginMethod_* call gets a logged no-op.
    var handler = {
      get: function (target, prop) {
        if (typeof prop === 'string' && prop.indexOf('pluginMethod_') === 0 &&
            target[prop] === undefined) {
          console.warn('[stubs] missing ' + prop + ' — returning no-op');
          return noop;
        }
        return target[prop];
      }
    };
    // We can't Proxy the prototype itself (would break SDK's instanceof),
    // so we just install the two known methods. New misses will console.warn
    // via the catch in the caller (but that may not exist) — manual probe
    // is good enough for the viewer.
    P.__viewerPluginMethodStubsInstalled = true;
    console.log('[stubs] installed pluginMethod_* stubs on baseEditorsApi.prototype');
    return true;
  };

  console.log('[stubs] viewer stubs installed:',
    'AscBuilder.{Word,Cell,Slide},',
    'AscCommon.{MacroRecorder, CDocumentMacros}');

})(window);
