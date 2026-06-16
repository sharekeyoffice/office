// wrapper-boot.js — runs INSIDE the editor iframe, before app.js
//
// Purpose:
//   1. License shim: make asc_getCanBranding() return true so
//      Common.UI.LayoutManager honours editorConfig.customization.layout
//      (per 10.0.2 finding — without this, the layout tree is silently ignored).
//      Also asc_getCustomization() → true so Main.js calls
//      Header.setBranding(customization) — that's what swaps the OnlyOffice
//      header logo for our customization.logo (Sharekey). Required for the
//      ONLYOFFICE trademark policy (brand must read "Sharekey").
//   2. SDK loader hints: same as the existing v0 viewer pattern.
//
// Injected into web-apps/deploy/.../index.html by inject-boot.sh after each
// grunt build. Must execute BEFORE require(['app']) in the deployed bootstrap.
//
// Idempotent — running twice is a no-op (we check before patching).

(function () {
  'use strict';

  // ---- License shim ----------------------------------------------------
  // Targets sdkjs/common/apiCommon.js:493 — the asc_CAscEditorPermissions
  // prototype. Stable cross-release API.
  function shimCanBranding() {
    if (!window.Asc || !Asc.asc_CAscEditorPermissions || !Asc.asc_CAscEditorPermissions.prototype) {
      return false;
    }
    var P = Asc.asc_CAscEditorPermissions.prototype;
    if (P.__wrapperShimmed) return true;
    P.asc_getCanBranding = function () { return true; };
    // Gates Main.js's `if (canBranding) headerView.setBranding(customization)`
    // (canBranding = asc_getCustomization()). Without this the header logo stays
    // the default OnlyOffice mark and customization.logo is ignored.
    P.asc_getCustomization = function () { return true; };
    P.__wrapperShimmed = true;
    return true;
  }

  // The Asc namespace is created by sdk-all.js, which loads asynchronously
  // via requirejs. Try once now, otherwise poll until it appears or until
  // a sane timeout (matches the editor's own require timeout).
  if (!shimCanBranding()) {
    var waited = 0;
    var iv = setInterval(function () {
      if (shimCanBranding() || (waited += 100) >= 30000) {
        clearInterval(iv);
      }
    }, 100);
  }

  // ---- SDK loader hints (mirror v0 viewer.html pattern) ----------------
  // AscNotLoadAllScript — prevents loadSdk() from fetching the nonexistent
  // sdk-all.js bundle when running in our trimmed deploy.
  // Asc.Addons.ooxml — enables the OOXML branch (does nothing useful
  // without sdkjs-ooxml addon, but prevents fallback errors).
  window.AscNotLoadAllScript = true;
  window.Asc = window.Asc || {};
  window.Asc.Addons = window.Asc.Addons || {};
  window.Asc.Addons.ooxml = true;

  // ---- Suppress OnlyOffice's "connection is too slow" tooltip --------
  // The deployed index.html sets a 30s setTimeout that pops a require-timeout
  // alert AND auto-reloads the page. Two harmless require-registry entries
  // (the bundle entry-point + es6-promise) keep the timer's clearTimeout from
  // firing even though the editor is fully booted. We:
  //   1. Cancel the timer if it's been scheduled (var-at-top-of-script-tag
  //      becomes a window property in browsers).
  //   2. Override the message function so any later trigger renders empty.
  //   3. Re-cancel after a short delay in case it was scheduled after us.
  function killRequireTimeout() {
    if (window.requireTimeoutID) {
      clearTimeout(window.requireTimeoutID);
      window.requireTimeoutID = 0;
    }
  }
  window.requireTimeourError = function () { return null; };
  killRequireTimeout();
  setTimeout(killRequireTimeout, 100);
  setTimeout(killRequireTimeout, 1000);
})();
