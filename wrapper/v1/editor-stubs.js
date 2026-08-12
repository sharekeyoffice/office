// editor-stubs.js — DocServer-emulating stub layer for edit-mode boot.
//
// Loaded INSIDE the editor iframe before app.js. Patches the SDK so the
// editor's normal load chain (onEditorPermissions → asc_LoadDocument →
// CoAuth → applyModeEditorElements → onDocumentContentReady) completes
// without a real DocumentServer.
//
// Each section addresses one cascade layer documented in
// analysis/plan/phases/04-editing-save.md. When a layer is fixed,
// the next layer becomes visible. We add stubs ONLY for what's needed
// — minimal, comment-annotated.
//
// Idempotent: running twice is a no-op (we mark patches with a flag).

(function () {
  'use strict';

  if (window.__editorStubsLoaded) return;
  window.__editorStubsLoaded = true;

  // ---------------------------------------------------------------------
  // Phase 4.2 polish — font picker fallback for editor mode.
  //
  // The editor's name-keyed font lookup falls through to ASCW3 (symbol-only
  // sentinel) for any font name not in our `common/AllFonts.js` exact list
  // — e.g. docs that say "Calibri" never resolve to our Carlito files
  // because GetFontIndex penalty-scores against an empty selection_bin.
  // Result: every glyph renders as a tofu box.
  //
  // Same fix as v0 viewer.html: prime FontPickerMap from g_font_infos and
  // wrap GetFontFileWeb so unknown names route through FONT_FALLBACK_MAP.
  //
  // Plus: __plainFonts must be true BEFORE any font fetch so the XOR-decode
  // step in CFontFileLoader.LoadFontArrayBuffer is skipped for our plain
  // TTFs. Setting it here is the earliest possible point in the iframe.
  // ---------------------------------------------------------------------
  window.__plainFonts = true;

  // Map of "we don't have this font name → use this real font name instead".
  // The picker shim consults this when a name isn't found in g_font_infos;
  // misses fall back to FONT_FALLBACK_DEFAULT.
  //
  // Aptos / Aptos Display / Aptos Narrow are the new Microsoft Office 2024
  // defaults that replaced Calibri. Bahnschrift, Sitka, Source Sans Pro,
  // Yu Gothic etc. are common in modern Office docs. All sans-serifs route
  // to Carlito (Calibri-metric-compatible), serifs to Liberation Serif /
  // Caladea (Cambria-metric-compatible), monospace to Liberation Mono.
  var FONT_FALLBACK_MAP = {
    'Helvetica':         'Arial',
    'Calibri':           'Carlito',
    'Calibri Light':     'Carlito',
    'Aptos':             'Carlito',
    'Aptos Display':     'Carlito',
    'Aptos Narrow':      'Carlito',
    'Aptos Mono':        'Liberation Mono',
    'Aptos SemiBold':    'Carlito',
    'Aptos Light':       'Carlito',
    'Cambria':           'Caladea',
    'Cambria Math':      'Caladea',
    'Tahoma':            'DejaVu Sans',
    'Verdana':           'DejaVu Sans',
    'Segoe UI':          'DejaVu Sans',
    'Segoe UI Light':    'DejaVu Sans',
    'Segoe UI Semibold': 'DejaVu Sans',
    'Segoe Print':       'DejaVu Sans',
    'Segoe Script':      'DejaVu Sans',
    'Bahnschrift':       'DejaVu Sans',
    'Trebuchet MS':      'DejaVu Sans',
    'Georgia':           'Liberation Serif',
    'Sitka Text':        'Liberation Serif',
    'Sitka Display':     'Liberation Serif',
    'Sitka Banner':      'Liberation Serif',
    'Sitka Heading':     'Liberation Serif',
    'Sitka Subheading':  'Liberation Serif',
    'Sitka Small':       'Liberation Serif',
    'Source Sans Pro':   'Liberation Sans',
    'Source Serif Pro':  'Liberation Serif',
    'Open Sans':         'Liberation Sans',
    'Roboto':            'Liberation Sans',
    'Roboto Mono':       'Liberation Mono',
    'Consolas':          'Liberation Mono',
    'Cascadia Code':     'Liberation Mono',
    'Cascadia Mono':     'Liberation Mono',
    'Lucida Console':    'Liberation Mono',
    'Symbol':            'OpenSymbol',
    'Wingdings':         'OpenSymbol',
    'Wingdings 2':       'OpenSymbol',
    'Wingdings 3':       'OpenSymbol',
    'Webdings':          'OpenSymbol'
  };
  var FONT_FALLBACK_DEFAULT = 'DejaVu Sans';

  // Register the FONT_FALLBACK_MAP entries as REAL aliases in g_font_infos
  // (and g_map_font_index), pointing at the same CFontInfo as their target.
  // Why this matters for cell specifically:
  //   The cell renderer caches a font INDEX per cell (numeric index into
  //   g_font_infos) at open time. If "Aptos Narrow" isn't a key in
  //   g_map_font_index, the cached index is undefined → the per-cell font
  //   lookup returns garbage → text renders as .notdef boxes (tofu).
  //
  // The picker shim alone isn't enough because it intercepts NAME->FILE
  // lookups, not NAME->INDEX. By aliasing Aptos's index to Carlito's,
  // the index-keyed pipelines also work.
  //
  // Idempotent: if g_map_font_index already has the alias name, skip.
  function registerFontAliasesOnce() {
    var AscFonts = window.AscFonts;
    if (!AscFonts || !AscFonts.g_font_infos || !AscFonts.g_map_font_index) return false;
    var infos = AscFonts.g_font_infos;
    var indexMap = AscFonts.g_map_font_index;
    if (infos.length < 2) return false;
    if (window.__fontAliasesRegistered) return true;
    var added = 0;
    Object.keys(FONT_FALLBACK_MAP).forEach(function (alias) {
      if (alias in indexMap) return;  // real font with that name exists already
      var targetIdx = indexMap[FONT_FALLBACK_MAP[alias]];
      if (targetIdx === undefined) return;  // target not loaded — skip silently
      // Re-use the CFontInfo at targetIdx so the file pointers match. The
      // SDK's CheckFontLoadStyles consults info.indexR/indexI/etc., which
      // already point at real .ttf files for the target font.
      indexMap[alias] = targetIdx;
      added++;
    });
    window.__fontAliasesRegistered = true;
    if (added) log('registered ' + added + ' font alias(es) in g_map_font_index');
    return true;
  }

  function primeFontPickerOnce() {
    if (!window.AscFonts || !window.AscFonts.g_fontApplication) return false;
    var app = window.AscFonts.g_fontApplication;
    if (!app.FontPickerMap) return false;
    var infos = window.AscFonts.g_font_infos || [];
    if (infos.length < 2) return false;  // manifest not yet populated

    // Register aliases first so the loops below see them.
    registerFontAliasesOnce();

    for (var i = 0; i < infos.length; i++) {
      var name = infos[i].Name;
      if (!name || name === 'ASCW3') continue;
      app.FontPickerMap[name] = { m_wsFontName: name };
    }
    // Also seed FontPickerMap with the alias names → their target font
    // names. Same reasoning as above: ensures consistent name-keyed
    // resolution in code paths that bypass the GetFontFileWeb shim.
    //
    // CRITICAL: overwrite existing entries that resolve to ASCW3 or to a
    // missing font. The SDK's own GetFontFileWeb may have populated the
    // map BEFORE our shim ran (e.g. during the workbook's font-table
    // preload step) and cached "Aptos Narrow" → ASCW3 (its no-glyphs
    // sentinel). If we leave that, every Aptos cell renders as tofu.
    var aliasIndexMap = window.AscFonts.g_map_font_index;
    Object.keys(FONT_FALLBACK_MAP).forEach(function (alias) {
      var existing = app.FontPickerMap[alias];
      var bad = !existing || existing.m_wsFontName === 'ASCW3' ||
                aliasIndexMap[existing.m_wsFontName] === undefined;
      if (bad) {
        app.FontPickerMap[alias] = { m_wsFontName: FONT_FALLBACK_MAP[alias] };
      }
    });
    if (!app.__viewerFallbackInstalled) {
      var orig = app.GetFontFileWeb.bind(app);
      // Helper: a cached picker entry is "good" only if its target font
      // exists in g_map_font_index AND isn't the ASCW3 sentinel (the
      // symbol-only no-glyphs fallback the SDK uses when its own picker
      // can't resolve a name). Trusting ASCW3 means rendering tofu.
      function isGoodCachedEntry(r) {
        return r && r.m_wsFontName && r.m_wsFontName !== 'ASCW3' &&
               window.AscFonts.g_map_font_index[r.m_wsFontName] !== undefined;
      }
      app.GetFontFileWeb = function (name, lStyle) {
        var r = app.FontPickerMap[name];
        // Don't shortcut on bad cache entries — re-resolve. Without this
        // check, an ASCW3 cached BEFORE our shim was installed (e.g. by
        // the SDK's own GetFontFileWeb during font preload) sticks around
        // forever and every cell rendered with that font shows tofu.
        if (isGoodCachedEntry(r)) return r;
        r = orig(name, lStyle);
        if (!isGoodCachedEntry(r)) {
          var target = FONT_FALLBACK_MAP[name] || FONT_FALLBACK_DEFAULT;
          var sub = { m_wsFontName: target };
          app.FontPickerMap[name] = sub;
          return sub;
        }
        return r;
      };
      app.__viewerFallbackInstalled = true;
    }
    return true;
  }
  // Expose for debugging / re-priming if a host wants to force it.
  window.__primeFontPickerOnce = primeFontPickerOnce;

  // ---------------------------------------------------------------------
  // PROTOTYPE (pending legal sign-off) — render with the REAL Microsoft Core
  // Fonts (Arial / Times New Roman / Courier New) instead of the Liberation
  // metric clones, to match OnlyOffice cloud (which ships ttf-mscorefonts).
  //
  // Mechanism: we DON'T touch the manifest (name→index stays Arial→0, …) or the
  // bundle. We only change the FILE the loader fetches for each Liberation slot
  // — CFontFileLoader.Id is the filename appended to fontFilesPath. So
  // Arial→index 0→g_font_files[0], and flipping that file's Id from
  // 'LiberationSans-Regular.ttf' to 'arial.ttf' makes the engine fetch the real
  // Arial from /fonts/arial.ttf. The real fonts are NOT in git — they're fetched
  // at build time by build/fetch-mscorefonts.sh (Microsoft's installer, EULA).
  //
  // SK_REAL_MSCOREFONTS is toggled by that script (off in the committed source).
  //
  // ---------------------------------------------------------------------
  var SK_REAL_MSCOREFONTS = true;
  var SK_MSCORE_REPOINT = {
    'LiberationSans-Regular.ttf':    'arial.ttf',
    'LiberationSans-Italic.ttf':     'ariali.ttf',
    'LiberationSans-Bold.ttf':       'arialbd.ttf',
    'LiberationSans-BoldItalic.ttf': 'arialbi.ttf',
    'LiberationSerif-Regular.ttf':    'times.ttf',
    'LiberationSerif-Italic.ttf':     'timesi.ttf',
    'LiberationSerif-Bold.ttf':       'timesbd.ttf',
    'LiberationSerif-BoldItalic.ttf': 'timesbi.ttf',
    'LiberationMono-Regular.ttf':    'cour.ttf',
    'LiberationMono-Italic.ttf':     'couri.ttf',
    'LiberationMono-Bold.ttf':       'courbd.ttf',
    'LiberationMono-BoldItalic.ttf': 'courbi.ttf'
  };
  function ensureRealMsFontsOnce() {
    if (!SK_REAL_MSCOREFONTS) return true;            // prototype off → no-op
    if (window.__skMsFontsRepointed) return true;     // idempotent
    var AscFonts = window.AscFonts;
    if (!AscFonts || !AscFonts.g_font_files || !AscFonts.g_font_files.length) return false;
    var changed = 0;
    AscFonts.g_font_files.forEach(function (f) {
      if (!f) return;
      var real = SK_MSCORE_REPOINT[f.Id];   // f.Id is the .ttf filename used in the fetch URL
      if (real) {
        f.Id = real;        // next load fetches /fonts/<real> (the real MS font)
        f.Status = -1;      // mark notloaded so the engine (re)fetches the new bytes
        changed++;
      }
    });
    if (changed) {
      window.__skMsFontsRepointed = true;
      log('PROTOTYPE: repointed ' + changed + ' Liberation slot(s) → real MS Core Fonts');
    }
    return changed > 0;
  }
  window.__ensureRealMsFontsOnce = ensureRealMsFontsOnce;

  // ---------------------------------------------------------------------
  // Layer 6b (slide cascade) — defensive guard on CTextShaper.FlushWord.
  //
  // In slide's edit-mode boot, _openDocumentEndCallback's initial
  // Recalculate() can race the font loader: a paragraph using a theme
  // font (+mj-lt / +mn-lt) reaches FlushWord before the picker has had a
  // chance to materialise the font into g_oTextMeasurer.m_oManager.m_pFont.
  // The FlushWord body reads `this.FontId.m_pFaceInfo.family_name` —
  // FontId is null in that race, the read throws, the SDK reports
  // errorCode=-82 ("error opening file"), and isDocumentLoadComplete
  // never flips. The presentation model itself is fine; only the first-
  // pass recalc died.
  //
  // Word + cell don't hit this because their initial paragraphs already
  // reference fonts the SDK preloads via `IsNeedDefaultFonts` (Arial,
  // Symbol, Wingdings, Courier New, Times New Roman). Slide uses theme
  // fonts that bypass that preload list.
  //
  // Patch: when FontId is missing, fall back to whatever font the
  // measurer last loaded (m_oManager.m_pFont). If that's also null,
  // bail cleanly — the buffer reset is still safe and the next recalc
  // (after fonts settle) renders the text properly. We mirror the
  // assignment FlushWord makes when it succeeds, so subsequent shaping
  // sees a non-null FontId.
  // ---------------------------------------------------------------------
  function patchTextShaper() {
    // CTextShaper lives on window.AscFonts (see end of common/libfont/textshaper.js).
    var TS = window.AscFonts && window.AscFonts.CTextShaper;
    if (!TS || !TS.prototype || TS.prototype.__shapeGuardInstalled) return false;
    var origFlush = TS.prototype.FlushWord;
    TS.prototype.FlushWord = function () {
      // The crash path reads `this.FontId.m_pFaceInfo.family_name`.
      // Guard both: FontId may be null, or .m_pFaceInfo may be null
      // (font loaded but face metadata not yet parsed).
      if (!this.FontId || !this.FontId.m_pFaceInfo) {
        var measurer = window.AscCommon && window.AscCommon.g_oTextMeasurer;
        var fallback = measurer && measurer.m_oManager && measurer.m_oManager.m_pFont;
        if (fallback && fallback.m_pFaceInfo) {
          this.FontId = fallback;
        } else {
          // Truly nothing usable yet — drop the buffer instead of crashing.
          // Subsequent passes (after fonts settle) will reshape properly.
          if (this.ClearBuffer) this.ClearBuffer();
          return;
        }
      }
      return origFlush.apply(this, arguments);
    };
    TS.prototype.__shapeGuardInstalled = true;
    log('CTextShaper.FlushWord guarded');
    return true;
  }

  if (!patchTextShaper()) {
    var triesTS = 0;
    var ivTS = setInterval(function () {
      if (patchTextShaper() || ++triesTS > 300) clearInterval(ivTS);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 6c (slide cascade) — defensive guard on CTextMeasurer.MeasureCode.
  //
  // After 6b unblocked the initial body-text recalc, slide hits a second
  // race during layout-thumbnail rendering: the image-loaded callback path
  //   asyncImageEndLoadedBackground → CDrawingDocument.CheckRasterImageOnScreen
  //   → CEditorPage.CheckLayouts → CLayoutThumbnailDrawer.GetThumbnail
  //   → MasterSlide.recalculate → CParagraphRecalculateStateWrap.Recalculate_Numbering
  //   → CPresentationBullet.Measure → CTextMeasurer.MeasureCode
  // calls `m_oManager.MeasureChar(unicode)`, which returns undefined when
  // the bullet's font face isn't yet available (the bullet uses a
  // theme-resolved font, separate from body text). MeasureCode then reads
  // `Temp.fAdvanceX` on undefined and crashes; SDK reports errorCode=-82.
  //
  // The presentation model is already populated and intact — only the
  // first-pass thumbnail measurement died on this race. Layer 6b doesn't
  // catch this because it guards the SHAPER (FlushWord), and MeasureCode
  // is a different code path (numbering layout, no shaper involvement).
  //
  // Patch: when MeasureChar returns falsy, return a zero-size measurement
  // so the bullet lays out as zero-width (cosmetic only, on a thumbnail
  // that's about to be re-rendered anyway). Preserves the same return
  // shape so callers see {Width:0, Height:0, Ascent:0} instead of throwing.
  // ---------------------------------------------------------------------
  function patchTextMeasurer() {
    var TM = window.AscCommon && window.AscCommon.g_oTextMeasurer;
    if (!TM || TM.__measureGuardInstalled) return false;
    var origMeasure = TM.MeasureCode;
    TM.MeasureCode = function (lUnicode) {
      try {
        return origMeasure.call(this, lUnicode);
      } catch (e) {
        // Most likely cause: m_oManager.MeasureChar returned undefined
        // because the requested font face isn't loaded yet. Bail with a
        // benign zero-size measurement; the next recalc reshapes properly.
        return { Width: 0, Height: 0, Ascent: 0 };
      }
    };
    // Same guard for Measure2Code (used in the same numbering / shaping
    // paths and reads .fAdvanceX off the same MeasureChar return).
    if (typeof TM.Measure2Code === 'function') {
      var origMeasure2 = TM.Measure2Code;
      TM.Measure2Code = function (lUnicode) {
        try {
          return origMeasure2.call(this, lUnicode);
        } catch (e) {
          return { Width: 0 };
        }
      };
    }
    TM.__measureGuardInstalled = true;
    log('CTextMeasurer.MeasureCode guarded');
    return true;
  }

  if (!patchTextMeasurer()) {
    var triesTM = 0;
    var ivTM = setInterval(function () {
      if (patchTextMeasurer() || ++triesTM > 300) clearInterval(ivTM);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 6d (slide cascade) — defensive guard on ParaRun.AddText.
  //
  // Slide hits a third race during layout-thumbnail rendering. Trigger:
  //   asyncImageEndLoadedBackground → CheckRasterImageOnScreen
  //     → CheckLayouts → CLayoutThumbnailDrawer.GetThumbnail
  //     → SlideLayout.recalculate2 → CShape2.recalculateContent2
  //     → CreateDocContentFromString → AddToContentFromString
  //     → ParaRun.AddText(undefined)
  //
  // The slide editor synthesises preview text for empty layout placeholders
  // ("Click to add title", "Click to add text" — for the New-Slide picker
  // panel). The prompt string comes from a localization map that may
  // resolve to `undefined` for some placeholder types (e.g. master slide
  // placeholders with phType the localization table doesn't have an entry
  // for). The SDK's ParaRun.AddText reads `sString.getUnicodeIterator()`
  // without a null check → TypeError → SDK reports errorCode=-82, the
  // open path bails, isDocumentLoadComplete never flips.
  //
  // The actual slide content is fine — the crash is during the LAYOUTS
  // PANEL preview render (the picker shown when you click "New Slide" or
  // change layout). Skipping the synthesis when there's no prompt text
  // just leaves that placeholder empty in the thumbnail; main slide
  // rendering is untouched.
  //
  // Word + cell don't hit this because they don't have a layouts panel
  // with on-the-fly thumbnail synthesis.
  // ---------------------------------------------------------------------
  function patchParaRunAddText() {
    var ParaRun = window.AscWord && window.AscWord.ParaRun;
    if (!ParaRun || !ParaRun.prototype || ParaRun.prototype.__addTextGuardInstalled) return false;
    var origAddText = ParaRun.prototype.AddText;
    ParaRun.prototype.AddText = function (sString, nPos) {
      // The crash path reads `sString.getUnicodeIterator()`. Guard for
      // null/undefined/non-string. An empty string is harmless — the loop
      // inside the original function exits immediately.
      if (sString == null || typeof sString !== 'string') {
        return; // skip — leaves the run empty, layout thumbnail renders
                // the placeholder shape as a blank box
      }
      return origAddText.call(this, sString, nPos);
    };
    ParaRun.prototype.__addTextGuardInstalled = true;
    log('ParaRun.AddText guarded');
    return true;
  }

  if (!patchParaRunAddText()) {
    var triesAR = 0;
    var ivAR = setInterval(function () {
      if (patchParaRunAddText() || ++triesAR > 300) clearInterval(ivAR);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 5b — Common.Locale stub so onLanguageLoaded() returns true.
  //
  // Each editor's Main.onEditorPermissions has an early-return guarded by
  // `!this.onLanguageLoaded()`, which checks
  // `Common.Locale.getCurrentLanguage()`. Without DocServer's locale init
  // step, that returns falsy → onEditorPermissions bails BEFORE setting
  // appOptions.isEdit / canLicense / customization → downstream callbacks
  // (onDocumentReady, setApi, reloadTranslations) crash on undefined
  // appOptions.X.
  //
  // The wrapper passes `editorConfig.lang` ('en' by default), so the
  // language IS configured — it just hasn't propagated to Common.Locale.
  // Stub Common.Locale to surface that lang directly. Idempotent.
  //
  // We also belt-and-braces force-populate the critical appOptions keys
  // after our synthetic onEditorPermissions, so even if some other
  // early-return path fires later we don't get half-initialized state.
  // (See patchMainLoadBinary below.)
  // ---------------------------------------------------------------------
  function patchCommonLocale() {
    var Common = window.Common;
    if (!Common || !Common.Locale) return false;
    if (Common.Locale.__langStubbed) return true;
    var configuredLang = (window.editor && window.editor.editorConfig && window.editor.editorConfig.lang) || 'en';
    var origGet = Common.Locale.getCurrentLanguage;
    Common.Locale.getCurrentLanguage = function () {
      var real = origGet ? origGet.call(this) : null;
      return real || configuredLang;
    };
    Common.Locale.__langStubbed = true;
    log('Common.Locale.getCurrentLanguage stubbed → "' + configuredLang + '"');
    return true;
  }

  if (!patchCommonLocale()) {
    var triesCL = 0;
    var ivCL = setInterval(function () {
      if (patchCommonLocale() || ++triesCL > 300) clearInterval(ivCL);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 5c — defensive guard on Common.UI.Button.updateHint.
  //
  // The slide editor's controller `onAppReady` iterates over its toolbar
  // buttons and calls `button.updateHint(localizedString)`. For some
  // buttons the localized string is undefined (we don't ship the full
  // localization map). Button.updateHint then hits this line:
  //
  //     $btn.attr('aria-label', (typeof hint == 'string') ? hint : hint[0]);
  //
  // — which reads `undefined[0]` when `hint` is null/undefined → throws
  // → propagates out of forEach → onAppReady abort.
  //
  // Fix: short-circuit when hint is missing, leaving the existing aria-
  // label / tooltip untouched. The editor still renders fine without
  // tooltips on those buttons.
  // ---------------------------------------------------------------------
  function patchButtonUpdateHint() {
    var Btn = window.Common && window.Common.UI && window.Common.UI.Button;
    if (!Btn || !Btn.prototype || Btn.prototype.__hintGuardInstalled) return false;
    var origUpdate = Btn.prototype.updateHint;
    Btn.prototype.updateHint = function (hint, isHtml) {
      if (hint == null) return;  // skip — leaves existing tooltip in place
      return origUpdate.call(this, hint, isHtml);
    };
    Btn.prototype.__hintGuardInstalled = true;
    log('Common.UI.Button.updateHint guarded');
    return true;
  }

  if (!patchButtonUpdateHint()) {
    var triesBH = 0;
    var ivBH = setInterval(function () {
      if (patchButtonUpdateHint() || ++triesBH > 300) clearInterval(ivBH);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 5d — defensive guard on Common.UI.MenuItem.setCaption.
  //
  // The HintManager's `updateShortcutHints` walks every menu item to
  // append " (Alt+X)" hints to its caption. For some menu items the
  // rendered anchor has no text node child — `cmpEl.find('> a').contents()
  // .last()[0]` returns undefined → `undefined.textContent = caption`
  // throws → updateShortcutHints aborts mid-iteration.
  //
  // The check `this.rendered` exists but it's not enough: a menu item
  // can be `rendered=true` while its inner anchor is built without a text
  // node (icon-only items, separators with custom markup, etc.).
  //
  // Fix: also check that the target node exists before assigning.
  // ---------------------------------------------------------------------
  function patchMenuItemSetCaption() {
    var MI = window.Common && window.Common.UI && window.Common.UI.MenuItem;
    if (!MI || !MI.prototype || MI.prototype.__captionGuardInstalled) return false;
    var origSetCaption = MI.prototype.setCaption;
    MI.prototype.setCaption = function (caption) {
      this.caption = caption;
      if (!this.rendered || !this.cmpEl) return;
      var node = this.cmpEl.find('> a').contents().last()[0];
      if (!node) return;  // anchor has no text node child — nothing to update
      node.textContent = caption;
    };
    MI.prototype.__captionGuardInstalled = true;
    log('Common.UI.MenuItem.setCaption guarded');
    return true;
  }

  if (!patchMenuItemSetCaption()) {
    var triesMC = 0;
    var ivMC = setInterval(function () {
      if (patchMenuItemSetCaption() || ++triesMC > 300) clearInterval(ivMC);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 7 (Phase 4.3) — save-byte extraction.
  //
  // The standard download path (`api.asc_DownloadAs(c_oAscFileType.X)`)
  // routes through `_downloadAsUsingServer` which posts a "save" command
  // to a real DocServer. We don't have one, and the subclass `_downloadAs`
  // overrides also assume a server upload to retrieve the converted file.
  //
  // BUT each editor exposes `asc_nativeGetFile()` (used by the desktop
  // edition for offline save), which serializes the in-memory document
  // model directly via `BinaryFileWriter.Write()` / `CBinaryFileWriter.
  // WriteDocument()` and returns a string of the form:
  //
  //     "<sig>;v<version>;<size>;<base64-encoded-payload>"
  //
  // where <sig> is "DOCY" / "XLSY" / "PPTY" — the same Editor.bin format
  // x2t consumes for `convertFromBin(bin, ext)` to round-trip back to
  // OOXML. No server, no async callback, no extra deps — just an in-band
  // serializer the editor already ships.
  //
  // We expose `window.__captureSave()` so the outer wrapper can pull the
  // bytes synchronously (the serialization itself is sync), then run them
  // through x2t to get docx/xlsx/pptx. The function returns a Uint8Array
  // of UTF-8 bytes (x2t-bridge's runConversion writes the input straight
  // to the WASM virtual FS, which is what x2t reads).
  // ---------------------------------------------------------------------
  window.__captureSave = function () {
    var ns = (window.DE || window.SSE || window.PE);
    if (!ns || !ns.controllers || !ns.controllers.Main) {
      throw new Error('captureSave: editor namespace not ready');
    }
    var api = ns.controllers.Main.api;
    if (!api) throw new Error('captureSave: api not ready');
    if (typeof api.asc_nativeGetFile !== 'function') {
      throw new Error('captureSave: asc_nativeGetFile missing on api');
    }
    // Returns a string with the signature header + base64 payload;
    // matches the on-disk format x2t produced for our load path.
    var binStr = api.asc_nativeGetFile();
    if (typeof binStr !== 'string') {
      throw new Error('captureSave: nativeGetFile returned non-string (type=' + typeof binStr + ')');
    }
    // The header + base64 body are pure ASCII, so per-char encode is safe
    // and avoids needing a TextEncoder polyfill check. Same approach as
    // the SDK's own native-bridge stub.
    var bytes = new Uint8Array(binStr.length);
    for (var i = 0; i < binStr.length; i++) bytes[i] = binStr.charCodeAt(i) & 0xff;
    // Record the history point these bytes represent so a CONFIRMED save
    // (save-ack) can mark exactly this point as the saved point — see
    // window.__skMarkSaved. We intentionally do NOT clear "modified" here:
    //   • the save isn't durable until the host acks, and
    //   • SetDocumentModified(false) only flips a boolean that the editor
    //     immediately recomputes from History.Index !== History.SavedIndex on
    //     the next interface update (e.g. a cursor move), so it flipped right
    //     back to dirty. Advancing SavedIndex on ack is the durable fix.
    try {
      window.__skSavedHistoryIndex =
        (window.AscCommon && AscCommon.History) ? AscCommon.History.Index : null;
    } catch (e) { window.__skSavedHistoryIndex = null; }
    log('captureSave: ' + bytes.length + ' bytes (' + binStr.slice(0, 16) + '…) @index=' + window.__skSavedHistoryIndex);
    return bytes;
  };

  // Mark the document as saved at the point captured by the last
  // __captureSave. Advances the SDK history "saved point"
  // (SavedIndex/UserSavedIndex) to that index, then recomputes the modified
  // flag. This does NOT touch the undo/redo stack — SavedIndex is only a marker
  // into the history, so Undo/Redo keep working exactly as before; the only
  // effect is that Have_Changes()/the modified flag now read clean until a real
  // content edit advances History.Index (cursor/selection moves don't).
  //
  // Call ONLY after the host confirms the save (save-ack ok) — never
  // optimistically, or a failed save would falsely look clean. Returns whether
  // the doc is STILL modified, i.e. the user edited PAST the saved point during
  // the upload (true ⇒ caller should show 'dirty', not 'saved').
  window.__skMarkSaved = function () {
    var ns  = (window.DE || window.SSE || window.PE);
    var api = ns && ns.controllers && ns.controllers.Main && ns.controllers.Main.api;
    var H   = window.AscCommon && AscCommon.History;
    if (!api || !H) return false;
    var idx = window.__skSavedHistoryIndex;
    try {
      if (typeof idx === 'number') {
        // Mirror what the editor's own save does (Reset_SavedIndex sets
        // SavedIndex = Index), but anchor to the CAPTURED index so edits made
        // during the upload stay correctly dirty. Set both markers since
        // Have_Changes compares against UserSavedIndex or SavedIndex.
        H.SavedIndex     = idx;
        H.UserSavedIndex = idx;
      } else if (typeof H.Reset_SavedIndex === 'function') {
        H.Reset_SavedIndex(true);
      }
      if (typeof H.ForceSave !== 'undefined') H.ForceSave = false;
      // Recompute "modified" from the advanced saved point + re-emit
      // asc_onDocumentModifiedChanged so the diskette/host reflect reality.
      if (typeof api.CheckChangedDocument === 'function') api.CheckChangedDocument();
    } catch (e) { log('markSaved threw:', e.message); }
    try {
      if (typeof H.Have_Changes === 'function') return !!H.Have_Changes();
      if (typeof api.isDocumentModified === 'function') return !!api.isDocumentModified();
    } catch (e) { /* fall through */ }
    return false;
  };

  function log() {
    if (window.console) {
      console.log.apply(console, ['[editor-stubs]'].concat([].slice.call(arguments)));
    }
  }

  // ---------------------------------------------------------------------
  // Layer 3 (Phase 4 cascade) — onServerVersion latch shim.
  //
  // Main.onServerVersion checks `window.compareVersions` AND compares the
  // build version reported by the SDK against the editor's about-menu
  // version. If they don't match it pops a warning AND latches
  // this.changeServerVersion = true, making the next call early-return.
  // We don't have a real DocServer reporting the build version, so we
  // bypass the version check entirely by setting the flag the function
  // looks for.
  // ---------------------------------------------------------------------
  window.compareVersions = true;

  // ---------------------------------------------------------------------
  // Layer 4 (Phase 4 cascade) — CDocsCoApi.auth no-op + synthesise success.
  //
  // baseEditorsApi.asc_LoadDocument() constructs a CDocsCoApi instance and
  // calls auth() which expects to talk to a DocServer WebSocket and call
  // onFirstLoadChangesEnd / onConnect / onAuthParticipantsChanged on the
  // editor api. We patch CDocsCoApi.prototype to short-circuit the auth
  // flow with a successful-no-changes synthetic response.
  //
  // The patch happens in two phases because CDocsCoApi may not be defined
  // yet at this script's load time. We try to patch immediately, and if
  // the class doesn't exist we set up a poller that catches it as soon
  // as the SDK bundle finishes evaluating.
  // ---------------------------------------------------------------------
  function patchCDocsCoApi() {
    // CDocsCoApi may live in different namespaces depending on the bundle.
    var CoApi =
      (window.AscCommon && window.AscCommon.CDocsCoApi) ||
      window.CDocsCoApi ||
      null;
    if (!CoApi || !CoApi.prototype) return false;
    if (CoApi.prototype.__stubPatched) return true;

    // The original `auth` posts an authentication message over the
    // CoAuth socket and waits for the server to respond. We replace it
    // with a synchronous no-op that synthesises:
    //   1. The "no remote changes" callback (so the editor stops waiting)
    //   2. A successful auth result so post-auth controllers proceed
    var origAuth = CoApi.prototype.auth;
    CoApi.prototype.auth = function (data) {
      log('CDocsCoApi.auth — no-op stub');
      var self = this;

      // Fire whatever callbacks the editor's load chain expects after auth.
      // These are typically set as properties on `this` by the SDK during
      // construction; we call them defensively (no-op if missing).
      function safeCall(cb) {
        if (typeof cb === 'function') {
          try { cb.call(self); }
          catch (e) { log('callback threw:', e.message); }
        }
      }

      // Most cascades fail when these are missing — call them as no-arg
      // synthesised events. Real DocServer would pass a participants list,
      // a changes payload, etc. — empty/null means "you're alone, no
      // pending edits".
      setTimeout(function () {
        safeCall(self.onFirstLoadChangesEnd);
        safeCall(self.onConnect);
        if (typeof self.onAuthParticipantsChanged === 'function') {
          try { self.onAuthParticipantsChanged.call(self, []); }
          catch (e) { log('onAuthParticipantsChanged threw:', e.message); }
        }
      }, 0);
    };

    // The full CoAuth surface includes several other socket-level methods
    // the editor may call. Stub the rest as no-ops so any incidental call
    // doesn't crash. Names sourced from the SDK source (search for
    // CDocsCoApi.prototype.* in the bundle).
    [
      'init', 'connect', 'disconnect', 'sendChanges', 'saveChanges',
      'unSaveLock', 'lock', 'unlock', 'requestEditRights', 'requestUsers',
      'sendUserAction', 'getUsers', 'sendCallback', 'sendUserEvent',
      'forceSave'
    ].forEach(function (name) {
      if (typeof CoApi.prototype[name] !== 'function') {
        CoApi.prototype[name] = function () { /* no-op stub */ };
      }
    });

    CoApi.prototype.__stubPatched = true;
    log('CDocsCoApi patched');
    return true;
  }

  if (!patchCDocsCoApi()) {
    var tries = 0;
    var iv = setInterval(function () {
      if (patchCDocsCoApi() || ++tries > 300) {
        clearInterval(iv);
      }
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 5 — Main.loadBinary path: trigger onEditorPermissions synthetically.
  //
  // The SDK's `asc_setLicenseInfo` (which fires `asc_onEditorPermissions`)
  // lives in the closed-source sdkjs-ooxml addon. In our offline setup that
  // event never fires, so Main's appOptions stays half-populated and
  // applyModeEditorElements (which configures ReviewChanges, Protection,
  // and other controllers' appConfig) never runs.
  //
  // Workaround: when Main.loadBinary is called, we (1) synthesise a
  // permissions params object, (2) drive Main.onEditorPermissions ourselves,
  // (3) THEN run the SDK's actual binary-open via OpenDocumentFromBin so
  // onDocumentContentReady fires AFTER controllers are configured.
  //
  // We must wait for Main controller to be available; like CDocsCoApi, it's
  // defined later in the boot sequence. Poll until then.
  // ---------------------------------------------------------------------
  function buildSyntheticPermissions(w) {
    if (!w.AscCommon || !w.AscCommon.asc_CAscEditorPermissions) return null;
    var perms = new w.AscCommon.asc_CAscEditorPermissions();
    perms.rights = w.Asc.c_oRights.Edit;
    perms.licenseType = w.Asc.c_oLicenseResult.Success;
    perms.canBranding = true;
    perms.customization = true;
    perms.buildVersion = '0.0.0';
    perms.buildNumber = 0;
    perms.liveViewerSupport = false;
    perms.canCoAuthoring = false;
    perms.canReaderMode = true;
    return perms;
  }

  function patchMainLoadBinary() {
    var ns = (window.DE || window.SSE || window.PE);
    if (!ns || !ns.controllers || !ns.controllers.Main) return false;
    var main = ns.controllers.Main;
    if (main.__loadBinaryPatched) return true;

    // Register a Common.Gateway listener for `opendocumentfrombinary`.
    // Main.onLaunch is supposed to do this (controller/Main.js:268), but in
    // our stubbed boot flow the registration's underlying jQuery binding
    // doesn't survive — likely because Backbone's Application.start runs
    // controller initialization in a way that overrides Main's listener
    // table. We add our own listener here as a durable fallback so
    // `editor.openDocument(bin)` from DocsAPI actually reaches our
    // intercept.
    if (window.Common && window.Common.Gateway && !main.__gwListenerAdded) {
      window.Common.Gateway.on('opendocumentfrombinary', function (data) {
        main.loadBinary(data);
      });
      // Also register the lower-priority command paths the editor uses
      window.Common.Gateway.on('opendocument', function (data) {
        if (typeof main.loadDocument === 'function') main.loadDocument(data);
      });
      main.__gwListenerAdded = true;
    }

    var origLoadBinary = main.loadBinary && main.loadBinary.bind(main);

    main.loadBinary = function (data) {
      log('loadBinary intercepted, bytes=' + (data && data.byteLength));
      var self = this;
      var bin = new Uint8Array(data);

      // 1. Make sure we're not blocked at the version-check latch.
      this.changeServerVersion = false;
      window.compareVersions = true;

      // 2. Stub onServerVersion so onEditorPermissions doesn't early-return.
      if (typeof this.onServerVersion === 'function' && !this.__serverVersionStubbed) {
        this.onServerVersion = function () { return false; };
        this.__serverVersionStubbed = true;
      }

      // 3. Provide minimal permissions on the document if missing.
      this.permissions = this.permissions || {
        edit: true, review: false, comment: false, fillForms: false,
        download: false, print: false, copy: true, modifyContentControl: false,
        modifyFilter: true, chat: false, protect: false
      };
      // Ensure user fields exist (getUserInitials reads .firstname/.lastname).
      this.editorConfig = this.editorConfig || {};
      this.editorConfig.user = this.editorConfig.user || {};
      var u = this.editorConfig.user;

      if (window.__skUser) {
        if (window.__skUser.id)
          u.id = window.__skUser.id;
        if (window.__skUser.name)
          u.name = u.fullname = window.__skUser.name;
      }

      if (!u.id) u.id = 'sk-editor-user';
      if (!u.fullname) u.fullname = u.name || 'Sharekey user';
      if (!u.name) u.name = u.fullname;
      if (!u.firstname) u.firstname = String(u.fullname).split(' ')[0];
      if (!u.lastname) u.lastname = String(u.fullname).split(' ').slice(1).join(' ');
      if (!u.image) u.image = '';
      this.editorConfig.lang = this.editorConfig.lang || 'en';
      this.editorConfig.mode = this.editorConfig.mode || 'edit';
      this.editorConfig.targetApp = this.editorConfig.targetApp || 'desktop';

      // Document descriptor — each editor's Main controller reads the
      // doc descriptor from a different field:
      //   word  → this.document         (set in DE.controllers.Main.loadDocument)
      //   slide → this.document         (set in PE.controllers.Main.loadDocument)
      //   cell  → this.appOptions.spreadsheet (set in SSE.controllers.Main.loadDocument)
      //
      // onEditorPermissions then reads `<descriptor>.info.favorite`,
      // `.fileType`, `.title`, etc. We populate the minimum keys those
      // call sites need so the synthetic permissions flow doesn't crash
      // before it reaches the appOptions.isEdit assignment (line 1493 in
      // cell's Main.js — anything that bails earlier leaves controllers
      // half-initialized).
      var fileTypeByEditor = {
        word: 'docx', slide: 'pptx', cell: 'xlsx'
      };
      var nsKey = (window.DE && 'word') || (window.PE && 'slide') || (window.SSE && 'cell') || 'word';
      var docDescriptor = {
        key:        'wrap-' + nsKey + '-' + Date.now(),
        title:      'Document.' + fileTypeByEditor[nsKey],
        url:        '',
        directUrl:  '',
        vkey:       '',
        fileType:   fileTypeByEditor[nsKey],
        token:      '',
        options:    {},
        permissions: this.permissions,
        info:       { favorite: false }
      };
      this.document = this.document || docDescriptor;
      this.appOptions = this.appOptions || {};
      // Cell-specific: appOptions.spreadsheet is the descriptor anchor
      // for SSE's onEditorPermissions (line 1471 reads .info.favorite).
      // Setting it here mirrors what SSE.controllers.Main.loadDocument
      // would normally populate from the postMessage `opendocument` payload.
      if (window.SSE && !this.appOptions.spreadsheet) {
        this.appOptions.spreadsheet = docDescriptor;
      }

      // Layer 7 — set DocInfo on the api before asc_LoadDocument runs.
      // asc_LoadDocument reads `this.DocInfo.get_Encrypted()`; without
      // DocInfo the call crashes. The v0 viewer also does this (see
      // master plan README "boot fixes that work").
      //
      // DocInfo also carries the user, and that is where every tracked change
      // gets its author: ReviewInfo.Update() copies DocInfo's user id and full
      // name into each new revision. asc_CDocInfo.get_UserName() returns null
      // when no UserInfo is attached, and the review panel then calls
      // Common.Utils.getUserInitials(null), which throws on .split(). So the
      // UserInfo below is not optional decoration — without it the editor
      // breaks the document on the first edit inside a tracked change.
      // @adr-0001
      try {
        if (this.api && window.Asc && window.Asc.asc_CDocInfo) {
          var info = this.api.DocInfo || new window.Asc.asc_CDocInfo();

          if (window.Asc.asc_CUserInfo && typeof info.put_UserInfo === 'function') {
            var userInfo = new window.Asc.asc_CUserInfo();
            userInfo.put_Id(u.id);
            userInfo.put_FullName(u.fullname);
            userInfo.put_FirstName(u.firstname);
            userInfo.put_LastName(u.lastname);
            info.put_UserInfo(userInfo);
            log('DocInfo user = ' + u.fullname + ' (' + u.id + ')');
          } else {
            log('WARNING: asc_CUserInfo or put_UserInfo missing — revisions will have no author');
          }

          if (typeof this.api.asc_setDocInfo === 'function') {
            this.api.asc_setDocInfo(info);
          } else {
            this.api.DocInfo = info;
          }
        }
      } catch (e) {
        log('asc_setDocInfo threw:', e.message);
      }

      // 4. Drive Main.onEditorPermissions synthetically. This populates
      //    appOptions, then calls applyModeCommonElements +
      //    applyModeEditorElements, which configure all the controllers
      //    (ReviewChanges, Protection, Comments, etc.) before the SDK
      //    fires onDocumentContentReady.
      try {
        var perms = buildSyntheticPermissions(window);
        if (perms && typeof this.onEditorPermissions === 'function') {
          this.onEditorPermissions(perms);
          log('synthetic onEditorPermissions complete, isEdit=', this.appOptions && this.appOptions.isEdit);
        }
      } catch (e) {
        log('onEditorPermissions threw:', e.message);
      }

      // 5. Hand the binary to the actual SDK loader. We use OpenDocumentFromBin
      //    directly (same path the v0 viewer uses) — it's more reliable than
      //    asc_openDocumentFromBytes which routes through onEndLoadFile and
      //    behaves differently in this offline configuration.
      try {
        if (this.api) {
          this.api.ServerIdWaitComplete = true;
          if (window.AscCommon && window.AscCommon.g_font_loader) {
            window.AscCommon.g_font_loader.fontFilesPath = '/fonts/';
          }
          // Install the picker-fallback shim now that g_font_infos is
          // populated by the SDK's checkAllFonts. Idempotent.
          if (primeFontPickerOnce()) {
            log('font picker shim installed');
          } else {
            log('warning: font picker shim could not install (manifest not ready)');
          }
          // PROTOTYPE: if real MS Core Fonts are enabled, repoint the Liberation
          // file slots to them BEFORE the document opens (so glyphs load real
          // Arial/Times/Courier). No-op unless SK_REAL_MSCOREFONTS is on.
          ensureRealMsFontsOnce();
          this.api.OpenDocumentFromBin(null, bin);
        } else if (origLoadBinary) {
          origLoadBinary(data);
        }
      } catch (e) {
        log('OpenDocumentFromBin threw:', e.message);
        if (origLoadBinary) origLoadBinary(data);
      }
    };
    main.__loadBinaryPatched = true;
    // Signal the OUTER wrapper that the open-listener + loadBinary patch are
    // installed, so wrapper-postmessage.onLoad doesn't dispatch openDocument
    // before anyone is listening. That race is the intermittent "stuck on
    // loading": x2t convert can finish and openDocument fire BEFORE this runs —
    // especially when a backgrounded tab throttles the editor bundle init while
    // x2t (synchronous WASM) stays fast — so the 'opendocumentfrombinary' event
    // lands with no listener and is dropped.
    window.__wrapperOpenReady = true;
    log('Main.loadBinary patched');
    return true;
  }

  if (!patchMainLoadBinary()) {
    var triesM = 0;
    var ivM = setInterval(function () {
      if (patchMainLoadBinary() || ++triesM > 300) clearInterval(ivM);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 8 — local image insert (DocServer-less upload bypass).
  //
  // asc_addImage() (api.js) opens the OS file picker, then routes the
  // picked File(s) through baseEditorsApi._uploadCallback →
  // AscCommon.UploadImageFiles (editorscommon.js), which POSTs each file to
  // the DocServer upload service: `sUploadServiceLocalUrl + '/' + documentId`.
  // With no DocServer (and documentId null) that's `POST /upload/null` →
  // 404 → asc_onError → onWarning(-13 "Image URL is incorrect"), and nothing
  // is inserted.
  //
  // We replace UploadImageFiles with a local implementation: read each File
  // as a base64 `data:` URL and hand those URLs back through the SAME
  // callback signature the SDK expects — callback(No, [url, ...]). Everything
  // downstream is unchanged: _uploadCallback → _addImageUrl → ImageLoader
  // loads the data: URL directly (no network) → AddImages / AddImageUrlAction
  // inserts it. Works for plain insert, change-image-url, and shape-texture
  // (the `obj` variants), since UploadImageFiles isn't aware of `obj`.
  //
  // Round-trip on save: the image's src stays the `data:` URL in the document
  // model — DocumentUrls.getImageLocal() returns null for `data:` (see
  // editorscommon.js), so it is NOT remapped to a media/ path. BinaryFileWriter
  // serialises the data: URL into the Editor.bin ImageMap, and x2t embeds the
  // decoded bytes into word/media of the output .docx. No upload, no
  // /working/media coordination — the bin is self-contained for inserted
  // images (existing doc images keep going through the media/ blob path).
  // ---------------------------------------------------------------------
  function readFileAsDataURL(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error || new Error('FileReader failed')); };
      fr.readAsDataURL(file);
    });
  }

  function patchUploadImageFiles() {
    var AscCommon = window.AscCommon;
    if (!AscCommon || typeof AscCommon.UploadImageFiles !== 'function') return false;
    if (AscCommon.__localUploadPatched) return true;

    AscCommon.UploadImageFiles = function (files, documentId, documentUserId, jwt, shardKey, wopiSrc, userSessionId, callback) {
      var ID = window.Asc && window.Asc.c_oAscError && window.Asc.c_oAscError.ID;
      var No = ID ? ID.No : 0;

      if (!files || files.length === 0) {
        callback(ID ? ID.UplImageFileCount : -1);
        return;
      }
      // Preserve the picked order — _addImageUrl inserts the urls in order.
      var list = [];
      for (var i = 0; i < files.length; i++) list.push(files[i]);

      Promise.all(list.map(readFileAsDataURL)).then(function (urls) {
        log('UploadImageFiles (local): ' + urls.length + ' image(s) → data: URL');
        callback(No, urls);
      }).catch(function (e) {
        log('UploadImageFiles (local) failed: ' + e.message);
        callback(ID ? ID.UplImageUrl : -1);
      });
    };
    AscCommon.__localUploadPatched = true;
    log('AscCommon.UploadImageFiles patched (local data: URL insert)');
    return true;
  }

  if (!patchUploadImageFiles()) {
    var triesUI = 0;
    var ivUI = setInterval(function () {
      if (patchUploadImageFiles() || ++triesUI > 300) clearInterval(ivUI);
    }, 100);
  }

  // ---------------------------------------------------------------------
  // Layer 9 (cell cascade) — fire asc_onInitEditorStyles so the spreadsheet's
  // delayed UI (right-click context menu, cell-style gallery) gets built.
  //
  // SSE builds its context menu lazily: a poller in SSE Main (the `timer_sl`
  // setInterval) only calls documentHolderView.createDelayedElements() +
  // toolbarController.createDelayedElements() once `window.styles_loaded`
  // flips true. That flag is set at the END of Toolbar.onApiInitEditorStyles,
  // which runs on the SDK's `asc_onInitEditorStyles` event.
  //
  // The SDK fires that event from spreadsheet_api._sendWorkbookStyles(), but
  // the only open-path caller is asc_ApplyColorScheme(), which skips it in
  // view mode ("На view-режиме не нужно отправлять стили", cell/api.js). In
  // our DocServer-less boot that path doesn't run, so the event never fires
  // (window.styles_loaded stays `undefined`) and **right-click shows no
  // menu** — createDelayedElements() was never reached.
  //
  // Fix: once the workbook view (api.wb / api.wbModel) exists, call
  // _sendWorkbookStyles() directly. That method has NO view-mode guard of its
  // own — it just triggers asc_onInitEditorStyles with the cell styles — so
  // it drives onApiInitEditorStyles → window.styles_loaded = true →
  // createDelayedElements(). We retry on an interval because
  // onApiInitEditorStyles early-returns (without setting the flag) if the
  // toolbar's style combo isn't built yet; we stop as soon as the flag flips.
  //
  // Cell-only: word (DE) and slide (PE) have no such gate, so we bail there.
  // ---------------------------------------------------------------------
  function kickEditorStyles() {
    var ns = window.SSE;
    if (!ns || !ns.controllers || !ns.controllers.Main) return false;
    var api = ns.controllers.Main.api;
    if (!api || typeof api._sendWorkbookStyles !== 'function') return false;
    if (!api.wb || !api.wbModel) return false;        // workbook view not drawn yet
    if (window.styles_loaded === true) return true;   // styles already initialised
    try {
      api._sendWorkbookStyles();
    } catch (e) {
      log('kickEditorStyles threw:', e.message);
    }
    return window.styles_loaded === true;
  }

  var triesSt = 0;
  var ivSt = setInterval(function () {
    // Not the cell editor (word/slide) → no styles_loaded gate to open.
    if (window.DE || window.PE) { clearInterval(ivSt); return; }
    if (window.styles_loaded === true) { clearInterval(ivSt); return; }
    if (kickEditorStyles() || ++triesSt > 600) {
      if (window.styles_loaded === true) log('asc_onInitEditorStyles kicked → styles_loaded=true');
      clearInterval(ivSt);
    }
  }, 100);

  // ---------------------------------------------------------------------
  // Layer 10 (cell cascade) — keep the cell editor's edit-only api callbacks
  // alive across the wrapper's view→edit mode transitions.
  //
  // SSE's DocumentHolder registers a block of edit-only api callbacks in its
  // setEvents() (run from createPostLoadElements on 'script:loaded'), gated on
  // permissions.isEdit AT THAT MOMENT. The most visible is:
  //   asc_onSetAFDialog — the column auto-filter settings dialog (clicking an
  //   applied filter). Also asc_onEditCell, asc_onEntriesListMenu /
  //   asc_onValidationListMenu (in-cell dropdowns), asc_onFormulaCompleteMenu,
  //   special-paste, autocorrect, etc.
  //
  // Two things break these in our setup:
  //  1. Boot-order race — setEvents can run before our synthetic
  //     onEditorPermissions flips isEdit true, so the whole block is skipped.
  //  2. View-mode application — the wrapper boots the sheet in view mode and
  //     applies edit/view via asc_setRestriction + an `editing:disable`
  //     broadcast. That drops the edit-only api callbacks (observed: after
  //     boot, asc_onSetAFDialog is gone while the non-gated asc_onHyperlinkClick
  //     survives), and entering edit mode never re-adds them.
  //
  // Fix: (re)assert the edit-only callbacks whenever they're MISSING and the
  // editor is edit-capable — via a bounded boot sweep AND every time edit mode
  // is (re)entered (`editing:disable` === false). The
  // asc_checkNeedCallback('asc_onSetAFDialog') guard makes it idempotent: we
  // only add when absent, so a single click never hits two onApiAutofilter
  // listeners (which would toggle the dialog open→closed). We also gate on
  // asc_onHyperlinkClick (DocumentHolder-exclusive, registered by setEvents) so
  // we never pre-empt setEvents and create a duplicate during boot.
  //
  // Cell-only. Mirrors DocumentHolderExt.setEvents' `if (isEdit)` block.
  // ---------------------------------------------------------------------
  function ensureCellEditCallbacks() {
    var ns = window.SSE;
    if (!ns || !ns.controllers || !ns.controllers.DocumentHolder || !ns.controllers.Main) return false;
    var dh = ns.controllers.DocumentHolder;
    var api = ns.controllers.Main.api;
    if (!api || !dh.permissions || typeof api.asc_checkNeedCallback !== 'function') return false;
    if (dh.permissions.isEdit !== true) return false;                     // edit capability off → nothing to do
    if (!api.asc_checkNeedCallback('asc_onHyperlinkClick')) return false; // setEvents hasn't run → don't pre-empt it
    if (api.asc_checkNeedCallback('asc_onSetAFDialog')) return true;      // already present → no duplicate
    function reg(name, fn) {
      if (typeof fn === 'function') api.asc_registerCallback(name, fn.bind(dh));
    }
    reg('asc_onSetAFDialog',             dh.onApiAutofilter);
    reg('asc_onEditCell',                dh.onApiEditCell);
    if (typeof dh.onEntriesListMenu === 'function') {
      api.asc_registerCallback('asc_onEntriesListMenu',    dh.onEntriesListMenu.bind(dh, false)); // Alt+Down
      api.asc_registerCallback('asc_onValidationListMenu', dh.onEntriesListMenu.bind(dh, true));
    }
    reg('asc_onFormulaCompleteMenu',     dh.onApiFormulaCompleteMenu);
    reg('asc_onShowSpecialPasteOptions', dh.onShowSpecialPasteOptions);
    reg('asc_onHideSpecialPasteOptions', dh.onHideSpecialPasteOptions);
    reg('asc_onToggleAutoCorrectOptions',dh.onToggleAutoCorrectOptions);
    reg('asc_onFormulaInfo',             dh.onFormulaInfo);
    reg('asc_ChangeCropState',           dh.onChangeCropState);
    reg('asc_onInputMessage',            dh.onInputMessage);
    reg('asc_onTableTotalMenu',          dh.onTableTotalMenu);
    reg('asc_onShowPivotHeaderDetailsDialog', dh.onShowPivotHeaderDetailsDialog);
    reg('asc_onShowPivotGroupDialog',    dh.onShowPivotGroupDialog);
    reg('asc_doubleClickOnTableOleObject', dh.onDoubleClickOnTableOleObject);
    reg('asc_onSingleChartSelectionChanged', dh.onSingleChartSelectionChanged);
    // Empty-placeholder image insert (uses the local-upload path from Layer 8).
    if (typeof api.asc_registerPlaceholderCallback === 'function' &&
        window.AscCommon && window.AscCommon.PlaceholderButtonType) {
      if (typeof dh.onInsertImage === 'function')
        api.asc_registerPlaceholderCallback(window.AscCommon.PlaceholderButtonType.Image, dh.onInsertImage.bind(dh));
      if (typeof dh.onInsertImageUrl === 'function')
        api.asc_registerPlaceholderCallback(window.AscCommon.PlaceholderButtonType.ImageUrl, dh.onInsertImageUrl.bind(dh));
    }
    log('cell edit callbacks ensured (asc_onSetAFDialog + edit-cell/list/validation/…)');
    return true;
  }

  // ---------------------------------------------------------------------
  // Layer 10b (cell cascade) — build the EDIT-mode right-click context menus
  // when the editor is edit-capable, regardless of the view-first boot.
  //
  // SSE creates its context menus lazily and PICKS edit-vs-viewer by isEdit at
  // boot, in two places, both racing the same way as Layer 10:
  //   - Main.js: `if (appOptions.isEdit) { … createDelayedElements() } else {
  //     createDelayedElementsViewer() }`
  //   - DocumentHolder.createPostLoadElements: `permissions.isEdit ?
  //     createDelayedElements() : createDelayedElementsViewer()`
  // In our view-first boot isEdit is false at that moment, so only the VIEWER
  // menu (viewModeMenu) is built — the edit menus (ssMenu, copyPasteMenu, imgMenu,
  // funcMenu) never are. Entering edit via asc_setRestriction re-runs neither
  // path, so a right-click in edit mode hits `showObjectMenu → fillMenuProps →
  // showPopupMenu(documentHolder.ssMenu, …)` with ssMenu === undefined, and
  // showPopupMenu silently no-ops. (The copy menu still works in VIEW because
  // viewModeMenu exists — matching the symptom: view menu shows, edit menu doesn't.)
  //
  // Fix: call documentHolderView.createDelayedElements() once edit-capable AND
  // styles are loaded (Layer 9 forces styles_loaded; createDelayedElements early-
  // returns without it). Idempotent two ways — createDelayedElements guards on its
  // own `pmiCut`, and we guard on it too — and self-healing on edit-mode entry.
  //
  // Cell-only. Mirrors the edit-branch createDelayedElements in Main / DocumentHolder.
  // ---------------------------------------------------------------------
  function ensureCellDelayedMenus() {
    var ns = window.SSE;
    if (!ns || !ns.controllers || !ns.controllers.DocumentHolder) return false;
    var dh = ns.controllers.DocumentHolder;
    if (!dh.permissions || dh.permissions.isEdit !== true) return false;  // edit capability off → nothing to do
    var view = dh.documentHolder;
    if (!view || typeof view.createDelayedElements !== 'function') return false;
    if (view.pmiCut) return true;             // edit menus already built (createDelayedElements' own guard var)
    if (!window.styles_loaded) return false;  // createDelayedElements would no-op without it → wait for Layer 9
    try {
      view.createDelayedElements();
    } catch (e) {
      log('ensureCellDelayedMenus threw:', e.message);
      return false;
    }
    if (view.pmiCut) log('cell edit context menus built (ssMenu/copyPasteMenu/imgMenu/funcMenu)');
    return !!view.pmiCut;
  }

  // Re-assert whenever edit mode is (re)entered — the view-mode broadcast drops
  // these, and edit-entry is exactly when they need to come back.
  function hookCellEditCbRefresh() {
    if (window.__cellEditCbHooked) return true;
    var C = window.Common;
    if (!C || !C.NotificationCenter || typeof C.NotificationCenter.on !== 'function') return false;
    C.NotificationCenter.on('editing:disable', function (disable) {
      if (!disable) {
        setTimeout(ensureCellEditCallbacks, 0);   // edit mode entered, after the editor's own handlers run
        setTimeout(ensureCellDelayedMenus, 0);    // …and make sure the edit-mode context menus exist
      }
    });
    window.__cellEditCbHooked = true;
    log('cell edit-callback refresh hook installed (editing:disable)');
    return true;
  }

  var triesCB = 0;
  var ivCB = setInterval(function () {
    if (window.DE || window.PE) { clearInterval(ivCB); return; }   // word/slide: N/A
    hookCellEditCbRefresh();
    ensureCellEditCallbacks();
    ensureCellDelayedMenus();
    if (++triesCB > 600) clearInterval(ivCB);                      // boot sweep ends; the editing:disable hook persists
  }, 100);

  // ---------------------------------------------------------------------
  // Layer 11 (all editors) — font picker without the DocServer thumbnail sprite.
  //
  // The font dropdown renders each font name as a pre-rasterised tile from a
  // sprite the DocumentServer normally generates and serves at
  // `<thumbnailsPath>/fonts_thumbnail*.png(.bin)`. We have no DocServer, so that
  // file 404s — and the sprite loader (CThumbnailLoader in web-apps
  // ComboBoxFonts.js) does NOT check the XHR status ("// TODO: check errors"):
  // it parses the 404 response body as a binary header, reads a garbage
  // width/heightOne, and calls `ctx.createImageData(width, heightOne)` with
  // absurd dimensions →
  //   RangeError: Failed to execute 'createImageData' … Out of memory
  // which the SDK reports as the generic critical "An error occurred during the
  // work with the document" modal the moment the user opens the font list.
  //
  // CThumbnailLoader is closure-private, but `Common.UI.ComboBoxFonts` is a
  // global prototype, so we override two of its methods:
  //   1. loadSprite()  — skip the 404 XHR; install a harmless stub whose
  //      getImage() returns a 1×1 canvas (so any stray caller can't OOM).
  //   2. updateVisibleFontsTiles() — the list <a.font-item> is EMPTY in the
  //      template (the name only ever came from the tile image), so render the
  //      font NAME as plain text instead. The picker stays fully usable; we just
  //      lose the per-font visual preview (impossible without the sprite).
  //
  // Install via poller — Common.UI.ComboBoxFonts is defined late (web-apps app.js).
  // ---------------------------------------------------------------------
  function patchFontComboThumbnails() {
    var CB = window.Common && window.Common.UI && window.Common.UI.ComboBoxFonts;
    if (!CB || !CB.prototype || CB.prototype.__skFontThumbPatched) return false;

    CB.prototype.loadSprite = function (callback) {
      // No DocServer sprite — stub it out so nothing createImageData()s a 404.
      this.spriteThumbs = {
        width: 0, heightOne: 0, height: 0, count: 0,
        load: function (u, cb) { if (cb) cb(); },
        getImage: function () { var c = document.createElement('canvas'); c.width = 1; c.height = 1; return c; }
      };
      if (callback) callback();
    };

    CB.prototype.updateVisibleFontsTiles = function () {
      var me = this;
      // The un-overridden flushVisibleFontsTiles() reads this.tiles.length on
      // menu show/hide — keep it a (always-empty, since we render text) array so
      // it never sees undefined.
      if (!me.tiles) me.tiles = [];
      if (!me.el || !me.store || typeof me.store.at !== 'function' || !window.$) return;
      var $el = window.$(me.el);
      for (var i = 0; i < me.store.length; i++) {
        var rec = me.store.at(i);
        if (!rec) continue;
        var id = rec.get('id'), name = rec.get('name');
        if (!id || !name) continue;
        var a = $el.find('#' + id + ' > a.font-item')[0];
        if (!a || a.getAttribute('data-sk-fontname')) continue;   // already filled
        a.textContent = name;
        a.style.lineHeight = a.style.height || '';                // vertical-center in the fixed-height row
        a.setAttribute('data-sk-fontname', '1');
      }
    };

    // We never create tile nodes (names are plain text), so there's nothing to
    // remove — just keep this.tiles a valid array. The original reads
    // this.tiles.length unconditionally and can run before updateVisibleFontsTiles
    // (e.g. onAfterShowMenu), which threw "Cannot read properties of undefined".
    CB.prototype.flushVisibleFontsTiles = function () {
      this.tiles = [];
    };

    CB.prototype.__skFontThumbPatched = true;
    log('Common.UI.ComboBoxFonts: thumbnail sprite disabled → font names as text (no DocServer sprite)');
    return true;
  }

  if (!patchFontComboThumbnails()) {
    var triesFC = 0;
    var ivFC = setInterval(function () {
      if (patchFontComboThumbnails() || ++triesFC > 600) clearInterval(ivFC);
    }, 100);
  }

  log('loaded');
})();
