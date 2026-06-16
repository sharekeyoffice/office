#!/bin/bash
# verify-stubs.sh — grep upstream for every symbol the wrapper patches.
# Reports ✅ if the symbol still exists at the documented location,
# ⚠️ if it moved, 🔴 if it can't be found at all.
#
# Doesn't run the editor — just static checks against the source tree.
# For runtime verification, open the editor and use the console
# one-liner from the cascade verification notes.
#
# Usage:
#   ./verify-stubs.sh
#
# Run AFTER build/01-pull-upstream.sh has cloned the pinned upstream into
# vendor/ — this checks the exact sources the build will bundle.
#
# Env vars (override the defaults below):
#   SDKJS      path to sdkjs/    (default: <repo>/vendor/sdkjs)
#   WEB_APPS   path to web-apps/ (default: <repo>/vendor/web-apps)

set -euo pipefail

REPO="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SDKJS="${SDKJS:-$REPO/vendor/sdkjs}"
WEB_APPS="${WEB_APPS:-$REPO/vendor/web-apps}"

red()    { printf '\033[31m%s\033[0m' "$1"; }
green()  { printf '\033[32m%s\033[0m' "$1"; }
yellow() { printf '\033[33m%s\033[0m' "$1"; }

check() {
  local layer="$1"
  local desc="$2"
  local pattern="$3"
  local path="$4"
  local hits
  if hits=$(grep -rn "$pattern" "$path" 2>/dev/null | head -3); then
    if [[ -n "$hits" ]]; then
      printf "%-6s %s — %s\n" "$layer" "$(green '✅')" "$desc"
      printf "%-6s        %s\n" "" "$(echo "$hits" | head -1)"
      return 0
    fi
  fi
  printf "%-6s %s — %s\n" "$layer" "$(red '🔴')" "$desc"
  printf "%-6s        searched: %s/%s\n" "" "$path" "$pattern"
  return 1
}

echo "Cascade guard verification"
echo "  sdkjs:    $SDKJS"
echo "  web-apps: $WEB_APPS"
echo

[[ -d "$SDKJS" ]]    || { echo "ERROR: sdkjs not found at $SDKJS" >&2; exit 1; }
[[ -d "$WEB_APPS" ]] || { echo "ERROR: web-apps not found at $WEB_APPS" >&2; exit 1; }

# Layer 3
check "3"   "compareVersions consulted in onServerVersion" \
      "compareVersions" \
      "$WEB_APPS/apps/spreadsheeteditor/main/app/controller/Main.js"

# Layer 4
check "4"   "CDocsCoApi.prototype.auth exists" \
      "CDocsCoApi.prototype.auth\\b" \
      "$SDKJS/common/"

# Layer 5
for ed in documenteditor spreadsheeteditor presentationeditor; do
  check "5"   "$ed Main.loadBinary exists" \
        "loadBinary: function" \
        "$WEB_APPS/apps/$ed/main/app/controller/Main.js"
done

# Layer 5b
check "5b"  "Common.Locale.getCurrentLanguage exists" \
      "Common.Locale.getCurrentLanguage" \
      "$WEB_APPS/apps/common/"

# Layer 5c
check "5c"  "Button.updateHint reads hint[0]" \
      "hint\\[0\\]" \
      "$WEB_APPS/apps/common/main/lib/component/Button.js"

# Layer 5d
check "5d"  "MenuItem.setCaption reads .last()[0].textContent" \
      ".last()\\[0\\].textContent" \
      "$WEB_APPS/apps/common/main/lib/component/MenuItem.js"

# Layer 5e
check "5e"  "appOptions.spreadsheet read in cell onEditorPermissions" \
      "appOptions.spreadsheet.info" \
      "$WEB_APPS/apps/spreadsheeteditor/main/app/controller/Main.js"

# Layer 6 (font picker)
check "6"   "g_fontApplication.GetFontFileWeb exists" \
      "GetFontFileWeb" \
      "$SDKJS/common/libfont/map.js"

# Layer 6b
check "6b"  "CTextShaper.prototype.FlushWord reads m_pFaceInfo" \
      "FontId.m_pFaceInfo.family_name" \
      "$SDKJS/common/libfont/textshaper.js"

# Layer 6c
check "6c"  "MeasureCode reads Temp.fAdvanceX" \
      "Temp.fAdvanceX" \
      "$SDKJS/common/libfont/textmeasurer.js"

# Layer 6d
check "6d"  "ParaRun.prototype.AddText calls getUnicodeIterator" \
      "sString.getUnicodeIterator" \
      "$SDKJS/word/Editor/Run.js"

# Layer 7
for f in word/api.js cell/api.js slide/api.js; do
  check "7"   "asc_nativeGetFile in $f" \
        "asc_nativeGetFile\\b" \
        "$SDKJS/$f"
done

# Layer 8
check "8"   "UploadImageFiles (DocServer upload we bypass)" \
      "function UploadImageFiles" \
      "$SDKJS/common/editorscommon.js"
check "8"   "DocumentUrls.getImageLocal keeps data: inline" \
      "indexOf('data:image')" \
      "$SDKJS/common/editorscommon.js"

# Layer 9 (cell)
check "9"   "styles_loaded gates createDelayedElements (cell Main)" \
      "styles_loaded" \
      "$WEB_APPS/apps/spreadsheeteditor/main/app/controller/Main.js"
check "9"   "_sendWorkbookStyles fires asc_onInitEditorStyles (cell api)" \
      "_sendWorkbookStyles" \
      "$SDKJS/cell/api.js"

# Layer 10 (cell)
check "10"  "DocumentHolder edit block gated on isEdit (asc_onSetAFDialog)" \
      "asc_onSetAFDialog" \
      "$WEB_APPS/apps/spreadsheeteditor/main/app/controller/DocumentHolderExt.js"
check "10"  "asc_checkNeedCallback idempotency probe (cell api)" \
      "asc_checkNeedCallback" \
      "$SDKJS/cell/api.js"

# Layer 11
check "11"  "ComboBoxFonts.updateVisibleFontsTiles (font sprite tiles)" \
      "updateVisibleFontsTiles" \
      "$WEB_APPS/apps/common/main/lib/component/ComboBoxFonts.js"
check "11"  "ComboBoxFonts createImageData (OOM site we bypass)" \
      "createImageData" \
      "$WEB_APPS/apps/common/main/lib/component/ComboBoxFonts.js"

echo
echo "If any layer is 🔴, that symbol moved or got removed upstream."
echo "Read the corresponding cascade layer notes for next steps."
