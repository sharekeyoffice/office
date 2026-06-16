#!/bin/bash
# 05-prune.sh — slim the assembled site's web-apps/ by removing unneeded fat.
# Operates on $OUT_DIR (default dist/). ~535 MB → ~50 MB.

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST="${OUT_DIR:-$REPO/dist}"
[[ "$DIST" = /* ]] || DIST="$REPO/$DIST"
WA="$DIST/web-apps/apps"
KEEP_LOCALES="${KEEP_LOCALES:-en}"
[[ -d "$WA" ]] || { echo "ERROR: $WA missing. Run 04 first." >&2; exit 1; }

before=$(du -sh "$DIST" | cut -f1)
echo "  - removing in-editor help docs"
rm -rf "$WA"/{documenteditor,spreadsheeteditor,presentationeditor}/main/resources/help 2>/dev/null || true
echo "  - removing mobile + embed variants"
rm -rf "$WA"/{documenteditor,spreadsheeteditor,presentationeditor}/{mobile,embed} 2>/dev/null || true
echo "  - removing legacy IE polyfills"
rm -rf "$WA"/{documenteditor,spreadsheeteditor,presentationeditor}/main/ie 2>/dev/null || true
echo "  - removing source maps"
find "$WA" -name "*.js.map" -delete 2>/dev/null || true
find "$WA" -name "*.css.map" -delete 2>/dev/null || true
echo "  - trimming locales to: $KEEP_LOCALES"
IFS=',' read -r -a KEEP_ARR <<< "$KEEP_LOCALES"
in_keep(){ local l="$1"; for k in "${KEEP_ARR[@]}"; do [[ "$l" == "$k" ]] && return 0; done; return 1; }
for ed in documenteditor spreadsheeteditor presentationeditor; do
  loc="$WA/$ed/main/locale"
  [[ -d "$loc" ]] || continue
  # web-apps ships locales EITHER as per-language dirs (locale/<lang>/) OR as
  # per-language files (locale/<lang>.json). Handle both.
  for d in "$loc"/*/;     do [[ -d "$d" ]] || continue; in_keep "$(basename "$d")"        || rm -rf "$d"; done
  for f in "$loc"/*.json; do [[ -f "$f" ]] || continue; in_keep "$(basename "$f" .json)"  || rm -f  "$f"; done
done
echo "  - removing pdfeditor / visioeditor / forms (top-level + per-editor)"
rm -rf "$WA"/pdfeditor "$WA"/visioeditor "$WA"/forms 2>/dev/null || true
rm -rf "$WA"/{documenteditor,spreadsheeteditor,presentationeditor}/forms 2>/dev/null || true
echo "  - removing common's mobile assets"
rm -rf "$WA"/common/mobile 2>/dev/null || true

echo "✓ pruned  (before: $before, after: $(du -sh "$DIST" | cut -f1))"
