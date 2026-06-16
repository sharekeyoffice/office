#!/bin/bash
# 04-assemble-dist.sh — assemble the static site at the URL paths the editor
# expects. Sources the wrapper from sharekey-office/wrapper/ (NOT viewerPoc).
#
# Output dir: $OUT_DIR (default: dist/). For a public/-comparison run, call with
#   OUT_DIR=public bash build/04-assemble-dist.sh
#
# Inputs:
#   wrapper/v1/            (glue, x2t, fonts, sdk-runtime)
#   wrapper/v1/bundle/     (editor bundles, from step 02)
#   vendor/web-apps/deploy/{web-apps,sdkjs}/ (from step 03)

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
V1="$REPO/wrapper/v1"
WA="$REPO/vendor/web-apps/deploy"
DIST="${OUT_DIR:-$REPO/dist}"
[[ "$DIST" = /* ]] || DIST="$REPO/$DIST"

for required in "$V1" "$WA"; do
  [[ -d "$required" ]] || { echo "ERROR: $required missing. Run earlier steps." >&2; exit 1; }
done

rm -rf "$DIST"; mkdir -p "$DIST"

echo "→ wrapper sources at root"
cp "$V1"/{edit.html,edit-host-demo.html,editor-stubs.js,wrapper-boot.js,wrapper-customization.js,wrapper-mount.js,wrapper-postmessage.js,wrapper-heartbeat.js,x2t-bridge.js} "$DIST/"
echo "→ service-worker stub"
echo "// stub — wrapper does not use a service worker" > "$DIST/document_editor_service_worker.js"
echo "→ x2t/ + fonts/"
cp -R "$V1/x2t" "$DIST/x2t"; cp -R "$V1/fonts" "$DIST/fonts"
echo "→ web-apps editor UI + sdkjs assets"
cp -R "$WA/web-apps" "$DIST/web-apps"; cp -R "$WA/sdkjs" "$DIST/sdkjs"
echo "→ overwrite AllFonts.js + libfont/engine with trimmed versions"
mkdir -p "$DIST/sdkjs/common/libfont/engine"
cp "$V1/sdk-runtime/common/AllFonts.js"               "$DIST/sdkjs/common/AllFonts.js"
cp "$V1/sdk-runtime/common/libfont/engine/fonts.js"   "$DIST/sdkjs/common/libfont/engine/fonts.js"
cp "$V1/sdk-runtime/common/libfont/engine/fonts.wasm" "$DIST/sdkjs/common/libfont/engine/fonts.wasm"
echo "→ overwrite sdk-all-min.js with our editor bundles"
for ed in word cell slide; do
  src="$V1/bundle/$ed.editor.bundle.js"; dst="$DIST/sdkjs/$ed"; mkdir -p "$dst"
  [[ -f "$src" ]] || { echo "ERROR: $src missing — run 02" >&2; exit 1; }
  cp "$src" "$dst/sdk-all-min.js"
  [[ -f "$src.gz" ]] && cp "$src.gz" "$dst/sdk-all-min.js.gz"
  [[ -f "$src.br" ]] && cp "$src.br" "$dst/sdk-all-min.js.br"
done
echo "→ cosmetic stubs"
mkdir -p "$DIST/sdkjs/common/Charts" "$DIST/sdkjs/common/Images/cursors"
echo "// stub" > "$DIST/sdkjs/common/Charts/ChartStyles.js"
echo "{}"     > "$DIST/sdkjs/common/Images/cursors/svg.json"

echo "→ applying overlay/* over $DIST"
cp -R "$REPO/overlay"/. "$DIST/"

# ── Substitute __ALLOWED_HOST_ORIGIN__ ─────────────────────────
ALLOWED_HOST_ORIGIN="${ALLOWED_HOST_ORIGIN:-}"
if [[ -z "$ALLOWED_HOST_ORIGIN" && -f "$REPO/config.json" ]]; then
  ALLOWED_HOST_ORIGIN=$(node -e "console.log((JSON.parse(require('fs').readFileSync('$REPO/config.json','utf8'))['allowedHostOrigin'])||'')")
fi
[[ -n "$ALLOWED_HOST_ORIGIN" ]] || { echo "ERROR: set ALLOWED_HOST_ORIGIN env or config.json" >&2; exit 1; }
# Each comma-separated entry must be an exact origin (https://host[:port]) OR a
# single-level wildcard subdomain (https://*.suffix). The wildcard is matched at
# runtime by window.matchHostOrigin in edit.html. Use a wildcard ONLY for
# dev/preview (e.g. https://*.dev.example.com); keep prod exact.
_IFS_SAVE="$IFS"; IFS=','
for _origin in $ALLOWED_HOST_ORIGIN; do
  _origin="${_origin#"${_origin%%[![:space:]]*}"}"; _origin="${_origin%"${_origin##*[![:space:]]}"}"  # trim
  [[ "$_origin" =~ ^https?://(\*\.)?[a-zA-Z0-9.-]+(:[0-9]+)?$ ]] || \
    { echo "ERROR: invalid ALLOWED_HOST_ORIGIN entry: '$_origin'" >&2; exit 1; }
done
IFS="$_IFS_SAVE"
if sed --version >/dev/null 2>&1; then SED=(sed -i); else SED=(sed -i ''); fi
"${SED[@]}" "s|__ALLOWED_HOST_ORIGIN__|$ALLOWED_HOST_ORIGIN|g" "$DIST/edit.html"
grep -q "__ALLOWED_HOST_ORIGIN__" "$DIST/edit.html" && { echo "ERROR: origin substitution failed" >&2; exit 1; }

echo "✓ assembled $DIST (origin: $ALLOWED_HOST_ORIGIN, pre-prune $(du -sh "$DIST" | cut -f1))"
