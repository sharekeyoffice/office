#!/bin/bash
# 03-deploy-web-apps.sh — grunt-build the web-apps editor UI, then inject our
# boot scripts. Uses the wrapper's inject-boot.sh (NOT the in-tree viewerPoc).
#
# Inputs:  vendor/web-apps/, wrapper/v1/web-apps-overlay/
# Outputs: vendor/web-apps/deploy/{web-apps,sdkjs}/

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
WA="$REPO/vendor/web-apps"
WRAPPER="$REPO/wrapper"

[[ -d "$WA/build" ]] || { echo "ERROR: $WA/build missing. Run 01 first." >&2; exit 1; }

if [[ ! -d "$WA/build/node_modules" ]]; then
  echo "→ installing web-apps build deps (this is the slow step)"
  ( cd "$WA/build" && npm install --silent --no-audit --no-fund )
fi
command -v npx >/dev/null 2>&1 || { echo "ERROR: npx not found" >&2; exit 1; }

cd "$WA/build"
# Build the shared common/api assets FIRST (the editor -component targets do
# NOT build them — Gruntfile deploy-<editor>-component = init + deploy-app only).
echo "→ grunt deploy-common-component"
npx grunt deploy-common-component
for ed in documenteditor spreadsheeteditor presentationeditor; do
  echo "→ grunt deploy-${ed}-component"
  npx grunt "deploy-${ed}-component"
done

for ed in documenteditor spreadsheeteditor presentationeditor; do
  idx="$WA/deploy/web-apps/apps/$ed/main/index.html"
  [[ -f "$idx" ]] || { echo "ERROR: missing $idx after grunt deploy" >&2; exit 1; }
done
[[ -f "$WA/deploy/web-apps/apps/api/documents/api.js" ]] || \
  { echo "ERROR: apps/api/documents/api.js missing — deploy-common-component didn't run?" >&2; exit 1; }

INJECT="$WRAPPER/v1/web-apps-overlay/inject-boot.sh"
[[ -x "$INJECT" ]] || chmod +x "$INJECT"
echo "→ injecting boot scripts into deployed index.htmls"
bash "$INJECT" --web-apps "$WA" --editors documenteditor,spreadsheeteditor,presentationeditor

for ed in documenteditor spreadsheeteditor presentationeditor; do
  grep -q "wrapper-boot-injected" "$WA/deploy/web-apps/apps/$ed/main/index.html" || \
    { echo "ERROR: boot not injected into $ed" >&2; exit 1; }
done
echo "✓ web-apps deployed + boot injected ($(du -sh "$WA/deploy" | cut -f1))"
