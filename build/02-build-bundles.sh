#!/bin/bash
# 02-build-bundles.sh — produce editor-mode SDK bundles from vendor/sdkjs
# using the wrapper's tooling (NOT the in-tree viewerPoc).
#
# Inputs:  vendor/sdkjs/ (prepared by 01), wrapper/ (manifests + bundle.js)
# Outputs: wrapper/v1/bundle/{word,cell,slide}.editor.bundle.js{,.gz,.br}

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
WRAPPER="$REPO/wrapper"
SDK="$REPO/vendor/sdkjs"

[[ -d "$SDK/configs" ]] || { echo "ERROR: $SDK missing/incomplete. Run 01 first." >&2; exit 1; }

# esbuild (pinned in the root package.json devDependencies for byte-determinism).
# bundle.js resolves it from $REPO/node_modules via Node's upward module search.
if [[ ! -d "$REPO/node_modules/esbuild" ]]; then
  echo "→ installing build deps (esbuild, pinned)"
  ( cd "$REPO" && npm install --silent --no-audit --no-fund )
fi

# Verify the manifests still derive from this upstream (informational).
echo "→ verifying manifests derive from vendor/sdkjs"
node "$WRAPPER/build/gen-manifest.js" --sdk-dir "$SDK" || \
  echo "  (manifests differ from a pure configs derivation — using committed manifests)"

for editor in word cell slide; do
  echo "→ building $editor editor bundle"
  ( cd "$WRAPPER" && node build/bundle.js \
      --mode editor --editor "$editor" --sdk-dir "$SDK" \
      --manifest-dir "$WRAPPER" --out-dir "$WRAPPER/v1/bundle" )
done

for editor in word cell slide; do
  out="$WRAPPER/v1/bundle/${editor}.editor.bundle.js"
  [[ -f "$out" ]] || { echo "ERROR: bundle missing: $out" >&2; exit 1; }
  printf "  %s: %s bytes\n" "$editor" "$(stat -f%z "$out" 2>/dev/null || stat -c%s "$out")"
done
echo "✓ editor bundles built"
