#!/bin/bash
# 01-pull-upstream.sh — clone/checkout pinned upstream into vendor/, then
# prepare the sdkjs tree for bundling (apply our patch + stage AllFonts).
#
# Source of truth for the wrapper itself is sharekey-office/wrapper/ (NOT a
# fork, NOT a local viewerPoc). This step only fetches UPSTREAM sources.
#
# Reads upstream-pins/{sdkjs,web-apps}.json.
# Outputs: vendor/sdkjs/ (patched + AllFonts staged), vendor/web-apps/

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
VENDOR="$REPO/vendor"
WRAPPER="$REPO/wrapper"
mkdir -p "$VENDOR"

get_field() {
  local rel="$1" field="$2"
  ( cd "$REPO" && node -e "const fs=require('fs'); const j=JSON.parse(fs.readFileSync(process.argv[1],'utf8')); console.log(j[process.argv[2]]||'')" "$rel" "$field" )
}

clone_pinned() {
  local name="$1" pin="upstream-pins/$1.json" dir="$VENDOR/$1"
  local remote ref; remote=$(get_field "$pin" remote); ref=$(get_field "$pin" ref)
  echo "→ $name: $remote @ $ref"
  if [[ ! -d "$dir/.git" ]]; then
    rm -rf "$dir"; mkdir -p "$dir"
    ( cd "$dir" && git init -q && git remote add origin "$remote" )
  fi
  # reset --hard (not checkout) so re-runs discard any previously-applied
  # patch / staged AllFonts and start from a clean pinned tree (idempotent).
  ( cd "$dir" && git fetch --depth 1 origin "$ref" --quiet && git reset --hard --quiet FETCH_HEAD )
  echo "  $name @ $(cd "$dir" && git rev-parse --short HEAD)"
}

clone_pinned sdkjs
clone_pinned web-apps

# ── Prepare sdkjs for bundling ──────────────────────────────────
echo "→ applying sdkjs patches"
for p in "$WRAPPER"/sdkjs-patches/*.patch; do
  [[ -e "$p" ]] || continue
  if git -C "$VENDOR/sdkjs" apply --reverse --check "$p" 2>/dev/null; then
    echo "  already applied $(basename "$p")"
  else
    git -C "$VENDOR/sdkjs" apply "$p" && echo "  applied $(basename "$p")"
  fi
done

# REQUIRED: a clean upstream clone has NO common/AllFonts.js (OnlyOffice
# generates it at build time). The bundle manifest references it, so stage our
# hand-written one into the sdk tree before bundling. (wrapper/REGENERATE.md §A/§D.)
echo "→ staging wrapper AllFonts.js into vendor/sdkjs/common/"
cp "$WRAPPER/v1/sdk-runtime/common/AllFonts.js" "$VENDOR/sdkjs/common/AllFonts.js"

# x2t is committed in the wrapper (offline-buildable). Sanity-check it.
[[ -f "$WRAPPER/v1/x2t/x2t.wasm" ]] || { echo "ERROR: wrapper/v1/x2t/x2t.wasm missing" >&2; exit 1; }

echo "✓ upstream pulled & sdkjs prepared"
