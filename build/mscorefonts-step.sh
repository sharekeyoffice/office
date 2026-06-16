#!/usr/bin/env bash
# mscorefonts-step.sh — build step shared by build.sh (dist/) and initialize.sh
# (public/). Puts the real Microsoft Core Fonts (Arial / Times New Roman /
# Courier New) into the given build artifact so the editor renders them instead
# of the Liberation clones (matching OnlyOffice cloud).
#
#   bash build/mscorefonts-step.sh <artifact-dir>     # default: fetch + enable
#   SK_SKIP_MSCOREFONTS=1 bash build/mscorefonts-step.sh <artifact-dir>  # Liberation-only
#
# ON BY DEFAULT (the product ships real Arial). Touches the artifact ONLY
# (<art>/fonts + <art>/editor-stubs.js flag) — never the committed source.
# PROPRIETARY fonts —
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
ART="${1:?usage: mscorefonts-step.sh <artifact-dir>}"
[[ "$ART" = /* ]] || ART="$REPO/$ART"

if [[ "${SK_SKIP_MSCOREFONTS:-}" == "1" ]]; then
  echo "→ MS Core Fonts SKIPPED (SK_SKIP_MSCOREFONTS=1) — Liberation-only build"
  # Force the artifact's runtime flag OFF so it can't 404 on absent fonts (the
  # committed source defaults ON for the always-present setup).
  SK_ENABLE_MSCOREFONTS=1 MSCORE_STUBS_FILE="$ART/editor-stubs.js" \
    bash "$SCRIPT_DIR/fetch-mscorefonts.sh" --revert
else
  echo "→ fetching Microsoft Core Fonts into $ART/fonts…"
  SK_ENABLE_MSCOREFONTS=1 MSCORE_FONTS_DIR="$ART/fonts" MSCORE_STUBS_FILE="$ART/editor-stubs.js" \
    bash "$SCRIPT_DIR/fetch-mscorefonts.sh"
fi
