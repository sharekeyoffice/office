#!/bin/bash
# build.sh — orchestrator. Runs all build steps in order.
#
# Usage:
#   bash build/build.sh
#
# Steps:
#   01-pull-upstream.sh    — clone/pull pinned sdkjs + web-apps
#   02-build-bundles.sh    — run viewerPoc/build/bundle.js for editor mode
#   03-deploy-web-apps.sh  — grunt deploy + inject-boot
#   04-assemble-dist.sh    — copy artifacts into dist/
#   05-prune.sh            — remove unneeded fat
#   06-compress.sh         — write .gz siblings for nginx's gzip_static
#
# Final step: copy overlay/* over dist/.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$REPO"

echo "════════════════════════════════════════════════════════════"
echo "  sharekey-office build"
echo "  $(date -u +%Y-%m-%dT%H:%M:%SZ)"
echo "════════════════════════════════════════════════════════════"

bash "$SCRIPT_DIR/01-pull-upstream.sh"
bash "$SCRIPT_DIR/02-build-bundles.sh"
bash "$SCRIPT_DIR/03-deploy-web-apps.sh"
bash "$SCRIPT_DIR/04-assemble-dist.sh"   # also applies overlay + substitutes __ALLOWED_HOST_ORIGIN__

# Microsoft Core Fonts — fetched BY DEFAULT so the editor renders real
# Arial / Times New Roman / Courier New (matching OnlyOffice cloud). For a
# license-clean Liberation-only build, set SK_SKIP_MSCOREFONTS=1. Touches the
# build artifact ONLY (fonts → <out>/fonts, flag → <out>/editor-stubs.js), never
# the committed source. PROPRIETARY fonts —
bash "$SCRIPT_DIR/mscorefonts-step.sh" "${OUT_DIR:-$REPO/dist}"

bash "$SCRIPT_DIR/05-prune.sh"

# Pre-compress for nginx's gzip_static. After the prune, so we only spend CPU on
# files that actually ship.
bash "$SCRIPT_DIR/06-compress.sh"

echo
echo "════════════════════════════════════════════════════════════"
echo "  Build complete. Output:"
du -sh "$REPO/dist"
echo "════════════════════════════════════════════════════════════"
echo
echo "Next:"
echo "  npm start                         # local dev server on :8080"
echo "  npm run docker:build              # build container image"
