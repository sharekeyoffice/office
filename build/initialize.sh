#!/bin/bash
# initialize.sh — one-shot DEV bootstrap.
#
# Builds a runnable static site into public/ from scratch (clones pinned
# upstream into vendor/, builds editor bundles + web-apps, assembles + prunes,
# bakes a dev main app origin into edit.html). Then `npm start` serves it.
#
# This is the DEV path (→ public/). Production uses build/build.sh (→ dist/)
# + Docker/k8s. This script does NOT deploy anything.
#
# Usage:
#   npm run initialize
#   ALLOWED_HOST_ORIGIN=https://app.example.com npm run initialize   # custom origin
#   SKIP_PRUNE=1 npm run initialize                                   # faster, larger public/
#
# Idempotent: safe to re-run (01 resets vendor to the pinned ref each time).

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
cd "$REPO"

export OUT_DIR="$REPO/public"
export ALLOWED_HOST_ORIGIN="${ALLOWED_HOST_ORIGIN:-http://localhost:3000}"

echo "════════════════════════════════════════════════════════════"
echo "  sharekey-office — DEV initialize"
echo "  output dir : $OUT_DIR"
echo "  baked origin: $ALLOWED_HOST_ORIGIN"
echo "               (override: ALLOWED_HOST_ORIGIN=… npm run initialize)"
echo "════════════════════════════════════════════════════════════"

bash "$SCRIPT_DIR/01-pull-upstream.sh"
bash "$SCRIPT_DIR/02-build-bundles.sh"
bash "$SCRIPT_DIR/03-deploy-web-apps.sh"
bash "$SCRIPT_DIR/04-assemble-dist.sh"

# Microsoft Core Fonts — fetched into public/ BY DEFAULT (same as the prod
# build) so local dev renders real Arial/Times/Courier and never 404s on a
# flag-on-but-fonts-absent mismatch. SK_SKIP_MSCOREFONTS=1 for Liberation-only.
# PROPRIETARY —
bash "$SCRIPT_DIR/mscorefonts-step.sh" "$OUT_DIR"

if [[ "${SKIP_PRUNE:-0}" == "1" ]]; then
  echo "→ skipping prune (SKIP_PRUNE=1)"
else
  bash "$SCRIPT_DIR/05-prune.sh"
fi

echo
echo "════════════════════════════════════════════════════════════"
echo "  ✓ public/ ready — $(du -sh "$OUT_DIR" | cut -f1)"
echo "════════════════════════════════════════════════════════════"
echo "Next:"
echo "  npm start            # editor on http://localhost:8080 (serves public/)"
echo "  npm run test:host    # main app simulator on http://localhost:3000"
echo
echo "  edit.html origin is baked to: $ALLOWED_HOST_ORIGIN"
echo "  to test against your real main app, re-run with that origin, e.g.:"
echo "    ALLOWED_HOST_ORIGIN=https://app.example.com npm run initialize"
