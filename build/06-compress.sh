#!/bin/bash
# 06-compress.sh — pre-compress the assembled site for nginx's `gzip_static on`.
#
# nginx serves <file>.gz when the client accepts gzip and that sibling exists;
# otherwise it sends the file raw. Before this step exactly three files in the
# whole build had a .gz — the sdk-all-min.js bundles, which 04 copies by hand —
# so everything else went over the wire uncompressed. The worst offender was
# x2t/x2t.wasm at ~34 MB: it ships a .br, but the image is stock nginx:alpine
# with no brotli module (see the commented-out brotli_static in
# docker/nginx.conf), so the .br was dead weight and the raw 34 MB was what
# actually travelled. That download, not the WASM compile, was the bulk of a
# cold editor open — measured 2337 ms against the deployed env versus 125 ms of
# compile time on localhost.
#
# Runs AFTER 05-prune.sh on purpose: prune drops ~485 MB of web-apps fat, and
# compressing before it would burn CPU on files that are about to be deleted and
# leave orphan .gz siblings behind.
#
# Operates on $OUT_DIR (default dist/).

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$SCRIPT_DIR/.." && pwd)"
DIST="${OUT_DIR:-$REPO/dist}"
[[ "$DIST" = /* ]] || DIST="$REPO/$DIST"

[[ -d "$DIST" ]] || { echo "ERROR: $DIST missing. Run 04 first." >&2; exit 1; }

# Text-ish and already-uncompressed binary formats. Deliberately NOT listed:
# png/jpg/webp/woff2 (already compressed — gzip would only add bytes and files).
# TTFs are included: they are uncompressed tables and shrink by roughly half,
# and the editor pulls 13 of them on every boot.
EXTENSIONS=(js css wasm ttf svg json html)

# Below this, the .gz is not worth the extra file — and for tiny inputs gzip can
# emit MORE bytes than the original, which nginx would then happily serve.
MIN_BYTES=1024

before=$(du -sh "$DIST" | cut -f1)

find_args=()
for ext in "${EXTENSIONS[@]}"; do
  find_args+=(-o -name "*.${ext}")
done
# Drop the leading -o so the group is a valid expression.
find_args=("${find_args[@]:1}")

# -f so a rebuild refreshes stale .gz siblings (e.g. the ones 04 copied for the
# sdk bundles); -k so the original stays — nginx needs both.
count=0
while IFS= read -r -d '' file; do
  gzip -9 -kf "$file"
  count=$((count + 1))
done < <(find "$DIST" \( "${find_args[@]}" \) -type f -size +"${MIN_BYTES}"c \
           ! -name '*.gz' ! -name '*.br' -print0)

echo "✓ compressed $count file(s)  (before: $before, after: $(du -sh "$DIST" | cut -f1))"

# Report the headline assets so a regression (a new fat asset with no .gz) is
# visible in the build log rather than only in a browser's network panel.
for f in x2t/x2t.wasm sdkjs/common/libfont/engine/fonts.wasm; do
  [[ -f "$DIST/$f" ]] || continue
  raw=$(wc -c < "$DIST/$f")
  gz=$([[ -f "$DIST/$f.gz" ]] && wc -c < "$DIST/$f.gz" || echo 0)
  awk -v r="$raw" -v g="$gz" -v n="$f" \
    'BEGIN { printf "    %s: %.1f MB → %.1f MB gzipped\n", n, r/1048576, g/1048576 }'
done
