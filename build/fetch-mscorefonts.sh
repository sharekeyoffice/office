#!/usr/bin/env bash
#
# PROTOTYPE — build-time fetch of Microsoft "Core fonts for the Web"
# (Arial / Times New Roman / Courier New). PENDING LEGAL SIGN-OFF.
#
# This mirrors exactly what OnlyOffice's DocumentServer does
# (ttf-mscorefonts-installer, EULA-accepted in their Dockerfile): it does NOT
# bundle any font in our source tree — it downloads Microsoft's ORIGINAL,
# unmodified installer archives at build time and extracts the .ttf locally.
# The extracted fonts land in a gitignored staging dir + public/fonts/ and are
# NEVER committed.
#
# ─────────────────────────────────────────────────────────────────────────────
# LICENSING — READ BEFORE ENABLING
# These are Microsoft's "Core fonts for the Web" (1996 EULA, program
# discontinued 2002). They are FREE OF CHARGE but the EULA only permits
# redistribution as Microsoft's original unmodified installer — NOT the
# extracted .ttf, and NOT repackaged into another product. Serving the
# EXTRACTED fonts from our public site to browsers (which is what the editor's
# canvas font engine needs) is the open question that requires LEGAL SIGN-OFF
# before this leaves prototype status.
# ─────────────────────────────────────────────────────────────────────────────
#
# Usage:
#   SK_ENABLE_MSCOREFONTS=1 build/fetch-mscorefonts.sh          # fetch + enable
#   SK_ENABLE_MSCOREFONTS=1 build/fetch-mscorefonts.sh --revert # disable (keeps files)
#
# Requires: curl, cabextract  (macOS: `brew install cabextract`)

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
STAGING="$REPO_ROOT/mscorefonts"          # gitignored staging cache (originals + extracted)
# Destinations are overridable so the same script serves dev (public/) and the
# production build (dist/). Defaults target dev. The production build sets
# MSCORE_FONTS_DIR=dist/fonts and MSCORE_STUBS_FILE=dist/editor-stubs.js so it
# NEVER flips the committed source flag — only the dist build artifact.
FONTS_DIR="${MSCORE_FONTS_DIR:-$REPO_ROOT/public/fonts}"   # gitignored in dev; served at /fonts/
STUBS="${MSCORE_STUBS_FILE:-$REPO_ROOT/wrapper/v1/editor-stubs.js}"

# Opt-in gate — refuse to download proprietary fonts unless explicitly asked.
if [[ "${SK_ENABLE_MSCOREFONTS:-}" != "1" ]]; then
  echo "Refusing to run: this downloads PROPRIETARY Microsoft fonts."
  echo "This is a prototype pending legal sign-off. To proceed deliberately:"
  echo "  SK_ENABLE_MSCOREFONTS=1 $0"
  exit 2
fi

flip_flag() { # $1 = true|false — toggle the runtime gate in editor-stubs.js
  if grep -q "var SK_REAL_MSCOREFONTS" "$STUBS"; then
    # portable in-place sed (BSD/macOS + GNU)
    sed -i.bak -E "s/(var SK_REAL_MSCOREFONTS = )(true|false)(;)/\1$1\3/" "$STUBS" && rm -f "$STUBS.bak"
    echo "→ editor-stubs.js: SK_REAL_MSCOREFONTS = $1"
  else
    echo "WARN: SK_REAL_MSCOREFONTS flag not found in editor-stubs.js (font repoint won't toggle)."
  fi
}

if [[ "${1:-}" == "--revert" ]]; then
  flip_flag false
  echo "Reverted. Real MS fonts stay on disk (gitignored) but the editor uses Liberation again."
  echo "Run 'npm run sync' to redeploy editor-stubs.js, then hard-reload the editor."
  exit 0
fi

command -v curl       >/dev/null || { echo "ERROR: curl not found";       exit 1; }
command -v cabextract >/dev/null || { echo "ERROR: cabextract not found (brew install cabextract)"; exit 1; }

mkdir -p "$STAGING/dl" "$STAGING/ttf" "$FONTS_DIR"

# Microsoft's original installers, via the long-standing SourceForge 'corefonts'
# mirror (the same source ttf-mscorefonts-installer uses). Each is a self-
# extracting CAB containing the unmodified Microsoft .ttf files.
MIRROR="https://downloads.sourceforge.net/corefonts"
INSTALLERS=(arial32.exe times32.exe courie32.exe)

echo "Downloading Microsoft Core Fonts installers (originals, unmodified)…"
for exe in "${INSTALLERS[@]}"; do
  if [[ ! -f "$STAGING/dl/$exe" ]]; then
    echo "  • $exe"
    curl -sSL --fail -o "$STAGING/dl/$exe" "$MIRROR/$exe"
  fi
  cabextract -L -q -d "$STAGING/ttf" "$STAGING/dl/$exe" >/dev/null   # -L => lowercase names
done

# Copy each face to the canonical filename editor-stubs.js expects, matching
# case-insensitively against the CAB's inner names.
copy_face() { # $1 = source glob (lowercased) ; $2 = destination filename
  local src
  src="$(find "$STAGING/ttf" -maxdepth 1 -type f -iname "$1" | head -1)"
  # Hard fail (not warn): a missing face here would ship a build whose runtime
  # repoint requests $2 and 404s → editor error -26 "Fonts are not loaded". Fail
  # loudly at build time instead. (Inner CAB names are stable, but a bad mirror
  # download / different installer version would surface here.)
  if [[ -z "$src" ]]; then
    echo "  ERROR: '$1' not found in extracted fonts — cannot produce $2." >&2
    echo "         The installer download may have failed or returned an unexpected file." >&2
    echo "         Extracted so far: $(ls "$STAGING/ttf" 2>/dev/null | tr '\n' ' ')" >&2
    exit 1
  fi
  cp "$src" "$FONTS_DIR/$2"
  echo "  ✓ $2  ($(basename "$src"))"
}

echo "Installing real faces into $FONTS_DIR …"
copy_face 'arial.ttf'    'arial.ttf'        # Arial Regular
copy_face 'ariali.ttf'   'ariali.ttf'       # Arial Italic
copy_face 'arialbd.ttf'  'arialbd.ttf'      # Arial Bold
copy_face 'arialbi.ttf'  'arialbi.ttf'      # Arial Bold Italic
copy_face 'times.ttf'    'times.ttf'        # Times New Roman Regular
copy_face 'timesi.ttf'   'timesi.ttf'       # Times New Roman Italic
copy_face 'timesbd.ttf'  'timesbd.ttf'      # Times New Roman Bold
copy_face 'timesbi.ttf'  'timesbi.ttf'      # Times New Roman Bold Italic
copy_face 'cour.ttf'     'cour.ttf'         # Courier New Regular
copy_face 'couri.ttf'    'couri.ttf'        # Courier New Italic
copy_face 'courbd.ttf'   'courbd.ttf'       # Courier New Bold
copy_face 'courbi.ttf'   'courbi.ttf'       # Courier New Bold Italic

flip_flag true

cat <<EOF

Done (PROTOTYPE). Real Arial / Times New Roman / Courier New are now in
  $FONTS_DIR
and the runtime repoint flag is ENABLED in
  $STUBS

EOF
# The sync/reload hint only applies to the dev (public/) target. The production
# build wires this step in after dist is assembled, so no sync is needed there.
if [[ "$FONTS_DIR" == "$REPO_ROOT/public/fonts" ]]; then
  cat <<EOF
Next (dev):
  npm run sync          # redeploy editor-stubs.js to public/
  # hard-reload the editor tab → Arial now renders as real Arial

To disable:   SK_ENABLE_MSCOREFONTS=1 $0 --revert  &&  npm run sync
EOF
fi
