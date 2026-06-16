#!/usr/bin/env bash
# Apply the wrapper build-config overlay to a web-apps checkout.
#
# NOTE: NOT used by the office build. build/03-deploy-web-apps.sh builds the
# full web-apps and slims afterward via build/05-prune.sh. This script is the
# ALTERNATIVE pre-build slimming path (edits build/{editor}.json so grunt emits
# a slim deploy). Kept for reference / standalone web-apps builds.
#
# Patches build/{editor}.json files to:
#   1. Skip mobile/embed/forms variants (build only "main")
#   2. Trim main/locale copy glob to en.json
#   3. Reduce main/resources/help copy to a single stub file
#      (replace:prepareHelp aborts on empty help array — see 10.0.3 caveat)
#
# Idempotent — running twice yields the same on-disk state. Safe against
# upstream changes: each jq filter assigns deterministic values, never appends.
#
# Usage:
#   apply-overlay.sh [--web-apps PATH] [--editors LIST]
#
#   --web-apps PATH   Path to web-apps clone. Default: ../../../web-apps
#                     (relative to this script)
#   --editors LIST    Comma-separated. Default: documenteditor
#                     (10.1.2 will add spreadsheeteditor,presentationeditor)
#
# Exit codes: 0 ok, 1 missing prereqs, 2 patch failure.

set -euo pipefail

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Auto-detect a web-apps source checkout (dev convenience; pass --web-apps to
# be explicit — e.g. sharekey-office/vendor/web-apps).
WEB_APPS=""
for candidate in \
  "$SCRIPT_DIR/../../../../web-apps" \
  "$SCRIPT_DIR/../../../../../../web-apps" \
  "$SCRIPT_DIR/../../../web-apps" \
; do
  if [[ -d "$candidate/build" ]]; then
    WEB_APPS="$( cd "$candidate" && pwd )"
    break
  fi
done
EDITORS="documenteditor"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --web-apps) WEB_APPS="$2"; shift 2 ;;
    --editors)  EDITORS="$2"; shift 2 ;;
    -h|--help)
      sed -n '2,/^$/p' "$0" | sed 's/^# \?//'
      exit 0
      ;;
    *) echo "unknown arg: $1" >&2; exit 1 ;;
  esac
done

if ! command -v jq >/dev/null 2>&1; then
  echo "ERROR: jq not found in PATH" >&2; exit 1
fi
if [[ -z "$WEB_APPS" || ! -d "$WEB_APPS/build" ]]; then
  echo "ERROR: web-apps path not found at: $WEB_APPS" >&2
  echo "       Pass --web-apps PATH explicitly." >&2
  exit 1
fi

echo "web-apps: $WEB_APPS"
echo "editors:  $EDITORS"
echo

# jq filter shared by all editors. The help-array reduction keeps a single
# tiny entry (en/Contents.json) — empty array breaks replace:prepareHelp.
JQ_FILTER='
  .tasks.deploy = ["increment-build", "deploy-app-main"]
  | .main.copy.localization[0].src = "en.json"
  | .main.copy.help = [{
      "expand": true,
      "cwd":  ("../apps/" + .name + "/main/resources/help/"),
      "src":  ["en/Contents.json"],
      "dest": ("../deploy/web-apps/apps/" + .name + "/main/resources/help/")
    }]
'

IFS=',' read -r -a EDITOR_ARR <<< "$EDITORS"
for editor in "${EDITOR_ARR[@]}"; do
  cfg="$WEB_APPS/build/${editor}.json"
  if [[ ! -f "$cfg" ]]; then
    echo "  skip $editor (no $cfg)"; continue
  fi

  tmp="$(mktemp -t overlay.XXXXXX)"
  if jq "$JQ_FILTER" "$cfg" > "$tmp"; then
    # Atomic replace
    mv "$tmp" "$cfg"
    echo "  patched $editor.json"
  else
    rm -f "$tmp"
    echo "ERROR: jq failed on $cfg" >&2
    exit 2
  fi
done

echo
echo "Done. To rebuild:"
echo "  cd $WEB_APPS/build && npx grunt deploy-documenteditor"
echo "To revert (e.g. after an upstream pull) before re-applying:"
echo "  git -C $WEB_APPS checkout HEAD -- build/*.json"
