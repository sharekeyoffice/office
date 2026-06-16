#!/usr/bin/env python3
"""
slim-manifest.py — generate viewer-only script manifests by excluding files
known to be editor-only / UI-only.

Reads:  viewerPoc/scripts-{editor}.js   (current full manifest)
Writes: viewerPoc/scripts-{editor}.viewer.js   (slimmed manifest)
        viewerPoc/build/exclusions-{editor}.json   (audit log)

Usage:
    python3 viewerPoc/build/slim-manifest.py [--dry-run]
"""
import os, re, sys, json, argparse
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
VIEWER = ROOT / "viewerPoc"

# Each rule is (pattern_substring_in_path, reason). Conservative — only
# things that are unambiguously editor-only / view-irrelevant.
# Patterns match anywhere in the path (substring, case-sensitive).
EXCLUSIONS = [
    # ===== Round 2: stubbed in viewerPoc/stubs.js =====
    # `viewerPoc/stubs.js` defines no-op replacements for these classes BEFORE
    # the SDK manifest loads. The real files can therefore be dropped.

    # AscBuilder.{Word,Cell,Slide}.init() called from each editor's _onEndLoadSdk.
    # 2.5 MB combined — biggest single win in round 2.
    ("/apiBuilder.js",                "stubbed: AscBuilder.{Word,Cell,Slide} (see stubs.js)"),

    # AscCommon.MacroRecorder (apiBase.js:210) and CDocumentMacros (apiBase.js:3316).
    ("common/macro-recorder.js",      "stubbed: AscCommon.MacroRecorder (see stubs.js)"),
    ("common/macros.js",              "stubbed: AscCommon.CDocumentMacros (see stubs.js)"),

    # Plugin host: ~170 KB combined. Asc.createPluginsManager (apiBase.js:3314)
    # and CPluginCtxMenuInfo (apiBase.js:5317) — stubbed.
    ("common/apiBase_plugins.js",     "stubbed: Asc.createPluginsManager + CPluginCtxMenuInfo"),
    ("/api_plugins.js",               "stubbed: per-editor plugin host (no plugins in viewer)"),
    # ===== End round 2 stubs =====

    # ---- Plugin host (no plugins in the viewer) ----
    # NOTE: keep `apiBase_plugins.js` and the per-editor `api_plugins.js` for
    # now — they're loaded into baseEditorsApi.prototype and removing the
    # mixin can crash unrelated code paths (similar to CDocsCoApi). Phase 6.2
    # candidate for proper stubbing.
    ("plugin-events.js",              "plugin event bridge — no plugins"),

    # ---- Macro recording ----
    # NOTE: keep both `macro-recorder.js` AND `macros.js` — apiBase.js:210
    # instantiates `new MacroRecorder()`, and apiBase.js:3316 (during
    # _onEndLoadSdk) does `new AscCommon.CDocumentMacros()` from macros.js.
    # Both are round-2 stub candidates.

    # ---- Spell check ----
    # NOTE: keep `spellcheckapi.js` — apiBase.js:224 instantiates CSpellCheckApi
    # unconditionally. The actual spell engine in `common/spell/` IS droppable
    # because it's only triggered by API calls we never make.
    ("common/spell/",                 "spell-check engine — viewer-irrelevant"),

    # ---- Collaboration / coauth ----
    # All five collab files are tightly coupled and unconditionally
    # instantiated during _onEndLoadSdk. Drop NOTHING here in round 1; the
    # whole subtree is a round-2 stub-bundle target.
    #   - docscoapicommon.js      → defines c_oEditorId enum
    #   - docscoapi.js            → CDocsCoApi (apiBase.js:146)
    #   - CollaborativeEditingBase.js → base class
    #   - */CollaborativeEditing.js   → per-editor subclass (apiBase.js:3287)
    #   - collaborativeHistory.js → CCollaborativeHistory (instantiated by
    #     CollaborativeEditingBase constructor at line 227)

    # ---- Hashing / crypto used for change tracking ----
    # NOTE: keep `common/digest/sha256.js` — `AscCommon.Digest.sha256` is
    # called at runtime by the paragraph style cache (Metafile.js:580 →
    # style-cache.js → ParaRun constructor). Used on EVERY paragraph during
    # deserialization, not just change-tracking.
    # NOTE: keep `stringserialize.js` — defines `AscCommon.Base64` for fonts.
    ("common/hash/",                  "hashing for change tracking"),
    ("common/random.js",              "PRNG for ids — used by editing flows"),
    ("common/keychainstorage.js",     "credential keychain — no auth flow"),

    # ---- Builder / editor-facing public API surface ----
    # NOTE: keep `apiBuilder.js` per editor — each editor's `_onEndLoadSdk`
    # calls `AscBuilder.<Editor>.init()` unconditionally (e.g. word/api.js:10275).
    # Stubbed in stubs.js (round 2).
    # NOTE: keep `fromToJSON.js` — mixes `ToJson` onto AscWord.CNumberingLvl.prototype
    # etc. The numbering panel update path (CNumberingCollection.AddToMultiLvlCollection
    # → CNumInfo.FromLvl) calls `numLvl.ToJson()` during UpdateInterface on every
    # word document open. Without it, `ToJson is not a function` after the first
    # recalculate finishes.

    # ---- Autocorrect / settings UI ----
    # ALL four settings files in `common/api/` are kept in round 1:
    #   - spellCheckSettings.js   → CDocumentSpellChecker constructor uses it
    #   - autoCorrectSettings.js  → CDocument constructor (Document.js:1584)
    #   - addTextSettings.js      → CAddTextSettings used from many paths
    #   - firstLetterExceptions.js → instantiated INSIDE autoCorrectSettings.js
    #   - restrictionSettings.js  → safest to keep alongside
    # All round 2 stub candidates.

    # ---- Custom XML / extra UI ----
    # NOTE: keep `serialize-custom-xml.js` — defines `BinaryCustomsTableReader`
    # which the word bin reader's main-table walker (Serialize2.js:8094)
    # instantiates during `OpenDocumentFromBin`. Touched on every word open.
    # NOTE: keep `deleted-text-recovery.js` — `CCollaborativeHistory.InitTextRecover`
    # (collaborativeHistory.js:310) instantiates `new AscCommon.DeletedTextRecovery()`
    # during `Apply_Changes` on the open-end callback path.

    # NOTE about stringserialize.js: kept (was previously dropped) — it defines
    # `AscCommon.Base64` which `libfont/map.js:CreateFontData2` needs at boot.
    # The module's other concerns (string change-hashing) are dead code in
    # viewer mode but the file is small and removing the symbol cascades.
]


def matches(path: str, patterns) -> tuple[bool, str]:
    for pat, reason in patterns:
        if pat in path:
            return True, reason
    return False, ""


def slim_one(editor: str, dry_run: bool):
    src = VIEWER / f"scripts-{editor}.js"
    dst = VIEWER / f"scripts-{editor}.viewer.js"
    audit_path = VIEWER / "build" / f"exclusions-{editor}.json"

    body = src.read_text()
    paths = re.findall(r'"\.\.([^"]+)"', body)

    kept, dropped = [], []
    for p in paths:
        is_ex, reason = matches(p, EXCLUSIONS)
        if is_ex:
            dropped.append({"path": p, "reason": reason})
        else:
            kept.append(p)

    print(f"{editor:6} kept={len(kept):4d}  dropped={len(dropped):4d}  "
          f"({100 * len(dropped) / len(paths):.1f}% reduction)")

    if dry_run:
        return

    # Write slimmed manifest in the same shape as the original
    lines = ["var sdk_scripts = ["]
    for i, p in enumerate(kept):
        suffix = "," if i < len(kept) - 1 else ""
        lines.append(f'\t".."{suffix}'.replace('".."', f'"..{p}"'))
    lines.append("];")
    lines.append("")
    dst.write_text("\n".join(lines))

    # Audit log
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    audit_path.write_text(json.dumps({
        "editor": editor,
        "source_manifest": src.name,
        "slimmed_manifest": dst.name,
        "kept_count": len(kept),
        "dropped_count": len(dropped),
        "exclusion_rules_applied": len(EXCLUSIONS),
        "dropped": dropped,
    }, indent=2))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true",
                    help="report only; don't write files")
    args = ap.parse_args()

    for editor in ("word", "cell", "slide"):
        slim_one(editor, args.dry_run)


if __name__ == "__main__":
    main()
