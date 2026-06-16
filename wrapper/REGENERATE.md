# Regenerating the wrapper from upstream (sdkjs + web-apps + cryptpad x2t)

This document does **two jobs at once**:

1. **Upgrade runbook** — how to rebuild the entire `wrapper/` foundation
   from clean upstream sources when bumping OnlyOffice.
2. **AGPL §13 corresponding-source documentation** — the "scripts used to
   control compilation and installation" that we must be able to hand to
   anyone using `office.*.example.com`, plus the provenance of every
   third-party artifact we redistribute.

> **The principle:** `wrapper/` is **regenerated from upstream**, never
> hand-copied. Everything here either (a) comes verbatim from a pinned
> upstream release, (b) is produced by a documented transform of an
> upstream file, or (c) is our own original source whose purpose is
> documented (these are our "stated modifications" under the OnlyOffice
> 9.4 §7 additional terms). If you can't trace an artifact to one of
> those three, it does not belong in `wrapper/`.

See also: the internal build-chain notes (two-tier architecture), [LICENSE](../LICENSE),
[NOTICE](../NOTICE), [THIRD-PARTY.md](../THIRD-PARTY.md).

## Pinned inputs

All three are pinned in [`upstream-pins/`](../upstream-pins/):

| Input | Repo | Pin file | License |
|---|---|---|---|
| sdkjs | `github.com/ONLYOFFICE/sdkjs` | `sdkjs.json` (`ref`) | AGPL-3.0 + ONLYOFFICE §7 terms (9.4+) |
| web-apps | `github.com/ONLYOFFICE/web-apps` | `web-apps.json` (`ref`) | AGPL-3.0 + ONLYOFFICE §7 terms (9.4+) |
| x2t (WASM) | `github.com/cryptpad/onlyoffice-x2t-wasm` | `x2t.json` (`url`+`sha512_url`) | AGPL-3.0 |

> **Pin reality check (2026-06):** the working sdkjs/web-apps clones track
> **upstream `ONLYOFFICE/*` directly** (no private fork, no local commits;
> sdkjs HEAD = `v9.3.1.11`). The old `git@github.com:<your-org>/sdkjs.git @
> v9.3.0+0` pin was fictional. Pin **real upstream tags**. Match
> `sdkjs`↔`web-apps` major.minor; match x2t's major to sdkjs.
>
> **Version choice (9.3.x vs 9.4+):** 9.4 added AGPL §7 *additional terms*
> requiring prominent in-UI ONLYOFFICE credit (see [NOTICE](../NOTICE) and
> the compliance section below). 9.3.x predates that contractual clause but
> is **still** bound by AGPL-3.0 §13 and the ONLYOFFICE Trademark Policy,
> and lacks 9.4's removal of the 20-connection limit. Default to **9.4+ and
> comply**; only stay ≤9.3.x on legal advice. **Get counsel** — ONLYOFFICE
> is actively enforcing this against debranded redeployments (the
> "Euro-Office" case).

## Artifact provenance + how to regenerate each

Legend: **UPSTREAM-VERBATIM** · **DERIVED** (documented transform of an
upstream file) · **ORIGINAL** (our source — not regenerated, purpose
documented) · **VENDORED-BINARY** (pinned third-party binary).

### A. Editor bundles — `wrapper/v1/bundle/*.editor.bundle.js{,.gz,.br}`  · DERIVED

Produced by [`wrapper/build/bundle.js`](build/bundle.js) from sdkjs source.

```bash
# from sharekey-office/, after 01-pull-upstream.sh has cloned vendor/sdkjs:
# ⚠️ PREREQUISITE: stage our hand-written AllFonts.js into the sdk tree FIRST.
# A clean upstream clone has NO common/AllFonts.js (OnlyOffice generates it at
# build time); the manifest references it, so the bundle is wrong (and ~1.4 KB
# short per editor) without this copy. A local dev sdkjs dir masks this because
# the file is already sitting there.
cp wrapper/v1/sdk-runtime/common/AllFonts.js vendor/sdkjs/common/AllFonts.js

cd wrapper
node build/bundle.js --mode editor --editor word  --sdk-dir ../vendor/sdkjs
node build/bundle.js --mode editor --editor cell  --sdk-dir ../vendor/sdkjs
node build/bundle.js --mode editor --editor slide --sdk-dir ../vendor/sdkjs
# → wrapper/v1/bundle/{word,cell,slide}.editor.bundle.js (+ .gz/.br)
```

- `--sdk-dir` is the **new flag** (see "bundle.js change" below) that points
  the concatenator at the upstream sdkjs checkout instead of assuming the
  parent dir. This is what lets `wrapper/` live outside the sdkjs tree.
- These outputs are **gitignored** (regenerated every build). They carry
  ONLYOFFICE-copyrighted code → see the **legal-comments caveat** below.

### B. Script manifests — `wrapper/scripts-{word,cell,slide}.js[.viewer.js]` · DERIVED

The full edit-mode manifest is a reformat of upstream
`sdkjs/configs/{word,cell,slide}.json` (the file load-order list).

```
Transform (document precisely when you next regenerate):
  sdkjs/configs/word.json  ──(reformat to one "../path" per line)──►  scripts-word.js
  scripts-word.js          ──(slim-manifest.py EXCLUSIONS)─────────►  scripts-word.viewer.js
```

On upgrade: re-pull `configs/*.json`, re-derive `scripts-*.js`, then run
`python3 wrapper/v2/slim-manifest.py` to regenerate the `.viewer.js` slim lists
+ `build/exclusions-*.json` audit logs. Watch `bundle.js` for
`[warn] missing:` — that means upstream renamed a file the manifest lists.

### C. sdkjs patch — `wrapper/sdkjs-patches/0001-externals-plainfonts.patch` · DERIVED

The single source change we make to upstream sdkjs: the `__plainFonts`
XOR-skip in `common/Drawings/Externals.js` (so our plain TTFs aren't
XOR-mangled). Capture it as a patch (one-time), then apply at build:

```bash
# capture once, from the sdkjs checkout that has the edit:
git -C <sdkjs> diff common/Drawings/Externals.js > wrapper/sdkjs-patches/0001-externals-plainfonts.patch
# apply at build (in 01-pull-upstream.sh, after checkout):
git -C vendor/sdkjs apply ../wrapper/sdkjs-patches/0001-externals-plainfonts.patch
```

If a future upstream merges/reverts this, refresh the patch. This patch is
a **modification of ONLYOFFICE code** → it must be disclosed as part of
corresponding source and noted in [NOTICE](../NOTICE).

### D. Font manifest — `wrapper/v1/sdk-runtime/common/AllFonts.js` · ORIGINAL

**Hand-written**, not derived from upstream's generated `AllFonts.js`. It
declares `window.__fonts_files` / `__fonts_infos` for exactly the 25 TTFs
we bundle (see its header for the row format). Regenerate only when the
bundled font set changes; keep the indexes in sync with `fonts/` (see
runbook 07). It is our original work.

> ⚠️ **It is also a BUILD INPUT, not just a runtime asset.** The bundle
> manifest lists `../common/AllFonts.js`, so `bundle.js` embeds it into
> `sdk-all-min.js`. A clean upstream sdkjs clone does **not** contain
> `common/AllFonts.js` → the bundle is wrong without staging our copy into
> `vendor/sdkjs/common/AllFonts.js` before building (see §A). This is the one
> hidden dependency a local dev sdkjs dir silently satisfies. `01-pull-upstream.sh`
> must perform this copy after checkout (along with applying the patches).

### E. libfont WASM engine — `wrapper/v1/sdk-runtime/common/libfont/engine/{fonts.js,fonts.wasm}` · UPSTREAM-VERBATIM

Copied as-is from `sdkjs/common/libfont/engine/` at the pinned ref
(HarfBuzz/FreeType shaping engine). ONLYOFFICE-copyrighted, AGPL-3.0.

```bash
cp vendor/sdkjs/common/libfont/engine/fonts.js   wrapper/v1/sdk-runtime/common/libfont/engine/
cp vendor/sdkjs/common/libfont/engine/fonts.wasm wrapper/v1/sdk-runtime/common/libfont/engine/
```

### F. x2t — `wrapper/v1/x2t/{x2t.wasm,x2t.js}{,.br}` · VENDORED-BINARY

CryptPad's prebuilt OnlyOffice x2t WASM (AGPL-3.0). **Committed** (per the
self-contained/offline-build decision), but its canonical source is the
pinned release:

```bash
URL=$(node -e "console.log(require('./upstream-pins/x2t.json').url)")
curl -fsSL "$URL" -o /tmp/x2t.zip
# verify against the published sha512 (currently NOT enforced by 01-pull — fix that):
curl -fsSL "$(node -e "console.log(require('./upstream-pins/x2t.json').sha512_url)")" -o /tmp/x2t.zip.sha512
( cd /tmp && shasum -a 512 -c x2t.zip.sha512 )
unzip -oq /tmp/x2t.zip -d wrapper/v1/x2t/
```

Corresponding-source for x2t lives in the cryptpad repo at the pinned tag —
link it from [THIRD-PARTY.md](../THIRD-PARTY.md).

### G. Bundled fonts — `wrapper/v1/fonts/*.ttf` (25 files) · VENDORED-BINARY

Substitution fonts for Microsoft families. Each has an upstream source +
license; full table + download commands in [THIRD-PARTY.md](../THIRD-PARTY.md)
(migrated from the old `fonts/README.md`). Summary:

| Family | Substitutes | Source | License |
|---|---|---|---|
| Liberation Sans/Serif/Mono (12) | Arial/Times/Courier | liberationfonts 2.1.5 | OFL-1.1 |
| Carlito (4) | Calibri | googlefonts/carlito | OFL-1.1 |
| Caladea (4) | Cambria | googlefonts/caladea | OFL-1.1 (verify) |
| DejaVu Sans (4) | broad Unicode | dejavu-fonts | Bitstream-Vera/permissive |
| OpenSymbol (1) | Symbol/Wingdings/Webdings | LibreOffice | verify (LGPL/MPL) |

> 🐞 **Compliance gap (fix during migration):** the font **LICENSE files are
> not currently shipped** in `fonts/`. OFL-1.1 **requires** the license text
> to accompany the fonts. Ship each family's license alongside its TTFs
> (`fonts/LICENSE-<family>.txt`) and list them in THIRD-PARTY.md.

### H. Wrapper glue — `wrapper/v1/{wrapper-*.js, editor-stubs.js, stubs.js, x2t-bridge.js, edit*.html, dev-server.js, web-apps-overlay/*}` · ORIGINAL

Our original source — **not regenerated**. `editor-stubs.js` is the 12
cascade patches (its targets are re-verified per bump via the runbook's
`verify-stubs.sh`); `x2t-bridge.js` wraps cryptpad x2t (itself AGPL-licensed
— see its header); the rest is wrapper plumbing. Document *purpose* per file
(done in the internal build-chain notes); this purpose
list also serves as our **"stated modifications"** for the §7 terms.

## Required code change: `bundle.js --sdk-dir`

Today `bundle.js` derives the sdkjs root as `path.resolve(__dirname,'..','..')`
([bundle.js:36-37](build/bundle.js)) — i.e. it assumes it lives *inside*
sdkjs. To run from `wrapper/`, add a `--sdk-dir` flag (and/or `OO_SDK_DIR`
env) that overrides `SDK_DIR`, while keeping the **manifest dir** resolved
relative to `bundle.js` itself (so `scripts-*.js` are read from `wrapper/`).
~5 lines. Output dir stays `wrapper/v1/bundle/`.

## End-to-end regeneration (clean room)

```bash
cd sharekey-office
bash build/01-pull-upstream.sh      # clone sdkjs+web-apps @ pins → vendor/; THEN:
                                    #   git -C vendor/sdkjs apply wrapper/sdkjs-patches/*.patch
                                    #   cp wrapper/v1/sdk-runtime/common/AllFonts.js vendor/sdkjs/common/  ← REQUIRED (§A/§D)
                                    #   (x2t is committed in wrapper/v1/x2t)
bash build/02-build-bundles.sh      # node wrapper/build/gen-manifest.js --sdk-dir vendor/sdkjs --verify
                                    # node wrapper/build/bundle.js --mode editor --sdk-dir ../vendor/sdkjs (× word/cell/slide)
#   E,G already in wrapper/ (committed); B,C,D-staging done by 01/02; H is ORIGINAL (in repo)
bash build/03-deploy-web-apps.sh    # grunt deploy-common-component + per-editor; inject-boot.sh
bash build/04-assemble-dist.sh      # assemble dist/ + overlay + __ALLOWED_HOST_ORIGIN__
bash build/05-prune.sh
```

> **Verified 2026-06-02:** steps through `02` reproduce the deployed
> `sdk-all-min.js` bundles **byte-for-byte from a clean `vendor/sdkjs` clone**
> (see Equivalence proof below). Steps `03+` (web-apps grunt) are not
> byte-deterministic.

**Equivalence check (first time only):** diff the regenerated tree against
the current hand-copied `viewerPoc/` to prove nothing undocumented was lost:

```bash
diff -rq <old>/viewerPoc/v1 sharekey-office/wrapper/v1 \
  | grep -v -E '/(bundle|x2t)/|SIZES.txt'   # expect: only generated/binary deltas
```

Any structural diff = an undocumented local tweak → document it here or drop it.

### Equivalence proof — CLEAN upstream clone (run 2026-06-02)

Done twice. First against the local dev sdkjs; then — the real test —
against a **fresh `git fetch` of `github.com/ONLYOFFICE/sdkjs` @ `b2f0aa1d5c`
into `vendor/sdkjs`** (no local sdkjs source used), patch applied, AllFonts
staged, manifests regenerated, bundles built via `wrapper/build/bundle.js`:

| step | result |
|---|---|
| `gen-manifest.js --verify` (from vendor/sdkjs) | ✅ all 3 manifests regenerate byte-exact (438/399/386) |
| bundles vs deployed `public/sdkjs/<ed>/sdk-all-min.js{,.gz,.br}` | ✅ **9/9 byte-identical** |

→ **The editor engine (`sdk-all-min.js`) reproduces byte-for-byte from a
clean upstream clone**, independent of any local sdkjs checkout. Inputs that
made it deterministic: exact upstream commit + `sdkjs-patches/*` + staged
`AllFonts.js` + committed manifests + same esbuild version.

**Hidden-dependency caught by the clean-room run:** the first clean build was
~1.4 KB short per editor and warned `missing: /common/AllFonts.js`. A clean
clone has no `common/AllFonts.js`; the bundle needs our hand-written one
staged in (§D). The local dev dir masked this. Now documented as a required
build step.

> **What this proof does NOT cover:** the `web-apps/` + non-bundle `sdkjs/`
> portion of `public/` comes from a grunt build of web-apps, which is **not
> byte-deterministic** (auto-incrementing build number, minifier output) — so
> a full `public/` byte-diff is not achievable by design. That half is
> verified *structurally* (file presence/paths), not byte-for-byte. The
> wrapper glue + fonts + x2t parts of `public/` are verbatim copies (trivially
> identical).

### Full-pipeline verification — whole `public/` from scratch (run 2026-06-03)

Ran the **entire pipeline** (`01`→`05`) end-to-end from clean upstream clones
into a fresh `public/` (origin baked `http://localhost:3000`), with `vendor/`
wiped first — **no local sdkjs, no `viewerPoc`**. Exit 0, ~4.5 min
(clone 50s · bundles 75s · web-apps grunt incl. `npm install` ~2 min ·
assemble+prune ~15s). Diffed against the prior `public_old`:

| Category | Result |
|---|---|
| `sdkjs/<ed>/sdk-all-min.js{,.gz,.br}` (×9) | ✅ byte-identical |
| wrapper glue, `AllFonts.js`, libfont `fonts.wasm`, x2t, 25 fonts, `edit.html` | ✅ byte-identical |
| `spreadsheeteditor`/`presentationeditor` `app.js` | ✅ byte-identical |
| `api.js`, `documenteditor/main/app.js` | ⚠️ differ **only by grunt's build counter** (`build:40` vs `41`); identical size — inherent, harmless |
| `document_editor_service_worker.js` | trivial stub-text difference |
| new **adds** | font license files + `edit-host-demo.html` |
| old had extra | `help/`, `mobile/`, `embed/`, `ie/`, `*.map` — correctly pruned in new |

**Conclusion: the from-scratch build is provably equivalent to the original**
(modulo the auto-incrementing grunt build number), with zero dependence on any
local sdkjs/viewerPoc checkout.

> 🐞 **Prune bug found & fixed during this run.** `05-prune.sh` trimmed
> `locale/<lang>/` *directories*, but web-apps 9.3 ships `locale/<lang>.json`
> *files* → the locale trim was a silent no-op (45 locales survived instead of
> just `en`). Now handles both forms and also removes per-editor `forms/`.
> Effect: `public/` 229 MB → 193 MB. (The original `public_old` only had `en`
> because it was built via the *other* slimming path — `apply-overlay` en-only
> — not the prune path.)

The generator (`bundle.js`), manifests, and hand-authored glue are *not*
derivable from the design notes (see gap below) — they are preserved as source in
`wrapper/`; only their **outputs** regenerate.

### Gap: the design notes cannot reconstruct the original source

The internal design notes documents the wrapper's *purpose, patch
sites, and procedures* but contains **no shippable source** (its README
says so explicitly). Therefore these must be **preserved as version-
controlled source in `wrapper/`** — they cannot be regenerated from the
those notes or from upstream:

- the generator + build data: `build/bundle.js`, `build/slim-manifest.py`,
  `build/exclusions-*.json`, `scripts-{word,cell,slide}.js[.viewer.js]`
- the hand-authored glue: `v1/{editor-stubs.js, wrapper-*.js, x2t-bridge.js,
  stubs.js, edit*.html, dev-server.js, web-apps-overlay/*}`
- the hand-written `v1/sdk-runtime/common/AllFonts.js`

Until they live in `wrapper/` under git, the **only** copy is the untracked
`…/sdkjs/.claude/worktrees/…/viewerPoc/` working dir — a single-point-of-
failure. Committing them into `wrapper/` is step 1 of the migration.

> Manifest derivation: **resolved.** `wrapper/build/gen-manifest.js` now
> derives `scripts-{word,cell,slide}.js` from `configs/*.json`
> (`PREPEND + .sdk.min + .sdk.common`, plus a single `common/zlib/zlib.js`
> insert for cell/slide) and verifies byte-exact against the committed
> manifests. Run `node build/gen-manifest.js --sdk-dir <sdk> [--write|--verify]`.

## Compliance checklist (run every release — see NOTICE / THIRD-PARTY.md)

- [ ] **§13 corresponding source** reachable from the served site (a `/source`
      link or in-editor "Source code" entry) → upstream pins + `wrapper/` +
      `sdkjs-patches/` + these build scripts.
- [ ] **§7 attribution (9.4+)**: prominent in-UI "Powered by ONLYOFFICE
      (modified by Sharekey)" + "ONLYOFFICE is a trademark of Ascensio System
      SIA" + license link (baked into `overlay/edit.html`).
- [ ] **Modifications stated**: `NOTICE` lists our changes + dates +
      "based on ONLYOFFICE by Ascensio System SIA."
- [ ] **License notices preserved**: see the legal-comments caveat below.
- [ ] **Trademark**: product/brand is "Sharekey", not "ONLYOFFICE"; no
      ONLYOFFICE logo/trade-dress as our branding; nominative use only.
- [ ] **Fonts**: ship every family's license file (gap H/G above).
- [ ] **Icons (CC BY-SA 4.0)**: attribute in THIRD-PARTY.md; any icon
      derivatives stay CC BY-SA 4.0.

### ⚠️ legal-comments caveat

`bundle.js` runs esbuild with `legalComments:'none'` ([bundle.js:127](build/bundle.js)),
which **strips ONLYOFFICE copyright/license headers from the shipped
bundles**. AGPL + the §7 terms require keeping legal notices. Resolve by
either (a) `legalComments:'inline'`/`'eof'` to retain them, or (b) serving a
`NOTICE`/`THIRD-PARTY` document alongside that reproduces the stripped
headers, with a header pointer in each bundle. Decide and record the choice
here before shipping.
