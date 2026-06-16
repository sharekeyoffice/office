# `sdk-runtime/common/` — trimmed SDK runtime overrides

These are **our versions** of two sdkjs `common/` runtime files. At build time
they **overwrite** the upstream copies so the deployed editor uses our trimmed
font set + the matching shaping engine.

| File | What it is | Why it's here |
|---|---|---|
| `AllFonts.js` | **Hand-written** font manifest — declares `window.__fonts_files` + `__fonts_infos` for exactly the 25 TTFs we bundle (see `../../fonts/`). | Upstream/DocServer *generates* `AllFonts.js` from the server's font directory; we have no server, so we hand-write a manifest for our bundled fonts. |
| `libfont/engine/fonts.js` | HarfBuzz/FreeType text-shaping glue (loads `fonts.wasm`). | Upstream-verbatim copy, kept here so it ships alongside our `AllFonts.js`. |
| `libfont/engine/fonts.wasm` | The WASM shaping engine itself. | Same. |

## How it's used in the build

`AllFonts.js` is used **twice** (see `wrapper/REGENERATE.md`):

1. **As a build input** — `build/01-pull-upstream.sh` copies it into
   `vendor/sdkjs/common/AllFonts.js` **before** bundling, because the bundle
   manifest lists `../common/AllFonts.js` and a clean upstream clone has **no**
   `AllFonts.js`. Without this staging the editor bundle is ~1.4 KB short and
   fonts mis-resolve. (This is the one hidden dependency the clean-room rebuild
   caught.)
2. **As a served runtime asset** — `build/04-assemble-dist.sh` copies all three
   files over the grunt-deployed `dist/sdkjs/common/...`, so the browser loads
   our trimmed manifest + engine at runtime.

## When to touch this

Only when the **bundled font set changes** — then regenerate `AllFonts.js`
(keep the file indexes in sync with `../../fonts/`; see
the font-addition procedure). `fonts.js`/`fonts.wasm` only change
if upstream bumps the HarfBuzz/FreeType engine.
