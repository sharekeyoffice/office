# `wrapper/v2/` — future read-only viewer mode

**Staging area for the planned v2: a strictly read-only viewer** — render
documents only, with **no editing** of any kind (no toolbars/edit surface).
Where `v1/` is the full editor (web-apps shell + the ~12 cascade patches),
`v2` will be the minimal viewer that loads the SDK to *display* a document and
nothing more.

> Status: **not built yet.** This folder currently holds the viewer-mode
> artifacts carried over from the original v0 PoC, which are the seeds of v2.

## What's here

| File | What it is |
|---|---|
| `stubs.js` | No-op replacements for SDK classes a viewer never needs (AscBuilder, MacroRecorder, plugin host, …). Loaded **before** the SDK so unconditional `new AscCommon.X()` calls resolve. Lets the slim viewer boot without pulling in editing/collab/plugin/builder code. |
| `scripts-{word,cell,slide}.viewer.js` | **Viewer-mode bundle manifests** — the slimmed file lists (vs the full `v1/`-mode `scripts-*.js`). Produced by trimming the editor manifests (round-1 exclusions + round-2 stubbed files). |

## How v2 will build (when implemented)

The same generator builds viewer bundles, pointed at this dir:

```bash
node wrapper/build/bundle.js --mode viewer --editor word \
  --sdk-dir vendor/sdkjs --manifest-dir wrapper/v2 --out-dir wrapper/v2/bundle
```

- `--mode viewer` reads `scripts-{editor}.viewer.js` (here) and emits
  `{editor}.bundle.js` (not `.editor.bundle.js`).
- `stubs.js` must be loaded as a `<script>` **before** the bundle.

> `slim-manifest.py` (in this dir) is the generator that produced the
> `scripts-*.viewer.js` here, from the full editor manifests
> (`../scripts-{editor}.js`). ⚠️ It's a v0-era script — its internal paths still
> assume the old `viewerPoc/` layout, so **update them before running** when v2
> is actually built. Until then the committed `.viewer.js` are a fixed snapshot.

## Not part of the v1 deployment

`sharekey-office` currently ships **only v1** (the editor). Nothing in `v2/`
is referenced by the build (`build/01-05`) or the served `edit.html` today.
