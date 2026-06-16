# `x2t/` — OnlyOffice x2t document converter (WebAssembly)

`x2t` ("X to T") is OnlyOffice's file-format converter. This is a **WebAssembly
build** of it that runs **in the browser**, used by the wrapper to convert
between OOXML and the editor's internal binary format — with **no
DocumentServer**:

- **Open:** `docx/xlsx/pptx` → `Editor.bin` (`DOCY`/`XLSY`/`PPTY`) → fed to the SDK.
- **Save:** `Editor.bin` (from `asc_nativeGetFile`) → `docx/xlsx/pptx`.

It's driven by `../x2t-bridge.js` (`window.X2TBridge`).

## Where it came from

| | |
|---|---|
| **Upstream** | **CryptPad** — [`cryptpad/onlyoffice-x2t-wasm`](https://github.com/cryptpad/onlyoffice-x2t-wasm) (a WASM packaging of ONLYOFFICE's x2t/core). |
| **Pinned at** | [`upstream-pins/x2t.json`](../../../upstream-pins/x2t.json) — release `url` + `sha512_url`. |
| **License** | **AGPL-3.0** (same as ONLYOFFICE core). See [`../../../THIRD-PARTY.md`](../../../THIRD-PARTY.md). |

## Files

| File | What |
|---|---|
| `x2t.wasm` | the converter (~34 MB) |
| `x2t.js` | Emscripten JS loader/glue for the WASM |
| `*.br` | pre-brotli'd companions for HTTP content-negotiation |

## Why it's committed (not downloaded)

We commit the artifacts so builds are **offline / self-contained**. The
canonical source is still the pinned CryptPad release — `01-pull-upstream.sh`
will fetch it from `x2t.json` if the WASM is ever missing. To **bump** x2t,
update `upstream-pins/x2t.json` (match the major to the sdkjs version) and
re-vendor; verify against the published `sha512`.

> AGPL note: because we redistribute (serve) this WASM, the deployed site must
> offer its corresponding source — link the pinned CryptPad release. See
> `THIRD-PARTY.md` / the §13 source link.
