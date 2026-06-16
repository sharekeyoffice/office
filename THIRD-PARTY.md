# Third-party components & attributions

This deployment redistributes the following third-party software and
assets. Each is the property of its respective authors under the license
shown. See [NOTICE](NOTICE) for the ONLYOFFICE attribution + statement of
modifications, and [`wrapper/REGENERATE.md`](wrapper/REGENERATE.md) for how
each artifact is produced/vendored.

> Verify the marked **(verify)** licenses against the exact upstream
> release you pin before shipping — they are best-effort here.

## 1. ONLYOFFICE Docs — sdkjs & web-apps

- **Copyright:** Ascensio System SIA
- **License:** GNU AGPL-3.0 (see [LICENSE](LICENSE)) **plus** the ONLYOFFICE
  additional terms (AGPL §7) introduced in ONLYOFFICE Docs 9.4.
- **Source:** `github.com/ONLYOFFICE/sdkjs`, `github.com/ONLYOFFICE/web-apps`
  at the refs pinned in [`upstream-pins/`](upstream-pins/).
- **Corresponding source (AGPL §13):** the served editor includes a link to
  the source — the pinned upstream refs + this repo (`wrapper/` +
  `wrapper/sdkjs-patches/` + `build/`). Reproduce via `wrapper/REGENERATE.md`.
- **Trademark:** "ONLYOFFICE is a trademark of Ascensio System SIA." Used
  nominatively only.

## 2. x2t — OOXML ⇄ Editor.bin converter (WASM)

- **Project:** CryptPad's `onlyoffice-x2t-wasm` (a WASM build of ONLYOFFICE's
  x2t / Ascensio core conversion tooling).
- **License:** AGPL-3.0.
- **Source:** `github.com/cryptpad/onlyoffice-x2t-wasm` at the release pinned
  in [`upstream-pins/x2t.json`](upstream-pins/x2t.json) (`url` + `sha512_url`).
- **Files:** `wrapper/v1/x2t/{x2t.wasm,x2t.js}` (+ `.br`). Wrapped by our
  `wrapper/v1/x2t-bridge.js` (AGPL-3.0, original work).

## 3. Bundled fonts — `wrapper/v1/fonts/`

Substitution fonts for Microsoft families. **OFL-1.1 requires each font's
license text to be distributed with it** — ship `LICENSE-<family>.txt`
alongside the TTFs (see the gap noted in REGENERATE.md §G/H).

| Family (files) | Substitutes for | Source | License |
|---|---|---|---|
| Liberation Sans / Serif / Mono (12) | Arial / Times New Roman / Courier New | `liberationfonts/liberation-fonts` v2.1.5 | OFL-1.1 |
| Carlito (4) | Calibri | `googlefonts/carlito` | OFL-1.1 |
| Caladea (4) | Cambria | `googlefonts/caladea` | OFL-1.1 *(verify)* |
| DejaVu Sans (4) | broad Unicode coverage | `dejavu-fonts.github.io` | Bitstream Vera / DejaVu license (permissive) |
| OpenSymbol (1) | Symbol / Wingdings / Webdings | LibreOffice | MPL-2.0 / LGPL *(verify)* |

Download commands are preserved in `wrapper/v1/fonts/README.md`.

## 4. Icon sets

- **ONLYOFFICE editor UI icons** (in `web-apps`): licensed **CC BY-SA 4.0**
  (per the ONLYOFFICE License & Trademark guide — non-code assets such as
  illustrations and icon sets). **Attribution required; share-alike** — any
  modifications to these icons must be released under CC BY-SA 4.0.
  Not trademarked (the restricted items are the ONLYOFFICE *logo* and
  *trade dress*, not the functional icons).
- **Tab favicons** (`overlay/icons/{word,cell,slide}*.svg`): Sharekey
  original work.

## 5. Build tooling

- `esbuild` (MIT) — used by `wrapper/build/bundle.js`.
- `grunt` + plugins (MIT) — web-apps build.

---

### What is NOT redistributed
No ONLYOFFICE logo or brand trade dress is used as Sharekey branding. No
ONLYOFFICE DocumentServer, commercial components, or closed-source
`sdkjs-ooxml` addon is included.

### References
- ONLYOFFICE License & Trademark Policy:
  https://www.onlyoffice.com/blog/2026/05/onlyoffice-license-and-trademark-policy
- ONLYOFFICE Docs 9.4 release (license update):
  https://www.onlyoffice.com/blog/2026/05/onlyoffice-docs-9-4
