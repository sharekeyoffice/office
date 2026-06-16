# Bundled fonts — copyrights, licenses & provenance

These TTFs are redistributed (committed here and served at `/fonts/*.ttf`).
Each family's license **requires its copyright notice + license text to
accompany the fonts** — that is what this directory satisfies. The fonts are
metric-compatible substitutes for Microsoft families; no Microsoft fonts are
distributed.

| Family (files) | Substitutes for | Copyright / Reserved Name | License | License file |
|---|---|---|---|---|
| **Liberation** Sans/Serif/Mono (12) | Arial / Times New Roman / Courier New | Copyright © 2007–2012 Red Hat, Inc. Reserved Font Name "Liberation". | OFL-1.1 | [OFL-1.1.txt](OFL-1.1.txt) |
| **Carlito** (4) | Calibri | Copyright 2013 The Carlito Project Authors (https://github.com/googlefonts/carlito). Reserved Font Name "Carlito". | OFL-1.1 | [OFL-1.1.txt](OFL-1.1.txt) |
| **Caladea** (4) | Cambria | Copyright © The Caladea Project Authors (Huerta Tipográfica). Reserved Font Name "Caladea". | OFL-1.1 | [OFL-1.1.txt](OFL-1.1.txt) |
| **DejaVu Sans** (4) | broad Unicode | © Bitstream (Vera); Arev glyphs © Tavmjong Bah; DejaVu changes public domain. | Bitstream Vera / public-domain | [DejaVu-LICENSE.txt](DejaVu-LICENSE.txt) |
| **OpenSymbol** (1) | Symbol / Wingdings / Webdings | LibreOffice / The Document Foundation | ⚠️ **verify** (LibreOffice ships it under MPL-2.0 / LGPL-3.0+) | _TODO: add `OpenSymbol-LICENSE.txt`_ |

## Provenance (download sources)

- **Liberation** — `github.com/liberationfonts/liberation-fonts` v2.1.5 (OFL-1.1).
- **Carlito** — `github.com/googlefonts/carlito` (OFL-1.1).
- **Caladea** — Google Fonts / `github.com/huertatipografica/Caladea` (OFL-1.1).
- **DejaVu** — `dejavu-fonts.github.io`.
- **OpenSymbol** — ships with LibreOffice (`github.com/LibreOffice/core`,
  `extras/source/truetype/symbol/`).

The exact download/extract commands are preserved in this directory's
`README.md`.

## ⚠️ Open item
`OpenSymbol.ttf` is bundled without its license file. Before shipping, add
`OpenSymbol-LICENSE.txt` with the LibreOffice/TDF notice for that face and
update the table above. (Flagged in `wrapper/REGENERATE.md` §G.)
