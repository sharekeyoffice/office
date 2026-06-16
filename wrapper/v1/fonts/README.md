# Bundled Fonts

This directory contains font files used by the offline viewer. Fonts are
**not committed** to the repository (gitignored) — download them per the
instructions below.

## Why these fonts

We need real `.ttf` files so that:

1. **sdkjs** has glyph metrics for layout (without them, text renders at
   zero width and is invisible).
2. **x2t.wasm** can resolve font references during OOXML conversion.

Phase 2.1 bundled the Liberation family (Arial / Times / Courier substitutes).
Phase 2.2 added Carlito (Calibri), Caladea (Cambria), DejaVu Sans (broad Unicode),
and OpenSymbol (Symbol/Wingdings/Webdings).

All fonts are licensed compatibly with our AGPL-3 distribution.

Total payload: **~10 MB** for 25 files.

## Download

All commands run from this directory:

```bash
cd wrapper/v1/fonts
```

### Liberation Sans / Serif / Mono — Phase 2.1 (12 files, ~4.4 MB)

OFL-1.1, metric-compatible with Arial / Times New Roman / Courier New.

```bash
curl -L -o /tmp/liberation-fonts.tar.gz \
  https://github.com/liberationfonts/liberation-fonts/files/7261482/liberation-fonts-ttf-2.1.5.tar.gz
tar -xzf /tmp/liberation-fonts.tar.gz --strip-components=1 \
    'liberation-fonts-ttf-2.1.5/LiberationSans-Regular.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationSans-Bold.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationSans-Italic.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationSans-BoldItalic.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationSerif-Regular.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationSerif-Bold.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationSerif-Italic.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationSerif-BoldItalic.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationMono-Regular.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationMono-Bold.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationMono-Italic.ttf' \
    'liberation-fonts-ttf-2.1.5/LiberationMono-BoldItalic.ttf'
rm /tmp/liberation-fonts.tar.gz
```

### Carlito — Phase 2.2 Calibri-substitute (4 files, ~2.7 MB)

Apache-2.0 / OFL-1.1 (Google Fonts). Metric-compatible with Calibri.

```bash
curl -sL -o /tmp/carlito.zip https://github.com/googlefonts/carlito/archive/refs/heads/main.zip
unzip -j -q /tmp/carlito.zip \
    'carlito-main/fonts/ttf/Carlito-Regular.ttf' \
    'carlito-main/fonts/ttf/Carlito-Italic.ttf' \
    'carlito-main/fonts/ttf/Carlito-Bold.ttf' \
    'carlito-main/fonts/ttf/Carlito-BoldItalic.ttf' -d .
rm /tmp/carlito.zip
```

### Caladea — Phase 2.2 Cambria-substitute (4 files, ~0.3 MB)

OFL-1.1 (Huerta Tipográfica). Metric-compatible with Cambria.

```bash
curl -sL -o /tmp/caladea.zip https://github.com/huertatipografica/Caladea/archive/refs/heads/master.zip
unzip -j -q /tmp/caladea.zip \
    'Caladea-master/fonts/ttf/Caladea-Regular.ttf' \
    'Caladea-master/fonts/ttf/Caladea-Italic.ttf' \
    'Caladea-master/fonts/ttf/Caladea-Bold.ttf' \
    'Caladea-master/fonts/ttf/Caladea-BoldItalic.ttf' -d .
rm /tmp/caladea.zip
```

### DejaVu Sans — Phase 2.2 broad Unicode (4 files, ~2.7 MB)

DejaVu license (free, GPL-compatible). Covers Cyrillic / Greek / Vietnamese /
extended Latin — used as the default fallback for any unrecognised font.

```bash
curl -sL -o /tmp/dejavu.tar.bz2 \
  https://github.com/dejavu-fonts/dejavu-fonts/releases/download/version_2_37/dejavu-fonts-ttf-2.37.tar.bz2
tar -xjf /tmp/dejavu.tar.bz2 --strip-components=2 \
    dejavu-fonts-ttf-2.37/ttf/DejaVuSans.ttf \
    dejavu-fonts-ttf-2.37/ttf/DejaVuSans-Oblique.ttf \
    dejavu-fonts-ttf-2.37/ttf/DejaVuSans-Bold.ttf \
    dejavu-fonts-ttf-2.37/ttf/DejaVuSans-BoldOblique.ttf
rm /tmp/dejavu.tar.bz2
```

### OpenSymbol — Phase 2.2 Symbol/Wingdings/Webdings (1 file, ~0.2 MB)

LGPL-3 (LibreOffice). Provides the dingbat / symbol glyphs that Microsoft's
Symbol & Wingdings cover. Not available as a standalone download — extract
from a LibreOffice installation:

```bash
# macOS
cp /Applications/LibreOffice.app/Contents/Resources/fonts/truetype/opens___.ttf OpenSymbol.ttf

# Linux (Debian/Ubuntu)
cp /usr/share/fonts/truetype/openoffice/opens___.ttf OpenSymbol.ttf
# or: apt install fonts-opensymbol; then locate opens___.ttf
```

If you don't have LibreOffice installed, download it once from
https://www.libreoffice.org/download/, copy the file out, and uninstall.

### Verify

```bash
ls -la *.ttf | wc -l  # should be 25
du -sh .              # should be ~10 MB
```

## Family-to-file mapping

The manifest at [common/AllFonts.js](../../common/AllFonts.js) maps Microsoft
font family names (referenced by `.docx` files) to these substitutes:

| Document references | Loaded file family |
|---|---|
| Arial / Helvetica          | LiberationSans |
| Times New Roman / Times    | LiberationSerif |
| Courier New / Courier      | LiberationMono |
| Calibri / Calibri Light    | Carlito |
| Cambria / Cambria Math     | Caladea |
| Symbol / Wingdings / Webdings | OpenSymbol |
| (anything else)            | DejaVu Sans (broad Unicode default) |

The "anything else → DejaVu Sans" fallback is implemented in [viewerPoc/index.html](../index.html)'s
`primeFontPicker()` (FALLBACK_MAP + FALLBACK_DEFAULT). Specific name overrides
(Tahoma → DejaVu Sans, Verdana → DejaVu Sans, Georgia → Liberation Serif, etc.)
live in the same map.

## Important — these are PLAIN TTFs

OnlyOffice's sdkjs by default expects font files to be pre-XOR'd (first 32 bytes
mangled with a known GUID). We've patched sdkjs to skip the XOR step when
`window.__plainFonts === true` (see [analysis/mitigation/font-xor-skip.md](../../analysis/mitigation/font-xor-skip.md)).

**Do not** XOR these files. Keep them as the upstream distributions ship them.
