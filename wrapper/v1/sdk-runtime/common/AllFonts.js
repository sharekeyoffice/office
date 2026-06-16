// AllFonts.js — Phase 2 manifest for the viewerPoc.
//
// In an OnlyOffice production deployment this file is GENERATED at build time
// from the server's font directory. For our offline PoC we hand-write a small
// manifest that points at the Liberation fonts in viewerPoc/fonts/.
//
// Manifest shape (verified in common/Drawings/Externals.js:664-676):
//
//   window["__fonts_files"] : Array<string>
//     One entry per font file. Strings are the file IDs that get appended to
//     fontFilesPath to form the URL.
//
//   window["__fonts_infos"] : Array<Array>
//     One entry per font FAMILY. Each row is 9 elements (thumbnail comes from
//     the family's POSITION in this array, not from the row):
//
//     [name, indexR, faceIdxR, indexI, faceIdxI, indexB, faceIdxB, indexBI, faceIdxBI]
//        0     1        2         3       4         5       6         7        8
//
//     Use index = -1 if a style is not available (synthesized at runtime).
//     Use faceIdx = 0 unless reading a TrueType Collection (.ttc).
//
//   window["g_fonts_selection_bin"] : string
//     Optional precompiled per-character font-coverage data. Empty string
//     skips the parsing path; sdkjs falls back to runtime detection.
//
// We map common Microsoft families to metric-compatible Liberation files:
//   Arial / Helvetica  -> LiberationSans
//   Times New Roman    -> LiberationSerif
//   Courier New        -> LiberationMono

(function () {
	// ---- File index assignments -----------------------------------------------
	// Order matters: __fonts_infos references files by their array index.
	var FILES = [
		"LiberationSans-Regular.ttf",      // 0
		"LiberationSans-Italic.ttf",       // 1
		"LiberationSans-Bold.ttf",         // 2
		"LiberationSans-BoldItalic.ttf",   // 3
		"LiberationSerif-Regular.ttf",     // 4
		"LiberationSerif-Italic.ttf",      // 5
		"LiberationSerif-Bold.ttf",        // 6
		"LiberationSerif-BoldItalic.ttf",  // 7
		"LiberationMono-Regular.ttf",      // 8
		"LiberationMono-Italic.ttf",       // 9
		"LiberationMono-Bold.ttf",         // 10
		"LiberationMono-BoldItalic.ttf",   // 11
		// Phase 2.2 additions
		"Carlito-Regular.ttf",             // 12 — Calibri-substitute
		"Carlito-Italic.ttf",              // 13
		"Carlito-Bold.ttf",                // 14
		"Carlito-BoldItalic.ttf",          // 15
		"Caladea-Regular.ttf",             // 16 — Cambria-substitute
		"Caladea-Italic.ttf",              // 17
		"Caladea-Bold.ttf",                // 18
		"Caladea-BoldItalic.ttf",          // 19
		"DejaVuSans.ttf",                  // 20 — broad Unicode coverage
		"DejaVuSans-Oblique.ttf",          // 21
		"DejaVuSans-Bold.ttf",             // 22
		"DejaVuSans-BoldOblique.ttf",      // 23
		"OpenSymbol.ttf"                   // 24 — Symbol/Wingdings/Webdings
	];

	// ---- Family entries -------------------------------------------------------
	// 9 elements per row: [name, R, fR, I, fI, B, fB, BI, fBI]
	// Multiple names can point at the same files (e.g. "Arial" and "Helvetica"
	// both map to LiberationSans). Use -1 for unavailable styles — sdkjs
	// synthesizes italic/bold from the regular face.
	var INFOS = [
		// name,             R,  fR,  I,  fI,  B,  fB, BI, fBI
		[ "Arial",            0,  0,   1,  0,   2,  0,   3,  0 ],
		[ "Helvetica",        0,  0,   1,  0,   2,  0,   3,  0 ],
		[ "Liberation Sans",  0,  0,   1,  0,   2,  0,   3,  0 ],

		[ "Times New Roman",  4,  0,   5,  0,   6,  0,   7,  0 ],
		[ "Times",            4,  0,   5,  0,   6,  0,   7,  0 ],
		[ "Liberation Serif", 4,  0,   5,  0,   6,  0,   7,  0 ],

		[ "Courier New",      8,  0,   9,  0,  10,  0,  11,  0 ],
		[ "Courier",          8,  0,   9,  0,  10,  0,  11,  0 ],
		[ "Liberation Mono",  8,  0,   9,  0,  10,  0,  11,  0 ],

		// Phase 2.2 — additional families
		[ "Calibri",         12,  0,  13,  0,  14,  0,  15,  0 ],
		[ "Calibri Light",   12,  0,  13,  0,  14,  0,  15,  0 ],
		[ "Carlito",         12,  0,  13,  0,  14,  0,  15,  0 ],

		[ "Cambria",         16,  0,  17,  0,  18,  0,  19,  0 ],
		[ "Cambria Math",    16,  0,  17,  0,  18,  0,  19,  0 ],
		[ "Caladea",         16,  0,  17,  0,  18,  0,  19,  0 ],

		[ "DejaVu Sans",     20,  0,  21,  0,  22,  0,  23,  0 ],

		// Symbol fonts: only one face exists; sdkjs synthesizes I/B/BI on demand.
		[ "Symbol",          24,  0,  -1,  0,  -1,  0,  -1,  0 ],
		[ "Wingdings",       24,  0,  -1,  0,  -1,  0,  -1,  0 ],
		[ "Webdings",        24,  0,  -1,  0,  -1,  0,  -1,  0 ],
		[ "OpenSymbol",      24,  0,  -1,  0,  -1,  0,  -1,  0 ]
	];

	window["__fonts_files"] = FILES;
	window["__fonts_infos"] = INFOS;

	// CFontSelectList.Init checks `g_fonts_selection_bin != ""` before decoding.
	// undefined != "" is true → would crash. Empty string takes the no-op branch.
	window["g_fonts_selection_bin"] = "";
})();
