/**
 * x2t-bridge.js
 *
 * Thin wrapper around CryptPad's onlyoffice-x2t-wasm. Exposes a clean async API
 * for OOXML <-> Editor.bin conversion in the browser.
 *
 * License: AGPL-3.0 (matches sdkjs and onlyoffice-x2t-wasm)
 *
 * Usage:
 *   <script src="./x2t-bridge.js"></script>
 *   <script>
 *     X2TBridge.configure({ wasmDir: './x2t-test/x2t/' });
 *     X2TBridge.convertToBin(docxUint8Array, 'document.docx').then(binUint8Array => {
 *       // binUint8Array starts with DOCY/XLSY/PPTY signature — feed to sdkjs
 *     });
 *   </script>
 *
 * Design notes:
 * - Singleton WASM module (loaded once, reused across conversions).
 * - Each conversion uses unique virtual-FS paths to avoid collisions across
 *   concurrent or rapid-fire conversions.
 * - Files are unlinked from virtual FS after each conversion to bound memory.
 * - Errors include the exit code from main1 so callers can distinguish
 *   "conversion failed" from "WASM module crashed" from "I/O error".
 */
(function (global) {
	'use strict';

	// --------------------------------------------------------------
	// Config
	// --------------------------------------------------------------
	var config = {
		// Directory containing x2t.js and x2t.wasm. Must end in '/'.
		wasmDir: './x2t/',
		// Optional: custom path to x2t.js (overrides wasmDir + 'x2t.js')
		jsUrl: null,
		// Verbose logging (calls console.debug/warn/error)
		verbose: false
	};

	function log(level, msg) {
		if (!config.verbose && level === 'debug') return;
		var fn = console[level] || console.log;
		fn.call(console, '[x2t-bridge]', msg);
	}

	// --------------------------------------------------------------
	// WASM module loader (singleton, lazy)
	// --------------------------------------------------------------
	var modulePromise = null;

	function loadModule() {
		if (modulePromise) return modulePromise;

		modulePromise = new Promise(function (resolve, reject) {
			// Per cryptpad/onlyoffice-x2t-wasm pre-js.js convention:
			// we set Module options on window before x2t.js runs.
			global.Module = global.Module || {};
			global.Module.noInitialRun = true;
			global.Module.noExitRuntime = true;

			// locateFile: resolve x2t.wasm relative to wasmDir
			global.Module.locateFile = function (path /*, prefix */) {
				return config.wasmDir + path;
			};

			// Capture stdout/stderr from x2t for diagnostics
			var prevPrint    = global.Module.print;
			var prevPrintErr = global.Module.printErr;
			global.Module.print    = function (m) { log('debug', '[x2t stdout] ' + m); if (prevPrint) prevPrint(m); };
			global.Module.printErr = function (m) { log('warn',  '[x2t stderr] ' + m); if (prevPrintErr) prevPrintErr(m); };

			global.Module.onAbort = function (reason) {
				log('error', 'x2t aborted: ' + reason);
				reject(new Error('x2t WASM aborted: ' + reason));
			};

			global.Module.onRuntimeInitialized = function () {
				log('debug', 'x2t runtime initialized');

				// Pre-create working directories for the first conversion
				try {
					['/working', '/working/themes', '/working/media', '/working/fonts']
						.forEach(function (d) {
							try { global.Module.FS.mkdir(d); } catch (_) { /* exists */ }
						});
				} catch (e) {
					log('warn', 'Could not pre-create working dirs: ' + e.message);
				}

				resolve(global.Module);
			};

			// Inject script tag.
			// IMPORTANT: cryptpad's pre-js.js does `new URL(myScript.getAttribute("src")).search`
			// which only accepts absolute URLs. Resolve relative paths against the page first.
			var rawUrl = config.jsUrl || (config.wasmDir + 'x2t.js');
			var url    = new URL(rawUrl, global.location.href).href;
			var script = document.createElement('script');
			script.src = url;
			script.onerror = function () {
				reject(new Error('Failed to load ' + url + ' — check wasmDir config'));
			};
			document.head.appendChild(script);

			log('debug', 'Loading x2t.js from ' + url);
		});

		return modulePromise;
	}

	// --------------------------------------------------------------
	// Conversion primitive
	// --------------------------------------------------------------
	// Counter to generate unique paths so concurrent calls don't collide
	var _convCounter = 0;

	/**
	 * Run x2t conversion: writes input + params, calls main1, reads output.
	 * Cleans up virtual-FS files even on failure.
	 *
	 * @param {Uint8Array} inputBytes
	 * @param {string}     inputExt    e.g. "docx"
	 * @param {string}     outputExt   e.g. "bin"
	 * @returns {Promise<Uint8Array>}
	 */
	function runConversion(inputBytes, inputExt, outputExt) {
		return loadModule().then(function (Module) {
			var id = ++_convCounter;
			var inputPath  = '/working/in_'  + id + '.' + inputExt;
			var outputPath = '/working/out_' + id + '.' + outputExt;
			var paramsPath = '/working/params_' + id + '.xml';

			var cleanup = function () {
				[inputPath, outputPath, paramsPath].forEach(function (p) {
					try { Module.FS.unlink(p); } catch (_) { /* ignore */ }
				});
			};

			try {
				Module.FS.writeFile(inputPath, inputBytes);
			} catch (e) {
				return Promise.reject(new Error('Could not write input to virtual FS: ' + e.message));
			}

			var params =
				'<?xml version="1.0" encoding="utf-8"?>' +
				'<TaskQueueDataConvert ' +
					'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
					'xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
					'<m_sFileFrom>' + inputPath + '</m_sFileFrom>' +
					'<m_sThemeDir>/working/themes</m_sThemeDir>' +
					'<m_sFileTo>' + outputPath + '</m_sFileTo>' +
					'<m_bIsNoBase64>false</m_bIsNoBase64>' +
				'</TaskQueueDataConvert>';

			try {
				Module.FS.writeFile(paramsPath, params);
			} catch (e) {
				cleanup();
				return Promise.reject(new Error('Could not write params: ' + e.message));
			}

			var t0 = (typeof performance !== 'undefined') ? performance.now() : Date.now();
			var exitCode;
			try {
				exitCode = Module.ccall('main1', 'number', ['string'], [paramsPath]);
			} catch (e) {
				cleanup();
				return Promise.reject(new Error('x2t threw during main1: ' + e.message));
			}
			var t1 = (typeof performance !== 'undefined') ? performance.now() : Date.now();
			log('debug', 'main1 returned ' + exitCode + ' in ' + (t1 - t0).toFixed(0) + 'ms (id=' + id + ')');

			if (exitCode !== 0) {
				cleanup();
				var err = new Error('x2t conversion failed (exit code ' + exitCode + ')');
				err.exitCode = exitCode;
				return Promise.reject(err);
			}

			var outputBytes;
			try {
				outputBytes = Module.FS.readFile(outputPath);
			} catch (e) {
				cleanup();
				return Promise.reject(new Error('Conversion succeeded but output is missing: ' + e.message));
			}

			// Copy out of WASM heap so we can free virtual-FS files immediately
			var copy = new Uint8Array(outputBytes.length);
			copy.set(outputBytes);
			cleanup();
			return copy;
		});
	}

	// --------------------------------------------------------------
	// Public API
	// --------------------------------------------------------------

	/**
	 * Detect file format from filename extension. Returns lowercase extension
	 * without leading dot, or null if not recognized.
	 */
	function getExt(name) {
		if (!name) return null;
		var m = String(name).toLowerCase().match(/\.([a-z0-9]+)$/);
		return m ? m[1] : null;
	}

	/**
	 * Detect editor type from a Uint8Array. Returns one of:
	 *   "ooxml" — ZIP-based (.docx/.xlsx/.pptx and OOXML siblings)
	 *   "bin"   — already an Editor.bin (DOCY/XLSY/PPTY/VSDY signature)
	 *   "other" — anything else
	 */
	function detectFormat(bytes) {
		if (!bytes || bytes.length < 4) return 'other';
		// PK ZIP magic
		if (bytes[0] === 0x50 && bytes[1] === 0x4B && bytes[2] === 0x03 && bytes[3] === 0x04) {
			return 'ooxml';
		}
		// Editor.bin signatures
		var sig = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
		if (sig === 'DOCY' || sig === 'XLSY' || sig === 'PPTY' || sig === 'VSDY') {
			return 'bin';
		}
		return 'other';
	}

	/**
	 * Get the 4-character signature of an Editor.bin Uint8Array.
	 * Returns "DOCY" / "XLSY" / "PPTY" / "VSDY" or null.
	 */
	function getBinSignature(bytes) {
		if (!bytes || bytes.length < 4) return null;
		var sig = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
		if (sig === 'DOCY' || sig === 'XLSY' || sig === 'PPTY' || sig === 'VSDY') return sig;
		return null;
	}

	/**
	 * Read the version field from an Editor.bin header.
	 * Header format: "DOCY;v<N>;<size>;<binary>"
	 * Returns the integer N, or null if the header is malformed.
	 */
	function getBinVersion(bytes) {
		if (!bytes || bytes.length < 8) return null;
		// Read up to ~32 bytes as ASCII to find the header
		var max = Math.min(32, bytes.length);
		var head = '';
		for (var i = 0; i < max; i++) head += String.fromCharCode(bytes[i]);
		var m = head.match(/^[A-Z]{4};v(\d+);/);
		return m ? parseInt(m[1], 10) : null;
	}

	/**
	 * Patch the version field in an Editor.bin header.
	 *
	 * x2t v9.3.0 emits headers like "DOCY;v5;..." (legacy version field) but the
	 * payload is raw binary, not base64-encoded. sdkjs's reader (Serialize2.js:7853)
	 * uses the rule:
	 *
	 *     if (Asc.c_nVersionNoBase64 !== AscCommon.CurFileVersion) base64-decode
	 *     else                                                       raw-binary
	 *
	 * With Asc.c_nVersionNoBase64 = 10 and a v5 header, sdkjs will try to base64-
	 * decode the raw bytes and produce garbage. Rewriting the header to v10
	 * forces the raw-binary branch.
	 *
	 * @param {Uint8Array} bytes        Editor.bin with header
	 * @param {number}     targetVersion  desired version number (default 10)
	 * @returns {Uint8Array}  new buffer with patched header (or the same buffer if
	 *                        already at targetVersion or no header found)
	 */
	function patchBinVersion(bytes, targetVersion) {
		targetVersion = targetVersion || 10;
		var current = getBinVersion(bytes);
		if (current === null || current === targetVersion) return bytes;

		// Header layout: <4-byte signature>;v<digits>;<rest...>
		// Find the byte index of the second ';' (end of version field)
		var sigEnd  = 4;                        // position of ';' after signature
		var verEnd  = -1;
		for (var i = sigEnd + 2; i < bytes.length && i < 32; i++) {
			if (bytes[i] === 0x3b) { verEnd = i; break; }  // ';'
		}
		if (verEnd === -1) return bytes;

		// Build replacement: "<sig>;v<targetVersion>"
		var sig = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
		var newHead = sig + ';v' + targetVersion;
		var newHeadBytes = new Uint8Array(newHead.length);
		for (var j = 0; j < newHead.length; j++) newHeadBytes[j] = newHead.charCodeAt(j);

		// Splice: newHeadBytes + bytes[verEnd .. end]
		var rest = bytes.subarray(verEnd);
		var out = new Uint8Array(newHeadBytes.length + rest.length);
		out.set(newHeadBytes, 0);
		out.set(rest, newHeadBytes.length);
		log('debug', 'Patched bin header v' + current + ' -> v' + targetVersion);
		return out;
	}

	/**
	 * Convert OOXML (.docx/.xlsx/.pptx/.odt/.ods/.odp) to Editor.bin.
	 *
	 * @param {Uint8Array | ArrayBuffer} input
	 * @param {string} fileName  used to infer input format from extension
	 * @returns {Promise<Uint8Array>}  Editor.bin bytes (DOCY/XLSY/PPTY signature)
	 */
	// Last conversion's extracted media (path → Blob URL). Set by convertToBin.
	// Consumers (viewer) call X2TBridge.getLastMedia() and register entries with
	// AscCommon.g_oDocumentUrls before calling api.OpenDocumentFromBin.
	var _lastMedia = {};
	function getLastMedia() { return _lastMedia; }

	function extractMediaFromFs(Module) {
		// Free old blob URLs so we don't leak across conversions.
		for (var k in _lastMedia) URL.revokeObjectURL(_lastMedia[k]);
		_lastMedia = {};
		var dir = '/working/media';
		var entries;
		try { entries = Module.FS.readdir(dir); } catch (e) { return _lastMedia; }
		entries.forEach(function (name) {
			if (name === '.' || name === '..') return;
			var path = dir + '/' + name;
			try {
				var stat = Module.FS.stat(path);
				if ((stat.mode & 0o170000) !== 0o100000) return;  // not a regular file
				var data = Module.FS.readFile(path);
				// Guess MIME from extension
				var ext = (name.split('.').pop() || '').toLowerCase();
				var mime = {
					png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
					gif: 'image/gif', svg: 'image/svg+xml', bmp: 'image/bmp',
					webp: 'image/webp', tiff: 'image/tiff', ico: 'image/x-icon',
					emf: 'application/octet-stream', wmf: 'application/octet-stream'
				}[ext] || 'application/octet-stream';
				var copy = new Uint8Array(data.length); copy.set(data);
				_lastMedia[name] = URL.createObjectURL(new Blob([copy], { type: mime }));
			} catch (e) {
				log('warn', 'media extract failed for ' + name + ': ' + e.message);
			}
		});
		log('debug', 'extracted ' + Object.keys(_lastMedia).length + ' media files: ' + Object.keys(_lastMedia).join(', '));
		return _lastMedia;
	}

	function convertToBin(input, fileName) {
		var bytes = (input instanceof Uint8Array) ? input : new Uint8Array(input);
		var ext = getExt(fileName) || 'docx';
		log('debug', 'convertToBin: ' + fileName + ' (' + bytes.length + ' bytes, ext=' + ext + ')');
		// Run conversion, then extract media from /working/media/ before the FS
		// gets reused on the next conversion. Media survive as Blob URLs.
		return runConversion(bytes, ext, 'bin').then(function (out) {
			return loadModule().then(function (Module) {
				extractMediaFromFs(Module);
				return out;
			});
		});
		// NOTE: do NOT patch the header to v10. x2t v9.3.0 emits a header of the
		// form "DOCY;v5;<dst_len>;<base64-payload>" — i.e. the payload IS base64-
		// encoded, despite the m_bIsNoBase64=false param. With the default
		// Asc.c_nVersionNoBase64 = 10 and the v5 header, the reader at
		// Serialize2.js:7853 sees c_nVersionNoBase64 (10) !== CurFileVersion (5)
		// and correctly takes the base64-decode branch.
		//
		// Patching the header to v10 makes CurFileVersion (parsed from the header
		// at line 7850) become 10, which then matches c_nVersionNoBase64=10 and
		// forces the raw-binary branch — reading ASCII base64 as raw bytes
		// produces garbage that happens to deserialize as one empty paragraph
		// without erroring.
	}

	/**
	 * Convert Editor.bin back to a target OOXML/ODF format.
	 *
	 * @param {Uint8Array | ArrayBuffer} input         Editor.bin bytes
	 * @param {string} outputExt                       e.g. "docx", "xlsx", "pptx"
	 * @returns {Promise<Uint8Array>}                  ZIP/OOXML bytes
	 */
	function convertFromBin(input, outputExt) {
		var bytes = (input instanceof Uint8Array) ? input : new Uint8Array(input);
		log('debug', 'convertFromBin: ' + bytes.length + ' bytes -> .' + outputExt);
		return runConversion(bytes, 'bin', outputExt);
	}

	/**
	 * Pre-load fonts into x2t's virtual filesystem.
	 *
	 * x2t reads fonts from /working/fonts/<FontName><suffix>.ttf where suffix is
	 * "" / "_Bold" / "_Italic" / "_Bold_Italic". CryptPad's reference integration
	 * (www/common/outer/x2t.js fetchFonts) does this once per conversion.
	 *
	 * For our offline viewer x2t doesn't strictly need fonts to convert OOXML to
	 * Editor.bin (verified: 21.docx converts fine with no fonts). But some shape /
	 * metric calculations may use them, so this helper is available for cases that
	 * need it.
	 *
	 * @param {string}   fontDirUrl   e.g. "./fonts/" — same dir sdkjs reads
	 * @param {Array}    fontInfos    same shape as window["__fonts_infos"]
	 * @param {Array}    fontFiles    same shape as window["__fonts_files"]
	 * @returns {Promise<void>}
	 */
	function preloadFonts(fontDirUrl, fontInfos, fontFiles) {
		if (!fontDirUrl || !fontInfos || !fontFiles) {
			return Promise.reject(new Error('preloadFonts: missing arguments'));
		}
		// Suffix per style index (matches CryptPad convention)
		// __fonts_infos row layout (after name+thumb):
		//   indexR (idx 2), faceR (3), indexI (4), faceI (5),
		//   indexB (idx 6), faceB (7), indexBI (8), faceBI (9)
		var STYLES = [
			{ slot: 2, suffix: ''             },
			{ slot: 4, suffix: '_Italic'      },
			{ slot: 6, suffix: '_Bold'        },
			{ slot: 8, suffix: '_Bold_Italic' }
		];

		return loadModule().then(function (Module) {
			var seen = {};                  // dedupe by output path
			var jobs = [];

			fontInfos.forEach(function (info) {
				var name = info[0];
				STYLES.forEach(function (s) {
					var fileIdx = info[s.slot];
					if (typeof fileIdx !== 'number' || fileIdx < 0) return;
					var fileId = fontFiles[fileIdx];
					if (!fileId) return;
					var outName = name + s.suffix + '.ttf';
					var outPath = '/working/fonts/' + outName;
					if (seen[outPath]) return;
					seen[outPath] = true;

					var url = fontDirUrl + (fontDirUrl.endsWith('/') ? '' : '/') + fileId;
					jobs.push(
						fetch(url).then(function (r) {
							if (!r.ok) throw new Error('Font fetch failed: ' + url + ' (' + r.status + ')');
							return r.arrayBuffer();
						}).then(function (buf) {
							Module.FS.writeFile(outPath, new Uint8Array(buf));
							log('debug', 'preloadFonts: wrote ' + outPath + ' (' + buf.byteLength + ' bytes)');
						})
					);
				});
			});

			return Promise.all(jobs).then(function () {
				log('debug', 'preloadFonts: complete (' + jobs.length + ' files)');
			});
		});
	}

	/**
	 * Configure the bridge. Call before first conversion.
	 *   - wasmDir: directory containing x2t.js and x2t.wasm (default './x2t/')
	 *   - jsUrl:   override URL for x2t.js (rare)
	 *   - verbose: enable console.debug logging
	 */
	function configure(opts) {
		opts = opts || {};
		if (typeof opts.wasmDir === 'string') {
			config.wasmDir = opts.wasmDir.endsWith('/') ? opts.wasmDir : opts.wasmDir + '/';
		}
		if (typeof opts.jsUrl === 'string') config.jsUrl = opts.jsUrl;
		if (typeof opts.verbose === 'boolean') config.verbose = opts.verbose;
	}

	/**
	 * Pre-load the WASM module (e.g. on app boot, while user navigates UI).
	 * Returns a promise that resolves when x2t is ready.
	 */
	function preload() {
		return loadModule().then(function () { /* discard module ref */ });
	}

	/**
	 * Returns true if the WASM module has finished loading.
	 */
	function isReady() {
		// modulePromise being non-null means loading was started; check if resolved
		// We can't synchronously inspect a Promise, so we expose this via a side flag.
		return _isReady;
	}
	var _isReady = false;
	// Side-effect: when modulePromise resolves, flip the flag
	var _origLoadModule = loadModule;
	loadModule = function () {
		var p = _origLoadModule();
		p.then(function () { _isReady = true; });
		return p;
	};

	// --------------------------------------------------------------
	// Export
	// --------------------------------------------------------------
	var X2TBridge = {
		configure:       configure,
		preload:         preload,
		preloadFonts:    preloadFonts,
		isReady:         isReady,
		convertToBin:    convertToBin,
		convertFromBin:  convertFromBin,
		detectFormat:    detectFormat,
		getBinSignature: getBinSignature,
		getBinVersion:   getBinVersion,
		patchBinVersion: patchBinVersion,
		getExt:          getExt,
		getLastMedia:    getLastMedia
	};

	if (typeof module !== 'undefined' && module.exports) {
		module.exports = X2TBridge;
	} else {
		global.X2TBridge = X2TBridge;
	}
})(typeof window !== 'undefined' ? window : this);
