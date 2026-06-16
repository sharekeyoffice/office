/**
 * x2t-node.mjs — load the x2t WASM converter under Node (headless).
 *
 * Mirrors wrapper/v1/x2t-bridge.js's runConversion(), but for Node instead of
 * the browser. The browser bridge sets a global `Module` then injects x2t.js
 * via a <script> tag; under CommonJS `require` that doesn't work because
 * x2t.js's top-level `var Module` is module-scoped and shadows our global.
 *
 * Trick: load x2t.js through `new Function('Module', ..., code)` and pass our
 * pre-built Module object as the `Module` parameter. x2t.js's `var Module;`
 * (no initializer) does NOT reset an existing same-named binding, and its
 * `if(!Module)` guard then sees our truthy object and uses it — exactly the
 * browser's "external script tag defines var Module" path. This lets us set
 * `onRuntimeInitialized` *before* the runtime boots, with no race.
 *
 * x2t.js's Node branch (`ENVIRONMENT_IS_NODE`) reads x2t.wasm via fs from its
 * own dir, so we pass the real x2t.js path as __filename/__dirname.
 */
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const require = createRequire(import.meta.url);
const HERE = path.dirname(fileURLToPath(import.meta.url));
// test/support → repo root → wrapper/v1/x2t
const X2T_DIR = path.resolve(HERE, '..', '..', 'wrapper', 'v1', 'x2t');
const X2T_JS = path.join(X2T_DIR, 'x2t.js');

let modulePromise = null;

/** Load (once) and return the initialized Emscripten Module. */
export function loadX2T() {
  if (modulePromise) return modulePromise;

  modulePromise = new Promise((resolve, reject) => {
    let code;
    try {
      code = readFileSync(X2T_JS, 'utf8');
    } catch (e) {
      reject(new Error(`Cannot read x2t.js at ${X2T_JS}: ${e.message}`));
      return;
    }

    const Module = {
      noInitialRun: true,
      noExitRuntime: true,
      print: () => {},
      printErr: () => {},
      // pre-js.js overwrites locateFile for Node (prefix = x2t.js dir), so this
      // is belt-and-suspenders; either way x2t.wasm resolves next to x2t.js.
      locateFile: (p) => path.join(X2T_DIR, p),
      onAbort: (reason) => reject(new Error('x2t WASM aborted: ' + reason)),
      onRuntimeInitialized: () => {
        // Pre-create the working dirs the bridge uses.
        for (const d of ['/working', '/working/themes', '/working/media', '/working/fonts']) {
          try { Module.FS.mkdir(d); } catch { /* exists */ }
        }
        resolve(Module);
      },
    };

    const fakeModule = { exports: {} };
    try {
      // Params x2t.js references as free identifiers in its Node branch.
      const factory = new Function(
        'Module', 'require', '__filename', '__dirname', 'module', 'exports', 'process',
        code,
      );
      factory(Module, require, X2T_JS, X2T_DIR, fakeModule, fakeModule.exports, process);
    } catch (e) {
      reject(new Error('Failed to evaluate x2t.js: ' + e.message));
    }
  });

  return modulePromise;
}

let convCounter = 0;

/**
 * Convert bytes between an OOXML extension and the editor `bin` format.
 *
 *   convert(docxBytes, { from: 'docx', to: 'bin'  })  → Editor.bin (DOCY/…)
 *   convert(binBytes,  { from: 'bin',  to: 'docx' })  → OOXML zip bytes
 *
 * @param {Uint8Array} inputBytes
 * @param {{from: string, to: string}} opts
 * @returns {Promise<Uint8Array>}
 */
export async function convert(inputBytes, { from, to }) {
  const Module = await loadX2T();
  const id = ++convCounter;
  const inputPath = `/working/in_${id}.${from}`;
  const outputPath = `/working/out_${id}.${to}`;
  const paramsPath = `/working/params_${id}.xml`;

  const cleanup = () => {
    for (const p of [inputPath, outputPath, paramsPath]) {
      try { Module.FS.unlink(p); } catch { /* ignore */ }
    }
  };

  Module.FS.writeFile(inputPath, inputBytes);

  // Same params XML the browser bridge writes.
  const params =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<TaskQueueDataConvert ' +
      'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
      'xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
      `<m_sFileFrom>${inputPath}</m_sFileFrom>` +
      '<m_sThemeDir>/working/themes</m_sThemeDir>' +
      `<m_sFileTo>${outputPath}</m_sFileTo>` +
      '<m_bIsNoBase64>false</m_bIsNoBase64>' +
    '</TaskQueueDataConvert>';
  Module.FS.writeFile(paramsPath, params);

  let exitCode;
  try {
    exitCode = Module.ccall('main1', 'number', ['string'], [paramsPath]);
  } catch (e) {
    cleanup();
    throw new Error('x2t threw during main1: ' + e.message);
  }

  if (exitCode !== 0) {
    cleanup();
    const err = new Error(`x2t conversion failed (exit code ${exitCode})`);
    err.exitCode = exitCode;
    throw err;
  }

  let out;
  try {
    out = Module.FS.readFile(outputPath);
  } catch (e) {
    cleanup();
    throw new Error('Conversion succeeded but output is missing: ' + e.message);
  }

  const copy = new Uint8Array(out.length);
  copy.set(out);
  cleanup();
  return copy;
}

/** Convenience: OOXML → Editor.bin → OOXML (same target ext). */
export async function roundTrip(ooxmlBytes, ext) {
  const bin = await convert(ooxmlBytes, { from: ext, to: 'bin' });
  const back = await convert(bin, { from: 'bin', to: ext });
  return { bin, back };
}
