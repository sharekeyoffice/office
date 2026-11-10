/**
 * x2t-perturbed.mjs — run one pptx → bin conversion in a fresh x2t module
 * whose heap layout has been shifted first. Meant to be spawned as a child
 * process: one process, one module, one verdict.
 *
 *   node test/support/x2t-perturbed.mjs <file.pptx> <perturbation>
 *
 *   perturbation:
 *     dup<N>        load N extra copies of the bundled fonts (x2t reads them
 *                   into its heap, so allocations land at other addresses)
 *     warm:<a,b>    convert other files first in the same module, the way the
 *                   browser reuses one module for the whole session
 *
 * Prints one line of JSON: {"ok":true} or {"ok":false,"error":"..."}.
 */
import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadX2T } from './x2t-node.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FONTS = path.resolve(HERE, '..', '..', 'wrapper', 'v1', 'fonts');

const [file, perturbation = 'dup0'] = process.argv.slice(2);

const Module = await loadX2T();
const fonts = readdirSync(FONTS).filter((f) => /\.ttf$/i.test(f));
for (const f of fonts) {
  Module.FS.writeFile('/working/fonts/' + f, new Uint8Array(readFileSync(path.join(FONTS, f))));
}
if (perturbation.startsWith('dup')) {
  const n = Number(perturbation.slice(3));
  for (let i = 0; i < n; i++) {
    const src = path.join(FONTS, fonts[i % fonts.length]);
    Module.FS.writeFile(`/working/fonts/zz${i}.ttf`, new Uint8Array(readFileSync(src)));
  }
}

let id = 0;

function convert(bytes, ext) {
  id++;
  Module.FS.writeFile(`/working/in${id}.${ext}`, bytes);
  Module.FS.writeFile(
    `/working/p${id}.xml`,
    '<?xml version="1.0" encoding="utf-8"?>' +
      '<TaskQueueDataConvert xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
      'xmlns:xsd="http://www.w3.org/2001/XMLSchema">' +
      `<m_sFileFrom>/working/in${id}.${ext}</m_sFileFrom>` +
      '<m_sThemeDir>/working/themes</m_sThemeDir>' +
      `<m_sFileTo>/working/out${id}.bin</m_sFileTo>` +
      '<m_sFontDir>/working/fonts</m_sFontDir>' +
      '<m_bIsNoBase64>false</m_bIsNoBase64>' +
      '</TaskQueueDataConvert>',
  );
  const rc = Module.ccall('main1', 'number', ['string'], [`/working/p${id}.xml`]);
  if (rc !== 0) throw new Error('exit code ' + rc);
  return Module.FS.readFile(`/working/out${id}.bin`);
}

try {
  if (perturbation.startsWith('warm:')) {
    for (const w of perturbation.slice(5).split(',')) {
      convert(new Uint8Array(readFileSync(w)), path.extname(w).slice(1));
    }
  }
  const bin = convert(new Uint8Array(readFileSync(file)), 'pptx');
  const sig = String.fromCharCode(bin[0], bin[1], bin[2], bin[3]);
  console.log(JSON.stringify({ ok: sig === 'PPTY', sig }));
} catch (e) {
  console.log(JSON.stringify({ ok: false, error: e.message }));
}
process.exit(0);
