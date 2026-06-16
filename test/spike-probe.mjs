/**
 * spike-probe.mjs — exploratory run (not a test). Proves headless x2t in Node
 * and shows what the round-trip diff looks like before we tune normalization.
 *
 *   node test/spike-probe.mjs
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { convert, roundTrip } from './support/x2t-node.mjs';
import { unzip, normalizeOoxml, diffOoxml } from './support/ooxml-normalize.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIX = (rel) => path.join(HERE, 'fixtures', rel);

const sig = (u8) => String.fromCharCode(u8[0], u8[1], u8[2], u8[3]);

const CASES = [
  { ext: 'docx', file: FIX('docx/minimal.docx') },
  { ext: 'xlsx', file: FIX('xlsx/minimal.xlsx') },
  { ext: 'pptx', file: FIX('pptx/minimal.pptx') },
];

for (const { ext, file } of CASES) {
  console.log(`\n=== ${ext.toUpperCase()}: ${path.basename(file)} ===`);
  const src = new Uint8Array(readFileSync(file));
  console.log(`  source: ${src.length} bytes, ${unzip(src).size} zip entries`);

  const t0 = performance.now();
  const { bin, back } = await roundTrip(src, ext);
  const t1 = performance.now();
  console.log(`  → bin: ${bin.length} bytes, signature "${sig(bin)}"`);
  console.log(`  → back: ${back.length} bytes, ${unzip(back).size} zip entries`);
  console.log(`  round-trip time: ${(t1 - t0).toFixed(0)}ms`);

  const srcParts = Object.keys(normalizeOoxml(src)).sort();
  const backParts = Object.keys(normalizeOoxml(back)).sort();
  console.log(`  parts: src=${srcParts.length} back=${backParts.length}`);

  const diffs = diffOoxml(src, back);
  console.log(`  normalized diffs: ${diffs.length}`);
  for (const d of diffs.slice(0, 8)) console.log('    ' + d.replace(/\n/g, '\n    '));
  if (diffs.length > 8) console.log(`    … and ${diffs.length - 8} more`);
}

console.log('\nDONE.');
