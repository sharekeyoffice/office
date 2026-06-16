/**
 * spike-fixedpoint.mjs — does x2t reach a fixed point?
 *
 * Source→bin→OOXML is NOT identity (x2t normalizes). The question that decides
 * the whole stability-test strategy: is the SECOND pass stable? i.e. does
 *   back1 = rt(src);  back2 = rt(back1)   ⇒   back1 ≈ back2 ?
 * If yes, golden-compare + re-save idempotency are valid assertions.
 * Also checks bin determinism: src→bin twice ⇒ identical bytes?
 *
 *   node test/spike-fixedpoint.mjs
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { convert, roundTrip } from './support/x2t-node.mjs';
import { diffOoxml } from './support/ooxml-normalize.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIX = (rel) => path.join(HERE, 'fixtures', rel);
const eq = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);

const CASES = [
  { ext: 'docx', file: FIX('docx/minimal.docx') },
  { ext: 'xlsx', file: FIX('xlsx/minimal.xlsx') },
  { ext: 'pptx', file: FIX('pptx/minimal.pptx') },
];

for (const { ext, file } of CASES) {
  console.log(`\n=== ${ext.toUpperCase()} ===`);
  const src = new Uint8Array(readFileSync(file));

  // bin determinism: same input → same bin bytes?
  const binA = await convert(src, { from: ext, to: 'bin' });
  const binB = await convert(src, { from: ext, to: 'bin' });
  console.log(`  bin determinism (src→bin twice): ${eq(binA, binB) ? 'IDENTICAL ✓' : 'DIFFERS ✗'}`);

  // fixed point: rt(src) vs rt(rt(src))
  const { back: back1 } = await roundTrip(src, ext);
  const { back: back2 } = await roundTrip(back1, ext);
  const srcVsBack1 = diffOoxml(src, back1).length;
  const back1VsBack2 = diffOoxml(back1, back2).length;
  console.log(`  normalized diffs src   → back1 : ${srcVsBack1}`);
  console.log(`  normalized diffs back1 → back2 : ${back1VsBack2}  ${back1VsBack2 === 0 ? 'FIXED POINT ✓' : '(still moving)'}`);

  // OOXML byte stability of the second pass (after normalization churn settles)
  const back1b = (await roundTrip(src, ext)).back;
  console.log(`  back determinism (rt(src) twice, raw bytes): ${eq(back1, back1b) ? 'IDENTICAL ✓' : 'differ (zip/rezip noise)'}`);
}

console.log('\nDONE.');
