/**
 * roundtrip.test.mjs — Tier B stability suite (headless x2t, no browser).
 *
 * Encodes what the spike validated about x2t's behaviour:
 *  - OOXML → bin → OOXML works for all three formats, producing valid output.
 *  - x2t NORMALIZES, so source≠round-trip is EXPECTED — we golden-compare the
 *    round-trip output, not assert identity with the source.
 *  - DOCX/PPTX: bin is deterministic and the round-trip is a FIXED POINT
 *    (re-saving is stable) → strong equality assertions.
 *  - XLSX: bin is NON-deterministic (cfRule GUIDs regenerated; an x2t comment-
 *    author-duplication defect) → we assert conversion validity + golden only,
 *    and document the instability rather than asserting a fixed point.
 *
 * Regenerate goldens:  UPDATE_GOLDEN=1 npx vitest run
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { convert, roundTrip } from '../support/x2t-node.mjs';
import { unzip, normalizeOoxml, diffOoxml } from '../support/ooxml-normalize.mjs';
import { matchGolden } from '../support/golden.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIX = (rel) => path.join(HERE, '..', 'fixtures', rel);
const read = (rel) => new Uint8Array(readFileSync(FIX(rel)));
const sig = (u8) => String.fromCharCode(u8[0], u8[1], u8[2], u8[3]);
const bytesEqual = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);

const FORMATS = [
  { ext: 'docx', binSig: 'DOCY', file: 'docx/minimal.docx', deterministic: true },
  { ext: 'xlsx', binSig: 'XLSY', file: 'xlsx/minimal.xlsx', deterministic: false },
  { ext: 'pptx', binSig: 'PPTY', file: 'pptx/minimal.pptx', deterministic: true },
];

describe('x2t headless round-trip (Tier B)', () => {
  for (const fmt of FORMATS) {
    describe(fmt.ext, () => {
      it('OOXML → Editor.bin produces the right signature', async () => {
        const bin = await convert(read(fmt.file), { from: fmt.ext, to: 'bin' });
        expect(bin.length).toBeGreaterThan(0);
        expect(sig(bin)).toBe(fmt.binSig);
      });

      it('Editor.bin → OOXML produces a valid, non-trivial ZIP', async () => {
        const { back } = await roundTrip(read(fmt.file), fmt.ext);
        const entries = unzip(back);                 // throws if not a valid ZIP
        expect(entries.size).toBeGreaterThan(3);
        expect(entries.has('[Content_Types].xml')).toBe(true);
      });

      it('round-trip output matches the committed golden', async () => {
        const { back } = await roundTrip(read(fmt.file), fmt.ext);
        const res = matchGolden(`${fmt.ext}/minimal.parts.json`, normalizeOoxml(back));
        if (!res.ok && !res.written) {
          // surface a readable hint on failure
          console.error(`golden mismatch for ${fmt.ext}; run UPDATE_GOLDEN=1 to refresh after reviewing`);
        }
        expect(res.ok).toBe(true);
      });

      if (fmt.deterministic) {
        it('bin is byte-deterministic (same input → same bytes)', async () => {
          const a = await convert(read(fmt.file), { from: fmt.ext, to: 'bin' });
          const b = await convert(read(fmt.file), { from: fmt.ext, to: 'bin' });
          expect(bytesEqual(a, b)).toBe(true);
        });

        it('round-trip is a fixed point (re-save is stable)', async () => {
          const { back: back1 } = await roundTrip(read(fmt.file), fmt.ext);
          const { back: back2 } = await roundTrip(back1, fmt.ext);
          expect(diffOoxml(back1, back2)).toEqual([]);
        });
      } else {
        // XLSX: characterize the known instability instead of failing on it.
        it('is documented non-deterministic (≤2 residual diffs on re-save)', async () => {
          const { back: back1 } = await roundTrip(read(fmt.file), fmt.ext);
          const { back: back2 } = await roundTrip(back1, fmt.ext);
          const diffs = diffOoxml(back1, back2);
          // Spike finding: 1 residual diff = x2t duplicating a comment author
          // prefix on round-trip. Pin a tight upper bound so NEW instability
          // (regressions) still trips the test.
          expect(diffs.length).toBeLessThanOrEqual(2);
        });
      }
    });
  }
});
