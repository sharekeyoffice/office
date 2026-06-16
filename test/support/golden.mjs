/**
 * golden.mjs — read/write golden snapshots for round-trip comparison.
 *
 * A golden is the normalized representation (see ooxml-normalize) of a file's
 * round-trip output, captured once and committed. Regenerate with:
 *
 *   UPDATE_GOLDEN=1 npx vitest run
 *
 * Always review the git diff of test/golden/ before committing — a changed
 * golden means our pipeline's output changed (upstream pin bump, bridge edit,
 * or a regression).
 */
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const GOLDEN_DIR = path.resolve(HERE, '..', 'golden');
const UPDATE = process.env.UPDATE_GOLDEN === '1';

/**
 * Compare a normalized object against the stored golden at `relPath`
 * (e.g. "docx/minimal.parts.json"). When UPDATE_GOLDEN=1, writes instead.
 * Returns { ok, missing, expected } — callers assert in the test.
 */
export function matchGolden(relPath, normalized) {
  const file = path.join(GOLDEN_DIR, relPath);
  const actual = JSON.stringify(normalized, Object.keys(normalized).sort(), 2);

  if (UPDATE || !existsSync(file)) {
    mkdirSync(path.dirname(file), { recursive: true });
    writeFileSync(file, actual + '\n');
    return { ok: true, written: true, missing: !existsSync(file) };
  }
  const expected = readFileSync(file, 'utf8').trimEnd();
  return { ok: actual === expected, written: false, expected, actual };
}
