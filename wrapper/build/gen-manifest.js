#!/usr/bin/env node
/**
 * gen-manifest.js — derive wrapper/scripts-{editor}.js from upstream
 * sdkjs/configs/{editor}.json. This makes the bundle manifest a documented,
 * reproducible function of upstream instead of a hand-copied file.
 *
 * The transform (reverse-engineered + verified byte-exact 2026-06-02):
 *
 *   manifest = PREPEND ++ configs.sdk.min ++ configs.sdk.common
 *              with per-editor EXTRA_INSERTS applied.
 *
 *   PREPEND (all editors): vendor/polyfill.js, common/AllFonts.js,
 *                          common/applyDocumentChanges.js
 *   EXTRA_INSERTS:
 *     cell, slide: insert common/zlib/zlib.js right AFTER word/apiCommon.js
 *                  (word already lists zlib in configs.sdk.min, so no insert)
 *
 * The `configs.sdk.desktop` section (which contains the closed/Local files
 * common/Local/{license,common}.js and <editor>/Local/api.js) is intentionally
 * NOT included — those are the DocumentServer/licensing bits we run without.
 *
 * Usage:
 *   node gen-manifest.js --sdk-dir <sdkjs> [--editor word|cell|slide] [--write] [--verify]
 *
 *   (default: --verify, no write — prints a diff vs the committed manifest)
 *   --write   overwrite wrapper/scripts-<editor>.js with the generated list
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');

const args = process.argv.slice(2);
const getArg = n => { const i = args.indexOf(n); return i >= 0 ? args[i + 1] : undefined; };
const SDK_DIR = path.resolve(getArg('--sdk-dir') || process.env.OO_SDK_DIR || '');
const MANIFEST_DIR = path.resolve(getArg('--manifest-dir') || path.join(__dirname, '..'));
const WRITE = args.includes('--write');
const EDITORS = getArg('--editor') ? [getArg('--editor')] : ['word', 'cell', 'slide'];

if (!SDK_DIR || !fs.existsSync(path.join(SDK_DIR, 'configs'))) {
  console.error('ERROR: --sdk-dir must point at an sdkjs checkout containing configs/. Got: ' + SDK_DIR);
  process.exit(2);
}

const PREPEND = ['vendor/polyfill.js', 'common/AllFonts.js', 'common/applyDocumentChanges.js'];
const EXTRA_INSERTS = {
  cell:  [{ after: 'word/apiCommon.js', file: 'common/zlib/zlib.js' }],
  slide: [{ after: 'word/apiCommon.js', file: 'common/zlib/zlib.js' }],
};

function generate(editor) {
  const cfg = JSON.parse(fs.readFileSync(path.join(SDK_DIR, 'configs', `${editor}.json`), 'utf8')).sdk;
  let list = [...PREPEND, ...cfg.min, ...cfg.common];
  for (const ins of (EXTRA_INSERTS[editor] || [])) {
    const at = list.indexOf(ins.after);
    if (at < 0) throw new Error(`anchor "${ins.after}" not found for ${editor} — upstream layout changed`);
    if (!list.includes(ins.file)) list.splice(at + 1, 0, ins.file);
  }
  return list;
}

function render(list) {
  return 'var sdk_scripts = [\n' + list.map(p => `\t"../${p}"`).join(',\n') + '\n];\n';
}

function committedPaths(editor) {
  const p = path.join(MANIFEST_DIR, `scripts-${editor}.js`);
  if (!fs.existsSync(p)) return null;
  return [...fs.readFileSync(p, 'utf8').matchAll(/"\.\.([^"]+)"/g)].map(m => m[1].replace(/^\//, ''));
}

let anyDiff = false;
for (const ed of EDITORS) {
  const gen = generate(ed);
  if (WRITE) {
    fs.writeFileSync(path.join(MANIFEST_DIR, `scripts-${ed}.js`), render(gen), 'utf8');
    console.log(`[${ed}] wrote ${gen.length} entries`);
    continue;
  }
  const cur = committedPaths(ed);
  if (!cur) { console.log(`[${ed}] generated ${gen.length} (no committed manifest to compare)`); continue; }
  let fd = -1;
  for (let i = 0; i < Math.max(gen.length, cur.length); i++) if (gen[i] !== cur[i]) { fd = i; break; }
  if (fd === -1 && gen.length === cur.length) {
    console.log(`[${ed}] ✅ generated == committed (${gen.length} entries)`);
  } else {
    anyDiff = true;
    console.log(`[${ed}] ⚠️ differs (gen ${gen.length} vs committed ${cur.length}); first diff @${fd}: gen=${gen[fd]} committed=${cur[fd]}`);
  }
}
process.exit(anyDiff ? 1 : 0);
