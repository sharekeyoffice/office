#!/usr/bin/env node
/**
 * bundle.js — wrapper editor/viewer bundle generator (relocatable).
 *
 * Identical build logic to the original (pre-consolidation) bundle.js, but
 * decoupled from living *inside* the sdkjs tree so it can run from
 * sharekey-office/wrapper/. Path resolution is now explicit:
 *
 *   --sdk-dir DIR        root of the sdkjs source checkout (the files the
 *                        manifests reference via "../path"). Default: env
 *                        OO_SDK_DIR, else two levels up (legacy in-tree mode).
 *   --manifest-dir DIR   where scripts-{editor}.js[.viewer.js] live.
 *                        Default: the dir containing this script's parent
 *                        (i.e. wrapper/), else legacy VIEWER_DIR.
 *   --out-dir DIR        output dir for bundles. Default: <manifest-dir>/v1/bundle
 *                        (editor mode) or <manifest-dir>/bundle (viewer mode).
 *
 *   --mode viewer|editor   (default viewer)
 *   --editor word|cell|slide   (default: all three)
 *   --report-only          size only, no minify, no write
 *   --skip-gzip            skip .gz/.br companions
 *
 * Build settings are kept byte-identical to the original so regenerated
 * bundles can be diffed against the committed reference bundles for an
 * equivalence proof. (NOTE: legalComments:'none' strips ONLYOFFICE headers;
 * that is a deliberate compliance divergence to revisit — see
 * wrapper/REGENERATE.md — but kept as-is here for the equivalence test.)
 */
'use strict';

const fs    = require('node:fs');
const path  = require('node:path');
const zlib  = require('node:zlib');

const args = process.argv.slice(2);
function getArg(name) {
  const i = args.indexOf(name);
  return i >= 0 ? args[i + 1] : undefined;
}

// Manifests live with this script's parent dir (wrapper/) by default.
const SELF_DIR     = __dirname;                          // .../wrapper/build
const MANIFEST_DIR = path.resolve(getArg('--manifest-dir') || path.join(SELF_DIR, '..'));
// sdkjs source root: explicit flag > env > legacy "two up from manifest".
const SDK_DIR      = path.resolve(
  getArg('--sdk-dir') || process.env.OO_SDK_DIR || path.join(MANIFEST_DIR, '..')
);

const FLAGS = {
  reportOnly: args.includes('--report-only'),
  skipGzip:   args.includes('--skip-gzip'),
  editor:     getArg('--editor'),
  mode:       getArg('--mode') || 'viewer',
};
if (FLAGS.mode !== 'viewer' && FLAGS.mode !== 'editor') {
  console.error('--mode must be "viewer" or "editor", got: ' + FLAGS.mode);
  process.exit(1);
}

function outDirForMode(mode) {
  const explicit = getArg('--out-dir');
  if (explicit) return path.resolve(explicit);
  return mode === 'editor'
    ? path.join(MANIFEST_DIR, 'v1', 'bundle')
    : path.join(MANIFEST_DIR, 'bundle');
}

const EDITORS = FLAGS.editor ? [FLAGS.editor] : ['word', 'cell', 'slide'];

function fmtSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / 1024 / 1024).toFixed(2) + ' MB';
}

function readManifest(editor) {
  const suffix = FLAGS.mode === 'editor' ? '.js' : '.viewer.js';
  const manifestPath = path.join(MANIFEST_DIR, `scripts-${editor}${suffix}`);
  const body = fs.readFileSync(manifestPath, 'utf8');
  const paths = [];
  const re = /"\.\.([^"]+)"/g;
  let m;
  while ((m = re.exec(body)) !== null) paths.push(m[1]);
  return paths;
}

async function concatAndMinify(editor, paths) {
  const esbuild = require('esbuild');
  const chunks = [];
  let totalRaw = 0;
  let missing = 0;
  for (const rel of paths) {
    const full = path.join(SDK_DIR, rel.replace(/^\//, ''));
    if (!fs.existsSync(full)) {
      console.warn(`  [warn] missing: ${rel}`);
      missing++;
      continue;
    }
    const src = fs.readFileSync(full, 'utf8');
    totalRaw += Buffer.byteLength(src, 'utf8');
    chunks.push(`/* ${rel} */\n${src}\n;\n`);
  }
  if (missing) console.warn(`  [warn] ${missing} files missing for ${editor}`);
  const source = chunks.join('');
  const t0 = Date.now();
  const result = await esbuild.transform(source, {
    minifyWhitespace: true,
    minifySyntax:    true,
    minifyIdentifiers: false,   // CRITICAL: cross-file top-level var sharing
    target: ['es2017'],
    legalComments: 'none',
    sourcemap: false,
  });
  return { source: result.code, totalRaw, minTime: Date.now() - t0 };
}

async function buildOne(editor) {
  console.log(`\n[${editor}]  sdk-dir=${SDK_DIR}`);
  const paths = readManifest(editor);
  console.log(`  manifest: ${paths.length} files`);
  if (FLAGS.reportOnly) {
    let totalRaw = 0;
    for (const rel of paths) {
      const full = path.join(SDK_DIR, rel.replace(/^\//, ''));
      if (fs.existsSync(full)) totalRaw += fs.statSync(full).size;
    }
    console.log(`  raw concat: ${fmtSize(totalRaw)}`);
    return;
  }
  const { source: minified, totalRaw, minTime } = await concatAndMinify(editor, paths);
  const minifiedBytes = Buffer.byteLength(minified, 'utf8');
  console.log(`  raw concat: ${fmtSize(totalRaw)}`);
  console.log(`  minified:   ${fmtSize(minifiedBytes)}  (${minTime} ms)`);
  const outDir = outDirForMode(FLAGS.mode);
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });
  const outName = FLAGS.mode === 'editor' ? `${editor}.editor.bundle.js` : `${editor}.bundle.js`;
  const outPath = path.join(outDir, outName);
  fs.writeFileSync(outPath, minified, 'utf8');
  if (!FLAGS.skipGzip) {
    fs.writeFileSync(outPath + '.gz', zlib.gzipSync(minified, { level: 9 }));
    fs.writeFileSync(outPath + '.br', zlib.brotliCompressSync(minified, {
      params: { [zlib.constants.BROTLI_PARAM_QUALITY]: 11 }
    }));
  }
  console.log(`  wrote: ${outPath}`);
}

(async function main() {
  for (const editor of EDITORS) await buildOne(editor);
})().catch(err => { console.error('bundle failed:', err); process.exit(1); });
