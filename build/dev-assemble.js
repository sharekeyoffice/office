#!/usr/bin/env node
// dev-assemble.js — fast dev loop for the EDITABLE layers only.
//
// Copies just the wrapper glue (`wrapper/v1/*.js`) + overlay assets
// (`overlay/`: edit.html, CSS, icons) into `public/`. The heavy trees
// (`web-apps/`, `sdkjs/`, `x2t/`, `fonts/`, the editor bundles) are produced
// by `npm run initialize` and are left untouched here — so a sync is a few KB,
// not ~200 MB.
//
//   node build/dev-assemble.js            # one-shot sync   (npm run sync)
//   node build/dev-assemble.js --watch    # sync + serve :8080 + re-sync on change   (npm run assemble)
//
// Env: ALLOWED_HOST_ORIGIN (default http://localhost:3000), PORT (8080).
// Prereq: run `npm run initialize` once first (builds public/ + the heavy trees).
//
// Note: re-sync does NOT auto-refresh the browser — hard-reload the editor tab.

'use strict';
const fs   = require('node:fs');
const path = require('node:path');
const cp   = require('node:child_process');

const REPO    = path.resolve(__dirname, '..');
const OVERLAY = path.join(REPO, 'overlay');
const V1      = path.join(REPO, 'wrapper', 'v1');
const PUBLIC  = path.join(REPO, 'public');
const ORIGIN  = process.env.ALLOWED_HOST_ORIGIN || 'http://localhost:3000';

// wrapper glue copied to the public/ root (NOT edit.html — overlay's wins;
// NOT bundle/fonts/x2t/sdk-runtime — those are heavy + built by initialize).
const GLUE = [
  'editor-stubs.js', 'wrapper-boot.js', 'wrapper-customization.js',
  'wrapper-mount.js', 'wrapper-postmessage.js', 'wrapper-heartbeat.js',
  'x2t-bridge.js', 'edit-host-demo.html',
];

function cpFile(src, dst) {
  fs.mkdirSync(path.dirname(dst), { recursive: true });
  fs.copyFileSync(src, dst);
}

function sync() {
  for (const f of GLUE) {
    const s = path.join(V1, f);
    if (fs.existsSync(s)) cpFile(s, path.join(PUBLIC, f));
  }
  for (const f of ['fonts.css', 'edit.css', 'connection-lost.css', 'turn-on-edit-mode.css', 'viewer-mode.css', 'welcome-screen.css']) {
    const s = path.join(OVERLAY, f);
    if (fs.existsSync(s)) cpFile(s, path.join(PUBLIC, f));
  }
  const iconsSrc = path.join(OVERLAY, 'icons');
  if (fs.existsSync(iconsSrc))
    for (const f of fs.readdirSync(iconsSrc))
      cpFile(path.join(iconsSrc, f), path.join(PUBLIC, 'icons', f));
  const imagesSrc = path.join(OVERLAY, 'images');
  if (fs.existsSync(imagesSrc))
    for (const f of fs.readdirSync(imagesSrc))
      cpFile(path.join(imagesSrc, f), path.join(PUBLIC, 'images', f));
  // edit.html with the dev origin baked in
  const html = fs.readFileSync(path.join(OVERLAY, 'edit.html'), 'utf8')
    .replace(/__ALLOWED_HOST_ORIGIN__/g, ORIGIN);
  fs.writeFileSync(path.join(PUBLIC, 'edit.html'), html);
  console.log(`[assemble] synced glue+overlay → public/ (origin ${ORIGIN})  ${new Date().toLocaleTimeString()}`);
}

// ── prereq check ─────────────────────────────────────────────
if (!fs.existsSync(path.join(PUBLIC, 'web-apps'))) {
  console.error('public/ is not built yet — run `npm run initialize` first.');
  process.exit(1);
}

sync();

if (process.argv.includes('--watch')) {
  const server = cp.spawn(process.execPath, [path.join(REPO, 'server', 'index.js')], {
    stdio: 'inherit',
    env: { ...process.env, PORT: process.env.PORT || '8080' },
  });
  process.on('SIGINT', () => { server.kill(); process.exit(0); });
  process.on('exit', () => server.kill());

  let timer = null;
  const onChange = () => { clearTimeout(timer); timer = setTimeout(sync, 150); };
  fs.watch(OVERLAY, { recursive: true }, onChange);   // edit.html, css, icons/
  fs.watch(V1,      { recursive: false }, onChange);  // top-level glue only (skips fonts/x2t/bundle subdirs)

  console.log('[assemble] watching overlay/ + wrapper/v1/ glue — edit, then hard-reload the editor tab. Ctrl-C to stop.');
}
