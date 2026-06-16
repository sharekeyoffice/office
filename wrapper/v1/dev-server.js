#!/usr/bin/env node
// dev-server.js — minimal static server for the v1 wrapper.
//
// Self-contained: serves everything from THIS DIRECTORY, except for the
// OnlyOffice web-apps deploy tree (the editor UI), which is too large to
// bundle inline and is built separately via `grunt deploy-{document,
// spreadsheet,presentation}editor`.
//
// Expected layout (relative to dev-server.js):
//   ./                              wrapper sources (wrapper-*.js, edit*.html, editor-stubs.js, x2t-bridge.js, stubs.js)
//   ./fonts/                        TTF font files (URL: /fonts/*)
//   ./x2t/                          x2t.js + x2t.wasm    (URL: /x2t/*)
//   ./sdk-runtime/AllFonts.js                            (URL: /sdkjs/common/AllFonts.js)
//   ./sdk-runtime/libfont/engine/{fonts.js,fonts.wasm}   (URL: /sdkjs/common/libfont/engine/*)
//   ./bundle/<editor>.editor.bundle.js[.gz|.br]          (URL: /sdkjs/<editor>/sdk-all-min.js)
//
// External, configured via env or auto-detect:
//   WEB_APPS_DEPLOY                 path to web-apps/deploy/ (sibling of sdkjs/, by default)
//                                   Serves /web-apps/* and /sdkjs/* (anything not handled above)
//
// Listens on env PORT (default 8080). No deps beyond Node stdlib.

'use strict';

const fs   = require('fs');
const http = require('http');
const path = require('path');
const url  = require('url');

const PORT     = parseInt(process.env.PORT || '8080', 10);
const SELF_DIR = __dirname;

// Resolve web-apps/deploy: env var wins, else walk up from SELF_DIR looking
// for a `web-apps/deploy` sibling. The wrapper itself doesn't need this dir
// to load — only the editor UI does. If it's missing, the dev-server still
// starts but `/web-apps/*` and most `/sdkjs/*` requests will 404.
let WEB_APPS_RESOLVED = process.env.WEB_APPS_DEPLOY || null;
if (!WEB_APPS_RESOLVED) {
  let cur = SELF_DIR;
  for (let i = 0; i < 10 && cur !== '/'; i++) {
    const cand = path.join(cur, 'web-apps', 'deploy');
    if (fs.existsSync(cand)) { WEB_APPS_RESOLVED = cand; break; }
    cur = path.dirname(cur);
  }
}
console.log('[dev-server] v1 dir:', SELF_DIR);
console.log('[dev-server] web-apps deploy:', WEB_APPS_RESOLVED || '(not found — editor UI requests will 404)');

const MIME = {
  '.html':'text/html; charset=utf-8',
  '.js':  'application/javascript; charset=utf-8',
  '.json':'application/json; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg':'image/jpeg',
  '.gif': 'image/gif',
  '.ico': 'image/x-icon',
  '.wasm':'application/wasm',
  '.woff':'font/woff',
  '.woff2':'font/woff2',
  '.ttf': 'font/ttf'
};

function send404(res, requested) {
  res.writeHead(404, { 'Content-Type': 'text/plain; charset=utf-8' });
  res.end('Not found: ' + requested);
}

function serveFile(res, abs, requested, extraHeaders) {
  fs.stat(abs, (err, st) => {
    if (err || !st.isFile()) return send404(res, requested);
    const ext = path.extname(abs).toLowerCase();
    const headers = Object.assign({
      'Content-Type': MIME[ext] || 'application/octet-stream',
      'Cache-Control': 'no-cache',
      'Cross-Origin-Opener-Policy': 'same-origin',
      'Cross-Origin-Embedder-Policy': 'require-corp'
    }, extraHeaders || {});
    res.writeHead(200, headers);
    fs.createReadStream(abs).pipe(res);
  });
}

// Wrapper-source files served straight from v1/.
const SELF_FILES = new Set([
  '/edit.html',
  '/edit-host-demo.html',
  '/wrapper-boot.js',
  '/wrapper-mount.js',
  '/wrapper-customization.js',
  '/wrapper-postmessage.js',
  '/editor-stubs.js',
  '/x2t-bridge.js',
  '/stubs.js'
]);

http.createServer((req, res) => {
  const u = url.parse(req.url);
  let p = decodeURIComponent(u.pathname || '/');

  if (p === '/') p = '/edit-host-demo.html';

  // 1) Wrapper source files.
  if (SELF_FILES.has(p)) {
    return serveFile(res, path.join(SELF_DIR, p), p);
  }

  // 2) Service-worker stub — editor's boot tries to register one; serve empty 200.
  if (p === '/document_editor_service_worker.js') {
    res.writeHead(200, { 'Content-Type': MIME['.js'], 'Cache-Control': 'no-cache' });
    res.end('// stub — wrapper does not use a service worker\n');
    return;
  }

  // 3) Cosmetic stubs that the editor expects but we don't ship.
  if (p === '/sdkjs/common/Charts/ChartStyles.js') {
    res.writeHead(200, { 'Content-Type': MIME['.js'], 'Cache-Control': 'no-cache' });
    res.end('// stub — chart styles not bundled\n');
    return;
  }
  if (p === '/sdkjs/common/Images/cursors/svg.json') {
    res.writeHead(200, { 'Content-Type': MIME['.json'], 'Cache-Control': 'no-cache' });
    res.end('{}');
    return;
  }
  // Coauthoring socket endpoint — DocServer-only, doesn't apply to wrapper.
  if (p.startsWith('/doc/') && p.includes('shardkey=') && p.includes('transport=polling')) {
    res.writeHead(200, { 'Content-Type': 'text/plain', 'Cache-Control': 'no-cache' });
    res.end('1:6');  // socket.io v4 noop frame
    return;
  }
  // Toolbar/anchor sprite icons that aren't critical for rendering.
  if (/^\/sdkjs\/common\/Images\/icons\/.*\.(png|svg|gif)$/.test(p)) {
    const tinyPng = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=', 'base64');
    res.writeHead(200, { 'Content-Type': MIME['.png'], 'Cache-Control': 'public, max-age=3600' });
    res.end(tinyPng);
    return;
  }

  // 4) sdkjs runtime files we ship in v1/sdk-runtime/ — the font manifest
  // (AllFonts.js) and HarfBuzz engine (fonts.js + fonts.wasm). Layout
  // mirrors the URL: `sdk-runtime/common/AllFonts.js` is served at
  // `/sdkjs/common/AllFonts.js`. Only fall through to web-apps deploy if
  // the file isn't shipped here.
  if (p === '/sdkjs/common/AllFonts.js' ||
      p === '/sdkjs/common/libfont/engine/fonts.js' ||
      p === '/sdkjs/common/libfont/engine/fonts.wasm') {
    const local = path.join(SELF_DIR, 'sdk-runtime', p.slice('/sdkjs/'.length));
    if (fs.existsSync(local)) return serveFile(res, local, p);
  }

  // 5) x2t WASM artifacts.
  if (p.startsWith('/x2t/')) {
    const sub = p.slice('/x2t/'.length);
    return serveFile(res, path.join(SELF_DIR, 'x2t', sub), p);
  }

  // 6) Bundled fonts.
  if (p.startsWith('/fonts/')) {
    const sub = p.slice('/fonts/'.length);
    return serveFile(res, path.join(SELF_DIR, 'fonts', sub), p);
  }

  // 7) Editor-mode SDK bundles. Negotiates Content-Encoding (br > gz >
  // identity) using pre-built companion files.
  {
    const sdkAllMatch = /^\/sdkjs\/(word|cell|slide)\/sdk-all(-min)?\.js$/.exec(p);
    if (sdkAllMatch) {
      const editor = sdkAllMatch[1];
      const accept = (req.headers['accept-encoding'] || '').toLowerCase();
      const wantsBr = accept.includes('br');
      const wantsGz = accept.includes('gzip');
      const bundlePath = path.join(SELF_DIR, 'bundle', editor + '.editor.bundle.js');
      if (fs.existsSync(bundlePath)) {
        if (wantsBr && fs.existsSync(bundlePath + '.br')) {
          return serveFile(res, bundlePath + '.br', p, { 'Content-Encoding': 'br', 'Vary': 'Accept-Encoding' });
        }
        if (wantsGz && fs.existsSync(bundlePath + '.gz')) {
          return serveFile(res, bundlePath + '.gz', p, { 'Content-Encoding': 'gzip', 'Vary': 'Accept-Encoding' });
        }
        return serveFile(res, bundlePath, p);
      }
      return send404(res, p + ' (run bundle.js to produce this)');
    }
  }

  // 8) OnlyOffice web-apps deploy (editor UI). Externally built — see
  // README/integration guide. Anything under /web-apps/* or /sdkjs/* not
  // handled above falls through here.
  if ((p.startsWith('/web-apps/') || p.startsWith('/sdkjs/')) && WEB_APPS_RESOLVED) {
    const safe = path.normalize(p).replace(/^(\.\.[\/\\])+/, '');
    const abs  = path.join(WEB_APPS_RESOLVED, safe);
    if (!abs.startsWith(WEB_APPS_RESOLVED)) return send404(res, p);
    return serveFile(res, abs, p);
  }

  // 9) Test-fixture convenience routes — useful for the standalone smoke
  // path; configurable via env so prod ignores them.
  if (process.env.TEST_FIXTURES_DIR && p.startsWith('/test-fixtures/')) {
    const sub = p.slice('/test-fixtures/'.length);
    const candidate = path.join(process.env.TEST_FIXTURES_DIR, sub);
    if (fs.existsSync(candidate)) return serveFile(res, candidate, p);
  }

  // 10) Last resort: any other file under v1/ (e.g. user-added assets).
  const localCandidate = path.join(SELF_DIR, p);
  if (localCandidate.startsWith(SELF_DIR)) {
    fs.stat(localCandidate, (err, st) => {
      if (!err && st.isFile()) return serveFile(res, localCandidate, p);
      return send404(res, p);
    });
    return;
  }

  send404(res, p);
}).listen(PORT, () => {
  console.log('[dev-server] listening on http://localhost:' + PORT);
  console.log('[dev-server] open http://localhost:' + PORT + '/edit-host-demo.html');
});
