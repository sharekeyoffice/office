#!/usr/bin/env node
// test/server.js — serves test/index.html on a different origin from
// the editor, so we can simulate the cross-origin main app ↔ editor
// flow locally.
//
// Run:  node test/server.js          # listens on :3000
//       PORT=3333 node test/server.js
//
// Editor is expected to run separately on :8080 (npm start in repo root).

'use strict';

const fs   = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const PORT = parseInt(process.env.PORT || '3000', 10);
const DIR  = __dirname;
const INDEX = path.join(DIR, 'index.html');

if (!fs.existsSync(INDEX)) {
  console.error('ERROR: test/index.html missing');
  process.exit(1);
}

http.createServer((req, res) => {
  const url = req.url.split('?')[0];
  // Only serve test/index.html and any sibling files inside test/.
  // Anything else: 404. Matches the editor server's no-SPA-fallback rule.
  let filePath;
  if (url === '/' || url === '/index.html') {
    filePath = INDEX;
  } else {
    const safe = path.normalize(url).replace(/^(\.\.[/\\])+/, '');
    filePath = path.join(DIR, safe);
    if (!filePath.startsWith(DIR)) {
      res.writeHead(403); res.end('Forbidden'); return;
    }
  }
  fs.stat(filePath, (err, st) => {
    if (err || !st.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('Not found: ' + url);
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    const mime = ext === '.html' ? 'text/html; charset=utf-8'
               : ext === '.js'   ? 'application/javascript; charset=utf-8'
               : ext === '.css'  ? 'text/css; charset=utf-8'
               : 'application/octet-stream';
    res.writeHead(200, {
      'Content-Type': mime,
      'Cache-Control': 'no-cache',
    });
    fs.createReadStream(filePath).pipe(res);
  });
}).listen(PORT, () => {
  console.log('[test-host] listening on http://localhost:' + PORT);
  console.log('[test-host] open http://localhost:' + PORT + '/');
  console.log();
  console.log('  Editor is expected at http://localhost:8080');
  console.log('  (start it with `npm start` in another terminal)');
  console.log();
  console.log('  IMPORTANT: edit public/edit.html and replace the');
  console.log('  __ALLOWED_HOST_ORIGIN__ placeholder with this URL:');
  console.log('    http://localhost:' + PORT);
});
