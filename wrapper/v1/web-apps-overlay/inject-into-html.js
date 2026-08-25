#!/usr/bin/env node
// inject-into-html.js — insert a snippet before an anchor string in a file.
// Used by inject-boot.sh (Windows has no working python3 in Git Bash).
'use strict';

const fs = require('node:fs');

const [filePath, anchor, injected] = process.argv.slice(2);
if (!filePath || !anchor || injected === undefined) {
  console.error('Usage: node inject-into-html.js <file> <anchor> <injected>');
  process.exit(1);
}

const src = fs.readFileSync(filePath, 'utf8');
if (!src.includes(anchor)) {
  console.error(`ERROR: anchor not found in ${filePath}`);
  process.exit(1);
}

fs.writeFileSync(filePath, src.replace(anchor, injected + '    ' + anchor), 'utf8');
