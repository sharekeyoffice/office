#!/usr/bin/env node
// run-bash.js — run a build/*.sh script with the right bash on each OS.
// On Windows, npm may pick WSL bash (no node on PATH). We require Git Bash.
// [adr:0001]
'use strict';

const { spawnSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');

const scriptArg = process.argv[2];
if (!scriptArg) {
  console.error('Usage: node build/run-bash.js <script-under-build/>');
  process.exit(1);
}

const repo = path.resolve(__dirname, '..');
const rel = scriptArg.replace(/\\/g, '/');
const script = path.resolve(repo, rel.startsWith('build/') ? rel : path.join('build', rel));

if (!fs.existsSync(script)) {
  console.error(`ERROR: script not found: ${script}`);
  process.exit(1);
}

function findGitBash() {
  const candidates = [
    process.env.ProgramFiles && path.join(process.env.ProgramFiles, 'Git', 'bin', 'bash.exe'),
    process.env['ProgramFiles(x86)'] && path.join(process.env['ProgramFiles(x86)'], 'Git', 'bin', 'bash.exe'),
    process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, 'Programs', 'Git', 'bin', 'bash.exe'),
  ].filter(Boolean);

  for (const bash of candidates) {
    if (fs.existsSync(bash)) return bash;
  }
  return null;
}

function isWslBash(bashPath) {
  const normalized = bashPath.replace(/\\/g, '/').toLowerCase();
  return normalized.includes('/windows/system32/bash.exe')
    || normalized.includes('/windowsapps/');
}

function resolveBash() {
  if (process.platform !== 'win32') return 'bash';

  const gitBash = findGitBash();
  if (gitBash) return gitBash;

  const which = spawnSync('where.exe', ['bash'], { encoding: 'utf8' });
  if (which.status === 0) {
    const first = which.stdout.trim().split(/\r?\n/)[0];
    if (first && isWslBash(first)) {
      console.error('ERROR: Git Bash not found, and the bash on PATH is WSL.');
      console.error('Install Git for Windows: https://git-scm.com/download/win');
      process.exit(1);
    }
    if (first && fs.existsSync(first)) return first;
  }

  console.error('ERROR: Git Bash not found. Install Git for Windows: https://git-scm.com/download/win');
  process.exit(1);
}

const bash = resolveBash();
const relScript = path.relative(repo, script).replace(/\\/g, '/');

const result = spawnSync(bash, [relScript], {
  cwd: repo,
  stdio: 'inherit',
  env: process.env,
});

process.exit(result.status ?? 1);
