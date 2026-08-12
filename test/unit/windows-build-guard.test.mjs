/**
 * windows-build-guard.test.mjs — structural guard for Windows dev bootstrap.
 * [adr:0001]
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
const pkg = JSON.parse(readFileSync(path.join(REPO, 'package.json'), 'utf8'));
const pullUpstream = readFileSync(path.join(REPO, 'build/01-pull-upstream.sh'), 'utf8');
const assemble = readFileSync(path.join(REPO, 'build/04-assemble-dist.sh'), 'utf8');

describe('windows build guard [adr:0001]', () => {
  it('routes initialize and build through run-bash.js', () => {
    expect(pkg.scripts.initialize).toContain('run-bash.js');
    expect(pkg.scripts.build).toContain('run-bash.js');
  });

  it('get_field reads JSON via process.argv', () => {
    expect(pullUpstream).toMatch(/get_field\(\)/);
    expect(pullUpstream).toMatch(/process\.argv\[1\]/);
    expect(pullUpstream).not.toMatch(/readFileSync\('\$1'/);
    expect(pullUpstream).not.toMatch(/readFileSync\('\$REPO/);
  });

  it('04-assemble reads config.json via process.argv when using inline node', () => {
    const configBlock = assemble.slice(assemble.indexOf('config.json'));
    expect(configBlock).toMatch(/process\.argv\[1\]/);
    expect(configBlock).not.toMatch(/readFileSync\('\$REPO/);
  });

  it('inject-boot uses node, not python3', () => {
    const injectBoot = readFileSync(
      path.join(REPO, 'wrapper/v1/web-apps-overlay/inject-boot.sh'),
      'utf8',
    );
    expect(injectBoot).toMatch(/inject-into-html\.js/);
    expect(injectBoot).not.toMatch(/python3\s+-/);
  });
});
