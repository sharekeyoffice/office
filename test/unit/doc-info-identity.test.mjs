// Guard check for ADR-0001 — the editor always has a user identity.
//
// editor-stubs.js runs inside the editor iframe and depends on the SDK being
// loaded, so it cannot be executed here. This guard pins the wiring instead:
// the identity has to reach asc_CDocInfo, and it has to reach the iframe
// before the document opens. Both are easy to drop by accident during an
// OnlyOffice upgrade, and neither fails loudly — the editor keeps working
// until someone types inside a tracked change.

import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const read = name => readFileSync(join(REPO, 'wrapper', 'v1', name), 'utf8');

describe('the wrapper puts a user on DocInfo [adr:0001][fp:ed3daf56]', () => {
	const stubs = read('editor-stubs.js');

	it('builds an asc_CUserInfo and sets its id and full name', () => {
		expect(stubs).toMatch(/new\s+window\.Asc\.asc_CUserInfo\s*\(\s*\)/);
		expect(stubs).toMatch(/\.put_Id\s*\(/);
		expect(stubs).toMatch(/\.put_FullName\s*\(/);
	});

	it('attaches it to DocInfo before handing DocInfo to the editor', () => {
		const attach = stubs.indexOf('.put_UserInfo(');
		const handOver = stubs.indexOf('asc_setDocInfo(info)');

		expect(attach).toBeGreaterThan(-1);
		expect(handOver).toBeGreaterThan(-1);
		// asc_setDocInfo copies the user out of DocInfo into api.User, so a
		// UserInfo attached afterwards would never reach the editor.
		expect(attach).toBeLessThan(handOver);
	});

	it('never leaves the name empty, whatever the host sent', () => {
		// getUserInitials does `username.split(' ')` with no guard, so an empty
		// name is as fatal as a missing one.
		expect(stubs).toMatch(/if\s*\(!u\.fullname\)\s*u\.fullname\s*=/);
		expect(stubs).toMatch(/if\s*\(!u\.id\)\s*u\.id\s*=/);
	});
});

describe('the identity reaches the editor [adr:0001][fp:ed3daf56]', () => {
	it('wrapper-mount passes user options into buildEditorConfig', () => {
		// A bare buildEditorConfig(type) call is the bug this ADR fixes: the
		// host identity never had a way in.
		expect(read('wrapper-mount.js')).toMatch(/buildEditorConfig\s*\(\s*type\s*,/);
	});

	it('wrapper-postmessage primes the iframe before the document opens', () => {
		const pm = read('wrapper-postmessage.js');
		const prime = pm.indexOf('self.primeUserIdentity()');
		const open = pm.indexOf('dispatchWhenReady');

		expect(prime).toBeGreaterThan(-1);
		expect(open).toBeGreaterThan(-1);
		// Revisions copy their author when they are created, so an identity
		// that lands after the document opened is already too late.
		expect(prime).toBeLessThan(open);
	});

	it('wrapper-postmessage ignores a blank identity instead of storing it', () => {
		expect(read('wrapper-postmessage.js')).toMatch(/rememberHostUser/);
	});
});
