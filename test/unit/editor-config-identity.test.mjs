// Unit check for ADR-0001 — the editor always has a user identity.
//
// wrapper-customization.js is browser glue: an IIFE that ends with `(window)`.
// It has no imports and no DOM calls, so it runs here with a plain object
// standing in for `window`.

import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const REPO = join(dirname(fileURLToPath(import.meta.url)), '..', '..');

function loadBuildEditorConfig() {
	const src = readFileSync(join(REPO, 'wrapper', 'v1', 'wrapper-customization.js'), 'utf8');
	const fakeWindow = {};
	// The file's own last line invokes the IIFE with `window`, so shadowing
	// `window` as a parameter is enough to capture the export.
	new Function('window', src)(fakeWindow);
	return fakeWindow.buildEditorConfig;
}

describe('buildEditorConfig user identity [adr:0001][fp:ed3daf56]', () => {
	it('uses the id and name the host supplied', () => {
		const built = loadBuildEditorConfig()('word', {
			userId: 'u-42',
			userName: 'Ada Lovelace'
		});

		expect(built.editorConfig.user.id).toBe('u-42');
		expect(built.editorConfig.user.name).toBe('Ada Lovelace');
	});

	it('falls back to a non-empty id and name when the host sends none', () => {
		const buildEditorConfig = loadBuildEditorConfig();

		for (const type of ['word', 'cell', 'slide']) {
			const user = buildEditorConfig(type).editorConfig.user;

			// Non-empty is the whole point. Common.Utils.getUserInitials calls
			// .split(' ') on the name with no null guard, so an empty or missing
			// name reaches the editor and throws on the first tracked change.
			expect(typeof user.id).toBe('string');
			expect(user.id.trim().length).toBeGreaterThan(0);
			expect(typeof user.name).toBe('string');
			expect(user.name.trim().length).toBeGreaterThan(0);
		}
	});

	it('falls back for an empty-string identity too, not only a missing one', () => {
		const user = loadBuildEditorConfig()('word', { userId: '', userName: '' })
			.editorConfig.user;

		expect(user.id.trim().length).toBeGreaterThan(0);
		expect(user.name.trim().length).toBeGreaterThan(0);
	});

	it('gives the fallback name a form a person can read', () => {
		const name = loadBuildEditorConfig()('word').editorConfig.user.name;

		// The fallback is written into the saved file and shown in the review
		// panel, so it must not be an internal token.
		expect(name).not.toMatch(/anon|null|undefined|^wrap-/i);
	});
});
