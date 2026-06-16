import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    // Tier A (unit) + Tier B (round-trip) run headless in Node — no browser.
    include: ['test/unit/**/*.test.mjs', 'test/roundtrip/**/*.test.mjs'],
    // x2t WASM init + conversions can take a moment on first load.
    testTimeout: 30_000,
    hookTimeout: 30_000,
  },
});
