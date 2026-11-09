/**
 * x2t-svg-save.test.mjs — a document with an SVG picture saves, in all three
 * formats, and keeps both its SVG and a PNG fallback.
 *
 * Background. On save (bin → OOXML) x2t draws a PNG fallback from every SVG,
 * and x2t-wasm traps while drawing it ("memory access out of bounds"). The
 * bridge's save shim swaps each SVG in /working/media for a tagged PNG before
 * the save and repairs the output after it (see "SVG save shim" in
 * x2t-bridge.js). This suite drives those steps against the real headless x2t.
 *
 * Fixtures test/fixtures/{docx,xlsx,pptx}/svg-image.* are synthetic: the
 * minimal fixture plus one picture stored the way Office stores an SVG.
 */
import { describe, it, expect, beforeAll, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { loadX2T, convert } from '../support/x2t-node.mjs';
import { unzip } from '../support/ooxml-normalize.mjs';

const require = createRequire(import.meta.url);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIX = (rel) => path.join(HERE, '..', 'fixtures', rel);
const read = (rel) => new Uint8Array(readFileSync(FIX(rel)));
const Bridge = require(path.join(HERE, '..', '..', 'wrapper', 'v1', 'x2t-bridge.js'));
const { prepare, fixOutput, readPngTag, rasterSize } = Bridge._svgSave;

const PNG_SIG = [0x89, 0x50, 0x4e, 0x47];
const isPng = (b) => PNG_SIG.every((v, i) => b[i] === v);

/** A 2×1 PNG whose pixels encode n, so every drawn SVG gets distinct bytes. */
function fakePng(n) {
  const chunk = (type, data) => {
    const len = Buffer.alloc(4); len.writeUInt32BE(data.length);
    const crc = Buffer.alloc(4); crc.writeUInt32BE(zlib.crc32(Buffer.concat([Buffer.from(type), data])));
    return Buffer.concat([len, Buffer.from(type), data, crc]);
  };
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(2, 0); ihdr.writeUInt32BE(1, 4); ihdr[8] = 8; ihdr[9] = 6;
  const raw = Buffer.from([0, n & 255, 0, 0, 255, 0, n & 255, 0, 255]);
  return new Uint8Array(Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0)),
  ]));
}

let Module;
let drawn;

beforeAll(async () => {
  Module = await loadX2T();
  Bridge.configure({
    rasterizeSvg: async (svg) => { drawn.push(svg); return fakePng(drawn.length); },
  });
});

beforeEach(() => {
  drawn = [];
  for (const n of Module.FS.readdir('/working/media')) {
    if (n !== '.' && n !== '..') Module.FS.unlink('/working/media/' + n);
  }
});

/** Open, then save through the shim the way convertFromBin does. */
async function openAndSave(ext) {
  const source = read(`${ext}/svg-image.${ext}`);
  const originalSvg = [...unzip(source)].find(([n]) => n.endsWith('.svg'))[1];
  const bin = await convert(source, { from: ext, to: 'bin' });
  const state = await prepare(Module);
  const out = await convert(bin, { from: 'bin', to: ext });
  const fixed = await fixOutput(out, state);
  return { originalSvg, state, fixed, parts: unzip(fixed) };
}

for (const ext of ['docx', 'xlsx', 'pptx']) {
  describe(`saving a ${ext} with an SVG picture`, () => {
    it('swaps the SVG for a tagged PNG before the save', async () => {
      await convert(read(`${ext}/svg-image.${ext}`), { from: ext, to: 'bin' });
      const state = await prepare(Module);

      expect(state.count).toBe(1);
      expect(drawn).toHaveLength(1);
      const media = Module.FS.readFile('/working/media/' + state.originals['1'].name);
      expect(isPng(media)).toBe(true);
      expect(readPngTag(media)).toBe('1');
    });

    it('keeps the original SVG in the saved file', async () => {
      const { originalSvg, parts } = await openAndSave(ext);
      const svgParts = [...parts].filter(([n]) => /\.svg$/i.test(n));

      expect(svgParts.length).toBeGreaterThan(0);
      for (const [name, bytes] of svgParts) {
        expect(isPng(bytes), name + ' must hold SVG, not PNG').toBe(false);
        expect(bytes.equals(originalSvg), name).toBe(true);
      }
    });

    it('writes the drawn PNG as a .png fallback, and the relationships point at it', async () => {
      const { parts } = await openAndSave(ext);
      const pngParts = [...parts].filter(([n]) => /\.png$/i.test(n));

      expect(pngParts).toHaveLength(1);
      const [pngName, pngBytes] = pngParts[0];
      expect(readPngTag(pngBytes)).toBe('1');

      const relsTargets = [...parts]
        .filter(([n]) => n.endsWith('.rels'))
        .flatMap(([, b]) => [...b.toString('utf8').matchAll(/Target="([^"]+)"/g)].map((m) => m[1]));
      expect(relsTargets.some((t) => t.endsWith(path.posix.basename(pngName)))).toBe(true);
      expect(relsTargets.filter((t) => /\.svg$/i.test(t)).every((t) => parts.has(
        [...parts.keys()].find((k) => k.endsWith(path.posix.basename(t))),
      ))).toBe(true);
      expect(parts.get('[Content_Types].xml').toString('utf8')).toMatch(/Extension="png"/);
    });

    it('opens again after the save, and saves a second time', async () => {
      const { fixed } = await openAndSave(ext);
      const bin2 = await convert(fixed, { from: ext, to: 'bin' });
      const state2 = await prepare(Module);
      const out2 = await convert(bin2, { from: 'bin', to: ext });
      const parts2 = unzip(await fixOutput(out2, state2));

      expect([...parts2.keys()].some((n) => /\.svg$/i.test(n))).toBe(true);
      expect([...parts2].filter(([n]) => /\.svg$/i.test(n)).every(([, b]) => !isPng(b))).toBe(true);
    });
  });
}

describe('svg raster size', () => {
  it('keeps the aspect from the viewBox and sizes the long side to 1024', () => {
    expect(rasterSize('<svg viewBox="0 0 40 20">')).toEqual({ width: 1024, height: 512 });
  });

  it('honours width/height up to 2048', () => {
    expect(rasterSize('<svg width="1500" height="300">')).toEqual({ width: 1500, height: 300 });
    expect(rasterSize('<svg width="5000" height="2500">')).toEqual({ width: 2048, height: 1024 });
  });

  it('falls back to 2:1 when the SVG gives no size', () => {
    expect(rasterSize('<svg>')).toEqual({ width: 1024, height: 512 });
  });
});
