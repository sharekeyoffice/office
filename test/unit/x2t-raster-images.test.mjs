import { beforeEach, describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { convert } from '../support/x2t-node.mjs';
import { unzip } from '../support/ooxml-normalize.mjs';

const require = createRequire(import.meta.url);
const bridge = require('../../wrapper/v1/x2t-bridge.js');
const read = relative => new Uint8Array(readFileSync(new URL('../fixtures/' + relative,
  import.meta.url)));
let converted;

beforeEach(() => {
  converted = [];
  bridge.configure({ rasterizeImage: async file => {
    converted.push(file.type);
    const png = read('images/' + (file.type === 'image/bmp' ? 'palette' : 'rgba') + '.png');

    return 'data:image/png;base64,' + Buffer.from(png).toString('base64');
  } });
});

for (const format of ['docx', 'xlsx', 'pptx']) {
  for (const imageType of ['bmp', 'tiff']) {
    it(`opens, saves and reopens a ${format} containing ${imageType.toUpperCase()}`, async () => {
      const input = read(`${format}/${imageType}-image.${format}`);
      const original = Uint8Array.from(input);
      const normalized = await bridge.normalizeRasterMedia(input);
      const parts = unzip(normalized);
      const pngName = [...parts.keys()].find(name => name.includes('/media/image.png'));

      expect(pngName).toBeDefined();
      expect([...parts.keys()].some(name => /\.(bmp|tiff?)$/i.test(name))).toBe(false);
      expect(parts.get(pngName).equals(Buffer.from(read('images/' +
        (imageType === 'bmp' ? 'palette' : 'rgba') + '.png')))).toBe(true);
      expect(parts.get('[Content_Types].xml').toString())
        .toContain('ContentType="image/png"');
      expect(input).toEqual(original);
      expect(converted).toEqual([imageType === 'bmp' ? 'image/bmp' : 'image/tiff']);

      const bin = await convert(normalized, { from: format, to: 'bin' });
      const saved = await convert(bin, { from: 'bin', to: format });
      const reopened = await convert(saved, { from: format, to: 'bin' });
      const savedAgain = await convert(reopened, { from: 'bin', to: format });
      const finalParts = unzip(savedAgain);

      expect([...finalParts.keys()].some(name => /\/media\/.*\.png$/i.test(name))).toBe(true);
      expect([...finalParts.keys()].some(name => /\.(bmp|tiff?)$/i.test(name))).toBe(false);
    });
  }
}

describe('raster media normalization boundaries', () => {
  it('handles name collisions, encoded names, quoted attributes and absolute targets', async () => {
    const input = read('docx/bmp-collision.docx');
    const before = unzip(input);
    const after = unzip(await bridge.normalizeRasterMedia(input));
    const rels = after.get('word/_rels/document.xml.rels').toString();
    const types = after.get('[Content_Types].xml').toString();

    expect(after.has('word/media/a & b.sk-raster-1.png')).toBe(true);
    expect(after.get('word/media/a & b.png')).toEqual(before.get('word/media/a & b.png'));
    expect(rels).toContain("Target='media/a%20%26%20b.sk-raster-1.png'");
    expect(rels).toContain("Target='/word/media/a%20%26%20b.sk-raster-1.png'");
    expect(rels).toContain("Target='https://example.invalid/a.bmp'");
    expect(types.match(/Extension=['"]png['"]/g)).toHaveLength(1);
    expect(types).toContain("PartName='/word/media/a%20%26%20b.sk-raster-1.png'");
    expect(types).toContain("ContentType='image/png'");
  });

  it('leaves an SVG package byte-for-byte unchanged', async () => {
    const input = read('docx/svg-image.docx');

    expect(await bridge.normalizeRasterMedia(input)).toBe(input);
    expect(converted).toEqual([]);
  });

  it('leaves unrelated OOXML parts byte-for-byte unchanged', async () => {
    const input = read('docx/bmp-image.docx');
    const before = unzip(input);
    const after = unzip(await bridge.normalizeRasterMedia(input));

    for (const [name, bytes] of before) {
      if (name === '[Content_Types].xml' || name.endsWith('.rels') || name.endsWith('.bmp')) {
        continue;
      }

      expect(after.get(name), name).toEqual(bytes);
    }
  });

  it('does not hide a rasterizer error or alter the input bytes', async () => {
    const input = read('docx/bmp-image.docx');
    const before = Uint8Array.from(input);
    bridge.configure({ rasterizeImage: async () => { throw new Error('could not decode'); } });

    await expect(bridge.normalizeRasterMedia(input)).rejects.toThrow('could not decode');
    expect(input).toEqual(before);
  });

  it('rejects a converter result that is not a PNG', async () => {
    bridge.configure({ rasterizeImage: async () => 'data:image/png;base64,AAAA' });

    await expect(bridge.normalizeRasterMedia(read('docx/bmp-image.docx')))
      .rejects.toThrow('invalid PNG');
  });
});
