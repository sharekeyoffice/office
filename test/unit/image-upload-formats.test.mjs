import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import vm from 'node:vm';

const require = createRequire(import.meta.url);
const decoder = require('../../wrapper/v1/tiff-decoder.js');
const read = name => readFileSync(new URL('../fixtures/images/' + name, import.meta.url));
const pixels = [
  255, 0, 0, 255, 0, 255, 0, 128, 0, 0, 255, 0,
  255, 255, 255, 255, 0, 0, 0, 255, 255, 255, 0, 1
];

function decodeFixture(name) {
  const bytes = read(name);

  return decoder.decode(bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.length));
}

describe('TIFF decoding before image insertion', () => {
  it.each(['rgba', 'lzw', 'deflate', 'packbits'])('preserves %s pixels and alpha', name => {
    const image = decodeFixture(name + '.tiff');

    expect([image.width, image.height]).toEqual([3, 2]);
    expect([...image.data]).toEqual(pixels);
  });

  it('retains orientation for the canvas transform', () => {
    expect(decodeFixture('rotated.tiff').orientation).toBe(6);
  });

  it('unpremultiplies associated alpha before drawing into a canvas', () => {
    expect([...decodeFixture('associated-alpha.tiff').data]).toEqual([128, 64, 0, 128]);
  });

  it('rejects an unsupported color space instead of producing a blank image', () => {
    expect(() => decodeFixture('lab.tiff')).toThrow('color format is not supported');
  });

  it('uses the first page of a multipage TIFF as the inserted image', () => {
    expect([...decodeFixture('multipage.tiff').data]).toEqual(pixels);
  });

  it('rejects a file without a TIFF image', () => {
    expect(() => decoder.decode(new ArrayBuffer(0))).toThrow();
  });
});

function uploadHarness({ converterFails = false } = {}) {
  const source = readFileSync(new URL('../../wrapper/v1/editor-stubs.js', import.meta.url), 'utf8');
  const begin = source.indexOf('  function readFileAsDataURL(file)');
  const end = source.indexOf('  if (!patchUploadImageFiles())', begin);
  const svgMedia = [];
  const converted = [];
  const window = {
    Asc: { c_oAscError: { ID: { No: 0, UplImageUrl: -4 } } },
    AscCommon: {
      UploadImageFiles() {},
      g_oDocumentUrls: { addImageUrl: (...args) => svgMedia.push(args) }
    },
    parent: { X2TBridge: {
      addMedia: async () => 'blob:svg',
      convertImageToPng: async file => {
        converted.push(file.name);
        if (converterFails) {
          throw new Error('invalid image');
        }

        return 'data:image/png;base64,converted';
      }
    } }
  };
  class FileReader {
    readAsDataURL(file) {
      this.result = 'data:' + file.type + ';base64,original';
      queueMicrotask(() => this.onload());
    }
  }
  vm.runInNewContext(source.slice(begin, end) + '\npatchUploadImageFiles();',
    { window, FileReader, log() {} });

  function upload(files) {
    return new Promise(resolve => {
      window.AscCommon.UploadImageFiles(files, null, null, null, null, null, null,
        (error, urls) => resolve({ error, urls }));
    });
  }

  return { upload, svgMedia, converted };
}

function file(name, type) {
  return { name, type, arrayBuffer: async () => new ArrayBuffer(1) };
}

describe('local upload format handling', () => {
  it('normalizes BMP/TIFF and preserves mixed-file order', async () => {
    const harness = uploadHarness();
    const result = await harness.upload([
      file('image.bmp', 'image/bmp'), file('image.tiff', 'image/tiff'),
      file('image.png', 'image/png'), file('image.jpg', 'image/jpeg')
    ]);

    expect(result.error).toBe(0);
    expect(Array.from(result.urls)).toEqual([
      'data:image/png;base64,converted', 'data:image/png;base64,converted',
      'data:image/png;base64,original', 'data:image/jpeg;base64,original'
    ]);
    expect(harness.converted).toEqual(['image.bmp', 'image.tiff']);
  });

  it.each(['photo.BMP', 'photo.TIF', 'photo.TIFF'])('handles %s without MIME', async name => {
    expect((await uploadHarness().upload([file(name, '')])).urls[0])
      .toBe('data:image/png;base64,converted');
  });

  it.each(['image/bmp', 'image/x-ms-bmp', 'image/tiff'])('handles %s without an extension',
    async type => {
      expect((await uploadHarness().upload([file('image', type)])).urls[0])
        .toBe('data:image/png;base64,converted');
    });

  it('reports conversion failure through the existing upload error callback', async () => {
    expect((await uploadHarness({ converterFails: true }).upload([file('bad.tif', '')])).error)
      .toBe(-4);
  });

  it('keeps the SVG media registration path', async () => {
    const harness = uploadHarness();

    expect((await harness.upload([file('image.svg', 'image/svg+xml')])).urls[0]).toBe('blob:svg');
    expect(harness.svgMedia).toHaveLength(1);
    expect(harness.svgMedia[0][0]).toMatch(/\.svg$/);
    expect(harness.converted).toHaveLength(0);
  });
});
