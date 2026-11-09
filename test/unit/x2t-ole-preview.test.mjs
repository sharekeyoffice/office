/**
 * x2t-ole-preview.test.mjs — a pptx whose OLE objects carry EMF/WMF previews
 * opens in x2t-wasm, and one x2t crash does not break the rest of the session.
 *
 * Background. x2t-wasm corrupts its heap while it converts an EMF preview of
 * an OLE object to SVG (think-cell data objects, embedded Excel, Visio). The
 * trap ("memory access out of bounds", "table index is out of bounds",
 * "null function") depends on the heap layout, so the x2t checks below run
 * each input in several fresh modules with shifted layouts.
 *
 * Fixture test/fixtures/pptx/ole-emf-preview.pptx is synthetic: minimal.pptx
 * plus one hidden OLE object whose preview is a header-only EMF, built the
 * way think-cell's "do not delete" objects are.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { unzip } from '../support/ooxml-normalize.mjs';

const require = createRequire(import.meta.url);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, '..', '..');
const BRIDGE_PATH = path.join(REPO, 'wrapper', 'v1', 'x2t-bridge.js');
const FIX = (rel) => path.join(REPO, 'test', 'fixtures', rel);
const read = (rel) => new Uint8Array(readFileSync(FIX(rel)));

const Bridge = require(BRIDGE_PATH);

const OLE_FIXTURE = 'pptx/ole-emf-preview.pptx';
const PREVIEW_PNG = 'ppt/media/e.sk-ole-preview.png';

describe('sanitizeOlePreviews', () => {
  it('repoints the OLE preview from the EMF to a PNG', async () => {
    const { bytes, replaced } = await Bridge.sanitizeOlePreviews(read(OLE_FIXTURE));
    expect(replaced).toBe(1);

    const parts = unzip(bytes);
    const rels = parts.get('ppt/slides/_rels/slide1.xml.rels').toString('utf8');
    expect(rels).toContain('Target="../media/e.sk-ole-preview.png"');
    expect(rels).not.toContain('Target="../media/e.emf"');

    const png = parts.get(PREVIEW_PNG);
    expect(png.subarray(0, 8)).toEqual(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
    expect(parts.get('[Content_Types].xml').toString('utf8')).toMatch(/<Default Extension="png"/);
  });

  it('keeps every other part byte for byte, the OLE data and the EMF included', async () => {
    const before = unzip(read(OLE_FIXTURE));
    const after = unzip((await Bridge.sanitizeOlePreviews(read(OLE_FIXTURE))).bytes);
    const changed = new Set(['ppt/slides/_rels/slide1.xml.rels', '[Content_Types].xml']);

    for (const [name, content] of before) {
      if (changed.has(name)) continue;
      expect(after.get(name), name).toEqual(content);
    }
    expect(after.has('ppt/media/e.emf')).toBe(true);
    expect(after.has('ppt/embeddings/o.bin')).toBe(true);
    expect([...after.keys()].sort()).toEqual([...before.keys(), PREVIEW_PNG].sort());
  });

  it('returns a pptx without OLE previews untouched', async () => {
    const input = read('pptx/minimal.pptx');
    const r = await Bridge.sanitizeOlePreviews(input);
    expect(r.replaced).toBe(0);
    expect(r.bytes).toBe(input);
  });

  it('returns bytes it cannot read as a ZIP untouched', async () => {
    const input = new TextEncoder().encode('not a zip at all');
    const r = await Bridge.sanitizeOlePreviews(input);
    expect(r.replaced).toBe(0);
    expect(r.bytes).toBe(input);
  });
});

describe('x2t opens the sanitized pptx in any heap layout', () => {
  const PERTURBATIONS = [
    'dup0', 'dup1', 'dup3', 'dup8', 'dup20',
    `warm:${FIX('docx/minimal.docx')}`,
    `warm:${FIX('xlsx/minimal.xlsx')}`,
  ];

  const runPerturbed = (file, perturbation) => JSON.parse(
    execFileSync(process.execPath, [path.join(REPO, 'test', 'support', 'x2t-perturbed.mjs'), file, perturbation], {
      encoding: 'utf8',
      timeout: 60_000,
    }).trim(),
  );

  it('converts to PPTY in every layout', async () => {
    const { bytes } = await Bridge.sanitizeOlePreviews(read(OLE_FIXTURE));
    const { mkdtempSync, writeFileSync } = await import('node:fs');
    const { tmpdir } = await import('node:os');
    const file = path.join(mkdtempSync(path.join(tmpdir(), 'x2t-ole-')), 'sanitized.pptx');
    writeFileSync(file, bytes);

    const verdicts = PERTURBATIONS.map((p) => [p.split(':')[0], runPerturbed(file, p)]);
    for (const [p, v] of verdicts) expect(v, p).toEqual({ ok: true, sig: 'PPTY' });
  }, 120_000);
});

/**
 * Load the bridge in a sandbox with a fake x2t.js. Each injected <script>
 * builds a new "module" generation; `crashOn` lists the generations whose
 * main1 traps the way x2t-wasm does.
 *
 * Like the real x2t.js, the fake runs once per global scope (page or
 * iframe). The real one declares top-level classes, so a second copy in the
 * same page throws "Identifier 'ExitStatus' has already been declared" and
 * never initializes; the fake stays silent the same way.
 */
function loadBridgeWithFakeX2T({ crashOn }) {
  const generations = [];
  const frames = [];
  const window = {};

  const makeFS = () => {
    const files = new Map();
    const dirs = new Set(['/']);
    return {
      files,
      mkdir: (d) => { if (dirs.has(d)) throw new Error('EEXIST'); dirs.add(d); },
      writeFile: (p, data) => files.set(p, typeof data === 'string' ? new TextEncoder().encode(data) : data),
      readFile: (p) => { if (!files.has(p)) throw new Error('ENOENT ' + p); return files.get(p); },
      unlink: (p) => files.delete(p),
      readdir: (d) => ['.', '..', ...[...files.keys()].filter((p) => p.startsWith(d + '/')).map((p) => p.slice(d.length + 1))],
      stat: () => ({ mode: 0o100000 }),
    };
  };

  // A document whose <script> runs the fake x2t.js in `win`'s global scope.
  const makeDocument = (win) => {
    let x2tRan = false;
    const doc = {
      head: {
        appendChild: () => {
          if (x2tRan) return;          // second copy in one scope: SyntaxError, no init
          x2tRan = true;
          const Module = win.Module;
          const gen = generations.length + 1;
          Module.FS = makeFS();
          Module.ccall = (_name, _ret, _types, [paramsPath]) => {
            if (crashOn.includes(gen)) throw new Error('memory access out of bounds');
            const params = new TextDecoder().decode(Module.FS.readFile(paramsPath));
            const to = params.match(/<m_sFileTo>([^<]+)<\/m_sFileTo>/)[1];
            Module.FS.writeFile(to, new TextEncoder().encode('PK-out-from-gen-' + gen));
            return 0;
          };
          generations.push(Module);
          setTimeout(() => Module.onRuntimeInitialized(), 0);
        },
      },
      createElement: (tag) => {
        if (tag !== 'iframe') return {};
        const frameWin = {};
        const frame = { style: {}, setAttribute() {}, contentWindow: frameWin, parentNode: null };
        frame.contentDocument = makeDocument(frameWin);
        return frame;
      },
      body: {
        appendChild: (frame) => { frame.parentNode = doc.body; frames.push(frame); },
        removeChild: (frame) => { frame.parentNode = null; },
      },
    };
    return doc;
  };
  const document = makeDocument(window);

  const sandbox = {
    window, document, console, setTimeout, clearTimeout, performance,
    URL, Blob, Response, TextEncoder, TextDecoder, DecompressionStream, atob,
  };
  sandbox.window.location = { href: 'https://office.test/edit.html' };
  vm.createContext(sandbox);
  vm.runInContext(readFileSync(BRIDGE_PATH, 'utf8'), sandbox);
  return { bridge: window.X2TBridge, generations, frames };
}

describe('a crashed x2t module is replaced', () => {
  it('retries a crashed conversion on a fresh module and carries the media over', async () => {
    const { bridge, generations } = loadBridgeWithFakeX2T({ crashOn: [1] });
    await bridge.preload();
    generations[0].FS.writeFile('/working/media/image1.png', new Uint8Array([1, 2, 3]));

    const out = await bridge.convertFromBin(new Uint8Array([0x50, 0x50, 0x54, 0x59]), 'pptx');

    expect(new TextDecoder().decode(out)).toBe('PK-out-from-gen-2');
    expect(generations).toHaveLength(2);
    expect(generations[1]).not.toBe(generations[0]);
    // Array.from: the bridge's Uint8Array comes from the vm realm.
    expect(Array.from(generations[1].FS.readFile('/working/media/image1.png'))).toEqual([1, 2, 3]);
  });

  it('runs the fresh module in its own iframe, not a second copy in the page', async () => {
    const { bridge, generations, frames } = loadBridgeWithFakeX2T({ crashOn: [1] });

    await bridge.convertFromBin(new Uint8Array([1]), 'pptx');

    expect(frames).toHaveLength(1);
    expect(generations[1]).toBe(frames[0].contentWindow.Module);
  });

  it('reports the crash when the retry crashes too, and leaves a fresh module for the next call', async () => {
    const { bridge, generations, frames } = loadBridgeWithFakeX2T({ crashOn: [1, 2] });

    await expect(bridge.convertFromBin(new Uint8Array([1]), 'pptx'))
      .rejects.toThrow('x2t threw during main1: memory access out of bounds');

    const out = await bridge.convertFromBin(new Uint8Array([1]), 'pptx');
    expect(new TextDecoder().decode(out)).toBe('PK-out-from-gen-3');
    expect(generations).toHaveLength(3);
    // Each reset replaces the previous iframe; only the newest stays attached.
    expect(frames.map((f) => f.parentNode !== null)).toEqual([false, true]);
  });
});
