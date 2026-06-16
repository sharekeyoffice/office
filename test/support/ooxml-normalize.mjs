/**
 * ooxml-normalize.mjs — make two OOXML files content-comparable.
 *
 * OOXML (.docx/.xlsx/.pptx) is a ZIP. You CANNOT byte-compare two saves of the
 * "same" document: ZIP entry order, local-header timestamps, and embedded IDs
 * (w:rsid*, w14:paraId/textId, docProps timestamps, calcChain, slide revision
 * ids) all churn. This module unzips, drops/strips the volatile bits, lightly
 * canonicalizes each XML part, and returns a stable, diffable representation.
 *
 * SPIKE-QUALITY CAVEAT: XML canonicalization here is regex-based (sort
 * attributes, collapse inter-tag whitespace). That is good enough to prove the
 * round-trip approach but is NOT a real XML canonicalizer. For the production
 * suite, swap in a proper parser (parse → sort attrs → re-serialize). The
 * volatile-strip list below is also docx-leaning and will need per-format
 * passes (see VOLATILE_*).
 *
 * Pure Node: minimal ZIP reader via zlib.inflateRawSync, no dependencies.
 */
import zlib from 'node:zlib';
import crypto from 'node:crypto';

// ---------------------------------------------------------------------------
// Minimal ZIP reader (central-directory based)
// ---------------------------------------------------------------------------
const SIG_EOCD = 0x06054b50;
const SIG_CEN = 0x02014b50;
const SIG_LOC = 0x04034b50;

function findEOCD(buf) {
  // EOCD is at the end; scan backward (comment field is usually empty).
  const min = Math.max(0, buf.length - 22 - 0xffff);
  for (let i = buf.length - 22; i >= min; i--) {
    if (buf.readUInt32LE(i) === SIG_EOCD) return i;
  }
  throw new Error('Not a ZIP: EOCD signature not found');
}

/** unzip(bytes) → Map<name, Buffer> of stored entry contents. */
export function unzip(bytes) {
  const buf = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  const eocd = findEOCD(buf);
  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16); // central dir offset

  const out = new Map();
  for (let n = 0; n < count; n++) {
    if (buf.readUInt32LE(off) !== SIG_CEN) throw new Error('Bad central dir entry');
    const method = buf.readUInt16LE(off + 10);
    const compSize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const localOff = buf.readUInt32LE(off + 42);
    const name = buf.toString('utf8', off + 46, off + 46 + nameLen);
    off += 46 + nameLen + extraLen + commentLen;

    // Read the local header to find where the data actually starts (its
    // name/extra lengths can differ from the central dir's).
    if (buf.readUInt32LE(localOff) !== SIG_LOC) throw new Error('Bad local header');
    const lNameLen = buf.readUInt16LE(localOff + 26);
    const lExtraLen = buf.readUInt16LE(localOff + 28);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    const raw = buf.subarray(dataStart, dataStart + compSize);

    let content;
    if (method === 0) content = Buffer.from(raw);            // stored
    else if (method === 8) content = zlib.inflateRawSync(raw); // deflate
    else throw new Error(`Unsupported ZIP method ${method} for ${name}`);

    if (!name.endsWith('/')) out.set(name, content);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Volatile-content stripping + light XML canonicalization
// ---------------------------------------------------------------------------

// Whole parts whose content is inherently non-deterministic / metadata-only.
const VOLATILE_PARTS = [
  /^docProps\/core\.xml$/,   // dcterms:created/modified, revision
  /^docProps\/app\.xml$/,    // TotalTime, AppVersion, etc.
  /^docProps\/custom\.xml$/,
  /^xl\/calcChain\.xml$/,    // recomputed on open, order churns
];

// Attribute patterns stripped from every XML part.
const VOLATILE_ATTRS = [
  /\s+w:rsid[A-Za-z]*="[^"]*"/g,         // revision save ids
  /\s+w14:paraId="[^"]*"/g,
  /\s+w14:textId="[^"]*"/g,
  /\s+w15:[A-Za-z]+="[^"]*"/g,
  // Brace-delimited GUIDs (e.g. xlsx x14:cfRule id="{…}") — x2t regenerates
  // these randomly on every conversion. Legit ids rarely use this format.
  /\s+id="\{[0-9A-Fa-f-]{36}\}"/g,
];

// Binary/opaque parts whose content embeds the volatile ids above and so
// can't be hashed stably. Compared as presence-only (key kept, value fixed).
const OPAQUE_PARTS = [
  /\.vml$/,                              // VML drawings embed cfRule guids etc.
];

// Whole elements stripped from every XML part.
const VOLATILE_ELEMENTS = [
  /<w:rsids>[\s\S]*?<\/w:rsids>/g,       // the rsid table in settings.xml
];

function isXmlPart(name) {
  return /\.(xml|rels)$/i.test(name);
}

/** Sort the attributes inside each start/empty tag for stable comparison. */
function sortAttributes(xml) {
  return xml.replace(/<([A-Za-z_][\w:.-]*)((?:\s+[^<>]*?)?)(\/?)>/g, (m, tag, attrs, selfClose) => {
    if (!attrs || !attrs.trim()) return `<${tag}${selfClose}>`;
    const found = [...attrs.matchAll(/([\w:.-]+)\s*=\s*"([^"]*)"/g)];
    if (found.length === 0) return m;
    const sorted = found
      .map(([, k, v]) => `${k}="${v}"`)
      .sort()
      .join(' ');
    return `<${tag} ${sorted}${selfClose}>`;
  });
}

function normalizeXml(text) {
  let s = text;
  for (const re of VOLATILE_ELEMENTS) s = s.replace(re, '');
  for (const re of VOLATILE_ATTRS) s = s.replace(re, '');
  s = s.replace(/<\?xml[^>]*\?>/, '');   // drop XML decl (encoding/standalone vary)
  s = s.replace(/>\s+</g, '><');          // collapse inter-tag whitespace
  s = sortAttributes(s);
  return s.trim();
}

function sha256(buf) {
  return crypto.createHash('sha256').update(buf).digest('hex');
}

/**
 * normalizeOoxml(bytes) → { [partName]: normalizedRepresentation }
 * XML parts → canonical XML string; binary parts → "sha256:<hash>".
 * Volatile whole-parts are dropped entirely.
 */
export function normalizeOoxml(bytes) {
  const entries = unzip(bytes);
  const result = {};
  for (const [name, content] of entries) {
    if (VOLATILE_PARTS.some((re) => re.test(name))) continue;
    if (OPAQUE_PARTS.some((re) => re.test(name))) { result[name] = 'opaque'; continue; }
    result[name] = isXmlPart(name)
      ? normalizeXml(content.toString('utf8'))
      : 'sha256:' + sha256(content);
  }
  return result;
}

/**
 * diffOoxml(a, b) → array of human-readable difference descriptions
 * (empty array === content-equivalent). Compares normalized representations.
 */
export function diffOoxml(aBytes, bBytes) {
  const a = normalizeOoxml(aBytes);
  const b = normalizeOoxml(bBytes);
  const names = new Set([...Object.keys(a), ...Object.keys(b)]);
  const diffs = [];
  for (const name of [...names].sort()) {
    if (!(name in a)) { diffs.push(`+ only in B: ${name}`); continue; }
    if (!(name in b)) { diffs.push(`- only in A: ${name}`); continue; }
    if (a[name] !== b[name]) {
      const where = firstDiffIndex(a[name], b[name]);
      diffs.push(
        `~ differs: ${name} (at char ${where})\n` +
        `    A: …${snippet(a[name], where)}…\n` +
        `    B: …${snippet(b[name], where)}…`,
      );
    }
  }
  return diffs;
}

function firstDiffIndex(a, b) {
  const n = Math.min(a.length, b.length);
  for (let i = 0; i < n; i++) if (a[i] !== b[i]) return i;
  return n;
}
function snippet(s, at) {
  return s.slice(Math.max(0, at - 40), at + 40);
}
