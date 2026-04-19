/**
 * xlsxPatch.js
 *
 * Patches an XLSX file at the ZIP + XML level.
 * Only the target cells are rewritten; every other byte in the archive is
 * preserved unchanged, so all original cell styles, borders, merged cells,
 * conditional formatting, column widths, row heights, table definitions,
 * defined names, and sheet-level settings are kept exactly as they were.
 *
 * Requires the Compression Streams API (Chrome 80+, Firefox 113+, Safari 16.4+).
 * Call `isPatchSupported()` to check before use.
 */

// ── feature detection ────────────────────────────────────────

export function isPatchSupported() {
  return (
    typeof CompressionStream !== 'undefined' &&
    typeof DecompressionStream !== 'undefined' &&
    typeof DOMParser !== 'undefined' &&
    typeof XMLSerializer !== 'undefined'
  );
}

// ── ZIP constants & CRC-32 ───────────────────────────────────

const SIG_LOCAL   = 0x04034b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_EOCD    = 0x06054b50;
const METHOD_STORE   = 0;
const METHOD_DEFLATE = 8;

const CRC32_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c;
  }
  return t;
})();

function crc32(data) {
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = CRC32_TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

// ── byte helpers ─────────────────────────────────────────────

function concatU8(arrays) {
  const total = arrays.reduce((n, a) => n + a.length, 0);
  const out = new Uint8Array(total);
  let pos = 0;
  for (const a of arrays) { out.set(a, pos); pos += a.length; }
  return out;
}

async function drainStream(readable) {
  const reader = readable.getReader();
  const chunks = [];
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (value) chunks.push(value);
  }
  return concatU8(chunks);
}

async function zlibInflate(data) {
  const ds = new DecompressionStream('deflate-raw');
  const w = ds.writable.getWriter();
  w.write(data); w.close();
  return drainStream(ds.readable);
}

async function zlibDeflate(data) {
  const cs = new CompressionStream('deflate-raw');
  const w = cs.writable.getWriter();
  w.write(data); w.close();
  return drainStream(cs.readable);
}

// ── ZIP reader ───────────────────────────────────────────────

function readZip(buffer) {
  const view = new DataView(buffer);
  const len  = buffer.byteLength;

  // Locate End of Central Directory (search backwards, allow comment up to 64 KB)
  let eocdPos = -1;
  for (let i = len - 22; i >= Math.max(0, len - 65558); i--) {
    if (view.getUint32(i, true) === SIG_EOCD) { eocdPos = i; break; }
  }
  if (eocdPos < 0) throw new Error('Not a valid ZIP/XLSX file');

  const cdCount  = view.getUint16(eocdPos + 8,  true);
  const cdOffset = view.getUint32(eocdPos + 16, true);

  const dec = new TextDecoder('utf-8');
  const entries = [];
  let pos = cdOffset;

  for (let i = 0; i < cdCount; i++) {
    if (view.getUint32(pos, true) !== SIG_CENTRAL) break;

    const compression = view.getUint16(pos + 10, true);
    const modTime     = view.getUint16(pos + 12, true);
    const modDate     = view.getUint16(pos + 14, true);
    const fileCrc     = view.getUint32(pos + 16, true);
    const compSize    = view.getUint32(pos + 20, true);
    const uncompSize  = view.getUint32(pos + 24, true);
    const nameLen     = view.getUint16(pos + 28, true);
    const extraLen    = view.getUint16(pos + 30, true);
    const commentLen  = view.getUint16(pos + 32, true);
    const localOff    = view.getUint32(pos + 42, true);
    const name        = dec.decode(new Uint8Array(buffer, pos + 46, nameLen));

    // Skip to actual data past local file header
    const lv        = new DataView(buffer, localOff);
    const lNameLen  = lv.getUint16(26, true);
    const lExtraLen = lv.getUint16(28, true);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;

    entries.push({
      name, compression, modTime, modDate,
      crc: fileCrc, compSize, uncompSize,
      data: new Uint8Array(buffer, dataStart, compSize),
    });

    pos += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

// ── ZIP writer ───────────────────────────────────────────────

function writeZip(entries) {
  const enc = new TextEncoder();
  const localParts  = [];
  const centralParts = [];
  const offsets = [];
  let offset = 0;

  for (const e of entries) {
    const nameBuf = enc.encode(e.name);
    offsets.push(offset);

    const lh = new Uint8Array(30 + nameBuf.length);
    const lv = new DataView(lh.buffer);
    lv.setUint32(0,  SIG_LOCAL,     true);
    lv.setUint16(4,  20,            true);  // version needed
    lv.setUint16(6,  0,             true);  // flags
    lv.setUint16(8,  e.compression, true);
    lv.setUint16(10, e.modTime || 0, true);
    lv.setUint16(12, e.modDate || 0, true);
    lv.setUint32(14, e.crc,          true);
    lv.setUint32(18, e.data.length,  true);
    lv.setUint32(22, e.uncompSize,   true);
    lv.setUint16(26, nameBuf.length, true);
    lv.setUint16(28, 0,              true);
    lh.set(nameBuf, 30);

    localParts.push(lh, e.data);
    offset += lh.length + e.data.length;
  }

  const cdStart = offset;

  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    const nameBuf = enc.encode(e.name);

    const cd = new Uint8Array(46 + nameBuf.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0,  SIG_CENTRAL,    true);
    cv.setUint16(4,  20,             true);
    cv.setUint16(6,  20,             true);
    cv.setUint16(8,  0,              true);
    cv.setUint16(10, e.compression,  true);
    cv.setUint16(12, e.modTime || 0, true);
    cv.setUint16(14, e.modDate || 0, true);
    cv.setUint32(16, e.crc,          true);
    cv.setUint32(20, e.data.length,  true);
    cv.setUint32(24, e.uncompSize,   true);
    cv.setUint16(28, nameBuf.length, true);
    cv.setUint16(30, 0,              true);
    cv.setUint16(32, 0,              true);
    cv.setUint16(34, 0,              true);
    cv.setUint16(36, 0,              true);
    cv.setUint32(38, 0,              true);
    cv.setUint32(42, offsets[i],     true);
    cd.set(nameBuf, 46);
    centralParts.push(cd);
  }

  const cdBytes = concatU8(centralParts);
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0,  SIG_EOCD,         true);
  ev.setUint16(4,  0,                true);
  ev.setUint16(6,  0,                true);
  ev.setUint16(8,  entries.length,   true);
  ev.setUint16(10, entries.length,   true);
  ev.setUint32(12, cdBytes.length,   true);
  ev.setUint32(16, cdStart,          true);
  ev.setUint16(20, 0,                true);

  return concatU8([...localParts, cdBytes, eocd]).buffer;
}

// ── decompress entry ─────────────────────────────────────────

async function entryBytes(entry) {
  if (entry.compression === METHOD_DEFLATE) return zlibInflate(entry.data);
  return entry.data; // METHOD_STORE
}

// ── cell address helpers ──────────────────────────────────────

function colIdxToName(idx) {
  let name = '';
  let n = idx + 1;
  while (n > 0) {
    name = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function colNameToIdx(name) {
  let idx = 0;
  for (const ch of name.toUpperCase()) idx = idx * 26 + (ch.charCodeAt(0) - 64);
  return idx - 1;
}

// Build a cell address like "C5" from 0-based row/col indices
function cellAddress(rowIdx, colIdx) {
  return colIdxToName(colIdx) + (rowIdx + 1);
}

// ── XML helpers ───────────────────────────────────────────────

const PARSER     = new DOMParser();
const SERIALIZER = new XMLSerializer();

function parseXml(str) {
  return PARSER.parseFromString(str, 'text/xml');
}

function serializeXml(doc) {
  let xml = SERIALIZER.serializeToString(doc);
  // Ensure UTF-8 declaration is present
  if (!xml.startsWith('<?xml')) {
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
  }
  return xml;
}

function byTag(parent, tag) {
  // getElementsByTagName works across any namespace
  return Array.from(parent.getElementsByTagName(tag));
}

function makeEl(doc, tag) {
  const ns = doc.documentElement?.namespaceURI;
  return ns ? doc.createElementNS(ns, tag) : doc.createElement(tag);
}

// ── shared string table ───────────────────────────────────────

function buildSst(sstDoc) {
  const strings = [];
  const idx = new Map();
  if (!sstDoc) return { strings, idx, sstDoc: null };
  const sis = byTag(sstDoc.documentElement, 'si');
  for (const si of sis) {
    const t = si.getElementsByTagName('t')[0];
    const s = t ? t.textContent : '';
    idx.set(s, strings.length);
    strings.push(s);
  }
  return { strings, idx, sstDoc };
}

function sstGetOrAdd({ strings, idx, sstDoc }, value) {
  if (idx.has(value)) return idx.get(value);
  const i = strings.length;
  strings.push(value);
  idx.set(value, i);
  if (sstDoc) {
    const si = makeEl(sstDoc, 'si');
    const t  = makeEl(sstDoc, 't');
    if (value.startsWith(' ') || value.endsWith(' ')) t.setAttribute('xml:space', 'preserve');
    t.textContent = value;
    si.appendChild(t);
    sstDoc.documentElement.appendChild(si);
  }
  return i;
}

function sstUpdateCounts({ strings, sstDoc }) {
  if (!sstDoc) return;
  const root = sstDoc.documentElement;
  root.setAttribute('count',       String(strings.length));
  root.setAttribute('uniqueCount', String(strings.length));
}

// ── worksheet patching ────────────────────────────────────────

function patchWorksheet(wsDoc, sst, fillMap, headers) {
  const sheetData = wsDoc.getElementsByTagName('sheetData')[0];
  if (!sheetData) return;

  // Index existing rows by their 1-based r attribute
  const rowMap = new Map();
  for (const rowEl of byTag(sheetData, 'row')) {
    rowMap.set(parseInt(rowEl.getAttribute('r') || '0', 10), rowEl);
  }

  for (const [rIdxStr, fills] of Object.entries(fillMap)) {
    const rowIdx = Number(rIdxStr);
    const rowNum = rowIdx + 1; // 1-based

    let rowEl = rowMap.get(rowNum);
    if (!rowEl) {
      rowEl = makeEl(wsDoc, 'row');
      rowEl.setAttribute('r', String(rowNum));
      // Insert in sorted order
      let inserted = false;
      for (const [r, el] of rowMap.entries()) {
        if (r > rowNum) { sheetData.insertBefore(rowEl, el); inserted = true; break; }
      }
      if (!inserted) sheetData.appendChild(rowEl);
      rowMap.set(rowNum, rowEl);
    }

    for (const [colName, rawValue] of Object.entries(fills)) {
      const colIdx = headers.indexOf(colName);
      if (colIdx === -1) continue;

      const ref    = cellAddress(rowIdx, colIdx);
      const strIdx = sstGetOrAdd(sst, String(rawValue));

      // Find existing cell with matching r attribute
      let cellEl = null;
      for (const c of byTag(rowEl, 'c')) {
        if (c.getAttribute('r') === ref) { cellEl = c; break; }
      }

      if (cellEl) {
        // Preserve the style index (s=) — only replace value + type
        const styleAttr = cellEl.getAttribute('s');
        while (cellEl.firstChild) cellEl.removeChild(cellEl.firstChild);
        cellEl.setAttribute('t', 's');
        if (styleAttr !== null) cellEl.setAttribute('s', styleAttr);
        const v = makeEl(wsDoc, 'v');
        v.textContent = String(strIdx);
        cellEl.appendChild(v);
      } else {
        // New cell — insert in column order
        const newCell = makeEl(wsDoc, 'c');
        newCell.setAttribute('r', ref);
        newCell.setAttribute('t', 's');
        const v = makeEl(wsDoc, 'v');
        v.textContent = String(strIdx);
        newCell.appendChild(v);

        let insertBefore = null;
        for (const c of byTag(rowEl, 'c')) {
          const cRef = c.getAttribute('r') || '';
          const cCol = colNameToIdx(cRef.match(/[A-Z]+/)?.[0] || 'A');
          if (cCol > colIdx) { insertBefore = c; break; }
        }
        rowEl.insertBefore(newCell, insertBefore);
      }
    }
  }
}

// ── resolve workbook rId → worksheet file path ────────────────

function decodeText(bytes) { return new TextDecoder('utf-8').decode(bytes); }

function xmlAttrValue(xml, tagHint, attrName) {
  // Quick regex to grab an attribute value — only used on workbook/rels XML
  const re = new RegExp(`\\b${attrName}="([^"]*)"`, 'g');
  const tagRe = new RegExp(`<${tagHint}\\b([^>]*)>`, 'gi');
  let m;
  while ((m = tagRe.exec(xml)) !== null) {
    const attrs = m[1];
    const am = re.exec(attrs);
    re.lastIndex = 0;
    if (am) return am[1];
  }
  return '';
}

async function resolveSheetPath(entries, sheetName) {
  const entryMap = new Map(entries.map(e => [e.name, e]));

  // Parse workbook.xml
  const wbEntry = entryMap.get('xl/workbook.xml');
  if (!wbEntry) throw new Error('xl/workbook.xml missing');
  const wbXml = decodeText(await entryBytes(wbEntry));

  // Find the sheet element matching the name and grab its r:id
  const sheetTagRe = /<sheet\b([^>]*)\/?>|<sheet\b([^>]*)>/gi;
  let rId = '';
  let m;
  while ((m = sheetTagRe.exec(wbXml)) !== null) {
    const attrs = m[1] || m[2] || '';
    const nameMatch = attrs.match(/\bname="([^"]*)"/);
    if (nameMatch && nameMatch[1] === sheetName) {
      const ridMatch = attrs.match(/(?:r:id|rId)="([^"]*)"/);
      if (ridMatch) { rId = ridMatch[1]; break; }
    }
  }
  if (!rId) throw new Error(`Sheet "${sheetName}" not found`);

  // Parse workbook.xml.rels
  const relsEntry = entryMap.get('xl/_rels/workbook.xml.rels');
  if (!relsEntry) throw new Error('xl/_rels/workbook.xml.rels missing');
  const relsXml = decodeText(await entryBytes(relsEntry));

  const relRe = new RegExp(`<Relationship\\b([^>]*)Id="${rId}"([^>]*)/>|<Relationship\\b([^>]*)Id="${rId}"([^>]*)>`, 'i');
  const relMatch = relsXml.match(relRe);
  if (!relMatch) throw new Error(`Relationship ${rId} not found`);
  const allAttrs = (relMatch[1] || '') + (relMatch[2] || '') + (relMatch[3] || '') + (relMatch[4] || '');
  const targetMatch = allAttrs.match(/\bTarget="([^"]*)"/i);
  if (!targetMatch) throw new Error('Target attribute not found in relationship');

  let path = targetMatch[1];
  if (!path.startsWith('xl/')) path = 'xl/' + path;
  return path;
}

// ── handle missing sharedStrings.xml ─────────────────────────

function buildEmptySst() {
  return parseXml(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>'
  );
}

async function ensureSstInContentTypes(entries) {
  // If sharedStrings.xml is already referenced, nothing to do
  const idx = entries.findIndex(e => e.name === 'xl/sharedStrings.xml');
  if (idx !== -1) return entries;

  // Add sharedStrings.xml to [Content_Types].xml
  const ctIdx = entries.findIndex(e => e.name === '[Content_Types].xml');
  if (ctIdx === -1) return entries; // Can't fix — just proceed

  const ctBytes = await entryBytes(entries[ctIdx]);
  let ctXml = decodeText(ctBytes);
  const override = '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
  ctXml = ctXml.replace('</Types>', override + '</Types>');
  const enc = new TextEncoder();
  const newCtBytes = enc.encode(ctXml);
  const newCtData  = await zlibDeflate(newCtBytes);
  const newEntries = [...entries];
  newEntries[ctIdx] = {
    ...entries[ctIdx],
    data: newCtData,
    uncompSize: newCtBytes.length,
    crc: crc32(newCtBytes),
    compression: METHOD_DEFLATE,
  };
  return newEntries;
}

async function ensureSstInRels(entries) {
  const idx = entries.findIndex(e => e.name === 'xl/sharedStrings.xml');
  if (idx !== -1) return entries; // already there

  const relsIdx = entries.findIndex(e => e.name === 'xl/_rels/workbook.xml.rels');
  if (relsIdx === -1) return entries;

  const relsBytes = await entryBytes(entries[relsIdx]);
  let relsXml = decodeText(relsBytes);
  const rel = '<Relationship Id="rIdSS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>';
  relsXml = relsXml.replace('</Relationships>', rel + '</Relationships>');
  const enc = new TextEncoder();
  const newBytes = enc.encode(relsXml);
  const newData  = await zlibDeflate(newBytes);
  const newEntries = [...entries];
  newEntries[relsIdx] = {
    ...entries[relsIdx],
    data: newData,
    uncompSize: newBytes.length,
    crc: crc32(newBytes),
    compression: METHOD_DEFLATE,
  };
  return newEntries;
}

// ── public API ────────────────────────────────────────────────

/**
 * Patch an XLSX file in-place.
 *
 * @param {ArrayBuffer} fileBuffer  Original XLSX bytes
 * @param {string}      sheetName   Sheet name to modify
 * @param {Object}      fillMap     { rowIndex0Based: { colName: value } }
 * @param {string[]}    headers     Column names in order (0-based index = column)
 * @returns {Promise<ArrayBuffer>}  Modified XLSX bytes
 */
export async function patchXlsxCells(fileBuffer, sheetName, fillMap, headers) {
  // 1. Parse ZIP
  let entries = readZip(fileBuffer);

  // 2. Resolve worksheet file path
  const sheetPath = await resolveSheetPath(entries, sheetName);
  const entryMap  = () => new Map(entries.map(e => [e.name, e]));

  // 3. Load shared strings (or create empty)
  const em = entryMap();
  const sstEntry = em.get('xl/sharedStrings.xml');
  let sstDoc;
  if (sstEntry) {
    sstDoc = parseXml(decodeText(await entryBytes(sstEntry)));
  } else {
    sstDoc = buildEmptySst();
    entries = await ensureSstInContentTypes(entries);
    entries = await ensureSstInRels(entries);
  }
  const sst = buildSst(sstDoc);

  // 4. Load and patch worksheet XML
  const wsEntry = entryMap().get(sheetPath);
  if (!wsEntry) throw new Error(`Worksheet "${sheetPath}" not found in ZIP`);
  const wsDoc = parseXml(decodeText(await entryBytes(wsEntry)));
  patchWorksheet(wsDoc, sst, fillMap, headers);

  // 5. Finalize SST counts
  sstUpdateCounts(sst);

  // 6. Serialize modified XMLs
  const enc        = new TextEncoder();
  const newWsText  = serializeXml(wsDoc);
  const newSstText = serializeXml(sstDoc);
  const newWsU8    = enc.encode(newWsText);
  const newSstU8   = enc.encode(newSstText);

  // 7. Re-compress
  const newWsComp  = await zlibDeflate(newWsU8);
  const newSstComp = await zlibDeflate(newSstU8);

  // 8. Rebuild entry list
  const newEntries = entries.map(e => {
    if (e.name === sheetPath) {
      return { ...e, data: newWsComp, uncompSize: newWsU8.length, crc: crc32(newWsU8), compression: METHOD_DEFLATE };
    }
    if (e.name === 'xl/sharedStrings.xml') {
      return { ...e, data: newSstComp, uncompSize: newSstU8.length, crc: crc32(newSstU8), compression: METHOD_DEFLATE };
    }
    return e;
  });

  // If SST was newly created, append it
  if (!sstEntry) {
    newEntries.push({
      name: 'xl/sharedStrings.xml',
      compression: METHOD_DEFLATE,
      modTime: 0, modDate: 0,
      data: newSstComp,
      uncompSize: newSstU8.length,
      crc: crc32(newSstU8),
    });
  }

  // 9. Write new ZIP
  return writeZip(newEntries);
}
