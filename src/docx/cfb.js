// Compound files (MS-CFB), the container a password-encrypted .docx is kept in: Word puts the
// encrypted package (EncryptedPackage), how it was encrypted (EncryptionInfo) and the data spaces
// that say so (\x06DataSpaces) in an OLE compound file instead of a ZIP. Only what that needs:
// reading streams by path, and writing a compound file of streams and storages (version 3, 512-byte
// sectors, small streams in the mini stream, as Word writes these files).

const FREE = 0xffffffff;
const END = 0xfffffffe;
const FAT_SECT = 0xfffffffd;
const DIFAT_SECT = 0xfffffffc;
const NONE = 0xffffffff;
const MAGIC = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
const MINI_CUTOFF = 4096;

const damaged = () => new Error('The file’s container is damaged: its parts can’t be read');

/**
 * Reads a compound file: { streams: Map(path -> Uint8Array) }, paths joined with '/' from the root
 * (e.g. 'EncryptionInfo', '\x06DataSpaces/Version').
 * @param {Uint8Array} bytes
 */
export function readCfb(bytes) {
  if (!MAGIC.every((b, i) => bytes[i] === b)) throw damaged();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u32 = (at) => view.getUint32(at, true);
  const sectorSize = 1 << view.getUint16(30, true);
  const miniSize = 1 << view.getUint16(32, true);
  if (sectorSize !== 512 && sectorSize !== 4096) throw damaged();
  const offset = (s) => (s + 1) * sectorSize;
  const sector = (s) => {
    if (offset(s) + sectorSize > bytes.length + sectorSize) throw damaged();
    return bytes.subarray(offset(s), Math.min(offset(s) + sectorSize, bytes.length));
  };
  // The FAT's sectors: 109 in the header, the rest in the DIFAT chain.
  const fatSectors = [];
  for (let i = 0; i < 109; i++) {
    const s = u32(76 + i * 4);
    if (s < DIFAT_SECT) fatSectors.push(s);
  }
  const perSector = sectorSize / 4;
  let difat = u32(68);
  for (let guard = 0; difat < DIFAT_SECT && guard < 1e6; guard++) {
    for (let i = 0; i < perSector - 1; i++) {
      const s = u32(offset(difat) + i * 4);
      if (s < DIFAT_SECT) fatSectors.push(s);
    }
    difat = u32(offset(difat) + (perSector - 1) * 4);
  }
  const fat = new Uint32Array(fatSectors.length * perSector);
  fatSectors.forEach((s, i) => {
    for (let k = 0; k < perSector; k++) fat[i * perSector + k] = u32(offset(s) + k * 4);
  });
  const chain = (start, table) => {
    const out = [];
    for (let s = start; s < DIFAT_SECT; s = table[s]) {
      if (s >= table.length || out.length > table.length) throw damaged();
      out.push(s);
    }
    return out;
  };
  const readChain = (start) => {
    const sectors = chain(start, fat);
    const out = new Uint8Array(sectors.length * sectorSize);
    sectors.forEach((s, i) => out.set(sector(s), i * sectorSize));
    return out;
  };
  const dir = readChain(u32(48));
  const dview = new DataView(dir.buffer);
  const entries = [];
  for (let at = 0; at + 128 <= dir.length; at += 128) {
    const nameLen = dview.getUint16(at + 64, true);
    let name = '';
    for (let i = 0; i + 2 < nameLen && i < 64; i += 2) name += String.fromCharCode(dview.getUint16(at + i, true));
    entries.push({
      name,
      type: dir[at + 66],
      left: dview.getUint32(at + 68, true),
      right: dview.getUint32(at + 72, true),
      child: dview.getUint32(at + 76, true),
      start: dview.getUint32(at + 116, true),
      size: dview.getUint32(at + 120, true) + (sectorSize === 4096 ? dview.getUint32(at + 124, true) * 2 ** 32 : 0),
    });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw damaged();
  const miniFatBytes = u32(64) ? readChain(u32(60)) : new Uint8Array(0);
  const miniFat = new Uint32Array(miniFatBytes.buffer, 0, miniFatBytes.length / 4);
  const miniStream = root.start < DIFAT_SECT ? readChain(root.start) : new Uint8Array(0);
  const streamOf = (e) => {
    if (e.size < MINI_CUTOFF) {
      const out = new Uint8Array(e.size);
      let at = 0;
      for (const s of chain(e.start, miniFat)) {
        const piece = miniStream.subarray(s * miniSize, s * miniSize + Math.min(miniSize, e.size - at));
        out.set(piece, at);
        at += piece.length;
        if (at >= e.size) break;
      }
      return out;
    }
    return readChain(e.start).subarray(0, e.size);
  };
  const streams = new Map();
  const seen = new Set();
  const walk = (id, prefix) => {
    if (id === NONE || id >= entries.length || seen.has(id)) return;
    seen.add(id);
    const e = entries[id];
    walk(e.left, prefix);
    walk(e.right, prefix);
    if (e.type === 2) streams.set(prefix + e.name, streamOf(e));
    else if (e.type === 1) walk(e.child, `${prefix}${e.name}/`);
  };
  walk(root.child, '');
  return { streams };
}

/** Sibling order in a storage, as MS-CFB has it: shorter names first, then by upper case. */
function compareNames(a, b) {
  if (a.length !== b.length) return a.length - b.length;
  const x = a.toUpperCase();
  const y = b.toUpperCase();
  return x < y ? -1 : x > y ? 1 : 0;
}

/**
 * Writes a compound file holding `streams`: Map(path -> Uint8Array), storages named by the paths
 * ('\x06DataSpaces/Version' is the stream Version in the storage \x06DataSpaces).
 * @returns {Uint8Array}
 */
export function writeCfb(streams) {
  // The tree of storages and streams.
  const root = { name: 'Root Entry', type: 5, children: new Map() };
  for (const [path, data] of streams) {
    const parts = path.split('/');
    let node = root;
    for (const part of parts.slice(0, -1)) {
      if (!node.children.has(part)) node.children.set(part, { name: part, type: 1, children: new Map() });
      node = node.children.get(part);
    }
    node.children.set(parts.at(-1), { name: parts.at(-1), type: 2, data });
  }
  const entries = [];
  const add = (node) => {
    node.id = entries.length;
    entries.push(node);
    node.left = NONE;
    node.right = NONE;
    node.child = NONE;
    node.color = 1; // black
    if (!node.children) return;
    const kids = [...node.children.values()].sort((a, b) => compareNames(a.name, b.name));
    for (const k of kids) add(k);
    // The siblings as a balanced red-black tree: balanced by halves, the deepest level red when the
    // tree isn't full, so every path has as many black nodes.
    const levels = [];
    const build = (list, depth) => {
      if (!list.length) return NONE;
      const mid = list.length >> 1;
      const n = list[mid];
      (levels[depth] ||= []).push(n);
      n.left = build(list.slice(0, mid), depth + 1);
      n.right = build(list.slice(mid + 1), depth + 1);
      return n.id;
    };
    node.child = build(kids, 0);
    const deepest = levels.at(-1);
    if (levels.length > 1 && deepest.length < 2 ** (levels.length - 1)) for (const n of deepest) n.color = 0;
  };
  add(root);

  // Small streams in the mini stream (64-byte sectors), the rest in sectors of their own.
  const miniFat = [];
  const miniParts = [];
  let miniLength = 0;
  const big = [];
  for (const e of entries) {
    if (e.type !== 2) continue;
    e.size = e.data.length;
    if (!e.size) e.start = END;
    else if (e.size < MINI_CUTOFF) {
      const count = Math.ceil(e.size / 64);
      e.start = miniFat.length;
      for (let i = 0; i < count; i++) miniFat.push(i === count - 1 ? END : e.start + i + 1);
      miniParts.push([miniLength, e.data]);
      miniLength += count * 64;
    } else big.push(e);
  }
  const miniStream = new Uint8Array(miniLength);
  for (const [at, data] of miniParts) miniStream.set(data, at);
  const sectorsFor = (n) => Math.ceil(n / 512);
  const miniFatBytes = new Uint8Array(sectorsFor(miniFat.length * 4) * 512).fill(0xff);
  new Uint32Array(miniFatBytes.buffer).set(miniFat);
  const dirBytes = new Uint8Array(sectorsFor(entries.length * 128) * 512);

  // Sectors in order: big streams, the mini stream, the mini FAT, the directory; then the FAT and DIFAT.
  const blocks = [];
  let next = 0;
  const place = (data) => {
    const count = sectorsFor(data.length);
    const start = count ? next : END;
    blocks.push({ start, count, data });
    next += count;
    return start;
  };
  for (const e of big) e.start = place(e.data);
  root.start = place(miniStream);
  root.size = miniLength;
  const miniFatStart = place(miniFatBytes);
  const dirStart = place(dirBytes);
  let fatCount = 1;
  let difatCount = 0;
  for (;;) {
    difatCount = Math.max(0, Math.ceil((fatCount - 109) / 127));
    if (fatCount * 128 >= next + fatCount + difatCount) break;
    fatCount++;
  }
  const fatStart = next;
  const difatStart = fatStart + fatCount;
  const total = difatStart + difatCount;
  const fat = new Uint32Array(fatCount * 128).fill(FREE);
  for (const b of blocks) for (let i = 0; i < b.count; i++) fat[b.start + i] = i === b.count - 1 ? END : b.start + i + 1;
  for (let i = 0; i < fatCount; i++) fat[fatStart + i] = FAT_SECT;
  for (let i = 0; i < difatCount; i++) fat[difatStart + i] = DIFAT_SECT;

  // The directory.
  const dv = new DataView(dirBytes.buffer);
  for (let i = entries.length; i < dirBytes.length / 128; i++) {
    dv.setUint32(i * 128 + 68, NONE, true);
    dv.setUint32(i * 128 + 72, NONE, true);
    dv.setUint32(i * 128 + 76, NONE, true);
  }
  entries.forEach((e, i) => {
    const at = i * 128;
    for (let k = 0; k < e.name.length && k < 31; k++) dv.setUint16(at + k * 2, e.name.charCodeAt(k), true);
    dv.setUint16(at + 64, (Math.min(e.name.length, 31) + 1) * 2, true);
    dirBytes[at + 66] = e.type;
    dirBytes[at + 67] = e.color;
    dv.setUint32(at + 68, e.left, true);
    dv.setUint32(at + 72, e.right, true);
    dv.setUint32(at + 76, e.child, true);
    dv.setUint32(at + 116, e.type === 1 ? 0 : e.start ?? END, true);
    dv.setUint32(at + 120, e.type === 1 ? 0 : e.size || 0, true);
  });

  const out = new Uint8Array(512 + total * 512);
  const ov = new DataView(out.buffer);
  out.set(MAGIC, 0);
  ov.setUint16(24, 0x3e, true); // minor version
  ov.setUint16(26, 3, true); // major version 3
  ov.setUint16(28, 0xfffe, true); // little-endian
  ov.setUint16(30, 9, true); // 512-byte sectors
  ov.setUint16(32, 6, true); // 64-byte mini sectors
  ov.setUint32(44, fatCount, true);
  ov.setUint32(48, dirStart, true);
  ov.setUint32(56, MINI_CUTOFF, true);
  ov.setUint32(60, miniFat.length ? miniFatStart : END, true);
  ov.setUint32(64, miniFatBytes.length / 512, true);
  ov.setUint32(68, difatCount ? difatStart : END, true);
  ov.setUint32(72, difatCount, true);
  for (let i = 0; i < 109; i++) ov.setUint32(76 + i * 4, i < fatCount ? fatStart + i : FREE, true);
  for (const b of blocks) if (b.count) out.set(b.data, 512 + b.start * 512);
  out.set(new Uint8Array(fat.buffer), 512 + fatStart * 512);
  for (let d = 0; d < difatCount; d++) {
    const at = 512 + (difatStart + d) * 512;
    for (let i = 0; i < 127; i++) {
      const f = 109 + d * 127 + i;
      ov.setUint32(at + i * 4, f < fatCount ? fatStart + f : FREE, true);
    }
    ov.setUint32(at + 127 * 4, d === difatCount - 1 ? END : difatStart + d + 1, true);
  }
  return out;
}
