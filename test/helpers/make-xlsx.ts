// A tiny ZIP writer for tests: each entry stored uncompressed (method 0). Little-endian, no data
// descriptors. Enough to craft arbitrary .xlsx fixtures (1904, missing parts, custom XML).
function crc32(bytes: Uint8Array): number {
  let c = ~0;
  for (const byte of bytes) {
    c ^= byte!;
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
  }
  return ~c >>> 0;
}

export function makeXlsx(entries: Record<string, string>): Uint8Array {
  const enc = new TextEncoder();
  const files = Object.entries(entries).map(([name, content]) => ({
    data: enc.encode(content),
    name: enc.encode(name),
  }));
  const locals: Uint8Array[] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  for (const f of files) {
    const crc = crc32(f.data);
    const lh = new Uint8Array(30 + f.name.length);
    const lv = new DataView(lh.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true); // version needed
    lv.setUint16(8, 0, true); // method 0 (stored)
    lv.setUint32(14, crc, true);
    lv.setUint32(18, f.data.length, true); // compressed size
    lv.setUint32(22, f.data.length, true); // uncompressed size
    lv.setUint16(26, f.name.length, true);
    lh.set(f.name, 30);
    locals.push(lh, f.data);

    const ch = new Uint8Array(46 + f.name.length);
    const cv = new DataView(ch.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(10, 0, true); // method 0
    cv.setUint32(16, crc, true);
    cv.setUint32(20, f.data.length, true);
    cv.setUint32(24, f.data.length, true);
    cv.setUint16(28, f.name.length, true);
    cv.setUint32(42, offset, true); // local header offset
    ch.set(f.name, 46);
    centrals.push(ch);

    offset += lh.length + f.data.length;
  }
  const cdSize = centrals.reduce((s, c) => s + c.length, 0);
  const cdOffset = offset;
  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, cdOffset, true);
  const parts = [...locals, ...centrals, eocd];
  const total = parts.reduce((s, p) => s + p.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}
