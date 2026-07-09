import { XlsxError } from "./limits";

const EOCD_SIG = 0x06054b50;
const CD_SIG = 0x02014b50;

/** Inflate a raw-DEFLATE block via the Web DecompressionStream, bounded by `maxBytes`. */
async function inflateRaw(data: Uint8Array, maxBytes: number): Promise<Uint8Array> {
  const ds = new DecompressionStream("deflate-raw");
  const writer = ds.writable.getWriter();
  // Narrow the backing buffer to ArrayBuffer for write()'s BufferSource type — a subarray view is
  // Uint8Array<ArrayBufferLike>, which TS won't accept (ArrayBufferLike admits SharedArrayBuffer).
  // Zero-copy when already ArrayBuffer-backed (the common case), which it always is here.
  const chunk =
    data.buffer instanceof ArrayBuffer
      ? new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
      : new Uint8Array(data);
  // Swallow these promises' rejections: cancelling the reader on a budget overflow aborts the
  // writable side, so the in-flight write/close reject with AbortError. Left dangling (`void`),
  // that surfaces as an unhandled rejection under Bun and pollutes the test run.
  writer.write(chunk).catch(() => {});
  writer.close().catch(() => {});
  const reader = ds.readable.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      await reader.cancel();
      throw new XlsxError("inflated output exceeds the byte budget");
    }
    chunks.push(value);
  }
  const out = new Uint8Array(total);
  let o = 0;
  for (const c of chunks) {
    out.set(c, o);
    o += c.byteLength;
  }
  return out;
}

/** Locate the End-of-Central-Directory record by scanning backward (the trailing comment can be up
 * to 65535 bytes). */
function findEocd(view: DataView): number {
  const max = Math.max(0, view.byteLength - 22 - 65535);
  for (let p = view.byteLength - 22; p >= max; p--) {
    if (view.getUint32(p, true) === EOCD_SIG) return p;
  }
  throw new XlsxError("not a valid .xlsx (no ZIP end-of-central-directory record)");
}

/** Parse a .xlsx (ZIP) into path → uncompressed bytes. Rejects non-ZIP / OLE-encrypted / oversized /
 * out-of-bounds inputs with XlsxError. */
export async function unzip(
  bytes: Uint8Array,
  maxInflatedBytes: number,
): Promise<Map<string, Uint8Array>> {
  if (
    bytes.length >= 8 &&
    bytes[0] === 0xd0 &&
    bytes[1] === 0xcf &&
    bytes[2] === 0x11 &&
    bytes[3] === 0xe0
  ) {
    throw new XlsxError("encrypted workbooks are unsupported");
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocd = findEocd(view);
  const cdCount = view.getUint16(eocd + 10, true);
  let p = view.getUint32(eocd + 16, true); // central directory offset
  const dec = new TextDecoder("utf-8");
  const out = new Map<string, Uint8Array>();
  let budget = maxInflatedBytes;
  for (let i = 0; i < cdCount; i++) {
    if (p + 46 > bytes.length || view.getUint32(p, true) !== CD_SIG) {
      throw new XlsxError("corrupt ZIP central directory");
    }
    const method = view.getUint16(p + 10, true);
    const compSize = view.getUint32(p + 20, true);
    const nameLen = view.getUint16(p + 28, true);
    const extraLen = view.getUint16(p + 30, true);
    const commentLen = view.getUint16(p + 32, true);
    const localOff = view.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    // Local header: sig(4) … nameLen@26(2) extraLen@28(2); data follows.
    if (localOff + 30 > bytes.length) throw new XlsxError("corrupt ZIP local header offset");
    const lNameLen = view.getUint16(localOff + 26, true);
    const lExtraLen = view.getUint16(localOff + 28, true);
    const dataStart = localOff + 30 + lNameLen + lExtraLen;
    if (dataStart + compSize > bytes.length) throw new XlsxError("corrupt ZIP entry data range");
    const comp = bytes.subarray(dataStart, dataStart + compSize);
    let data: Uint8Array;
    if (method === 0) {
      data = comp;
    } else if (method === 8) {
      data = await inflateRaw(comp, budget);
    } else {
      throw new XlsxError(`unsupported ZIP compression method ${method}`);
    }
    budget -= data.byteLength;
    if (budget < 0) throw new XlsxError("inflated output exceeds the byte budget");
    if (!name.endsWith("/")) out.set(name, data);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}
