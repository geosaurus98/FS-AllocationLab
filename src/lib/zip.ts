/**
 * Minimal, dependency-free zip reader and writer for Office files (.xlsx, .docx).
 * Uses the platform's DecompressionStream / CompressionStream ('deflate-raw'). No zip64.
 */

export interface ZipEntry {
  name: string;
  method: number;
  crc: number;
  compressedSize: number;
  size: number;
  /** The entry's stored (possibly compressed) bytes. */
  raw: Uint8Array;
}

const u16 = (b: Uint8Array, o: number) => b[o]! | (b[o + 1]! << 8);
const u32 = (b: Uint8Array, o: number) => (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0;

/** Read a zip's central directory. Entry data is sliced, not decompressed. */
export function readZip(bytes: Uint8Array | ArrayBuffer, what = 'Office file'): ZipEntry[] {
  const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) {
    if (u32(b, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error(`Not a valid ${what} (zip directory not found).`);
  const count = u16(b, eocd + 10);
  let p = u32(b, eocd + 16);
  const dec = new TextDecoder();
  const out: ZipEntry[] = [];
  for (let i = 0; i < count; i++) {
    if (u32(b, p) !== 0x02014b50) throw new Error(`Corrupt ${what} (bad zip central directory).`);
    const method = u16(b, p + 10);
    const crc = u32(b, p + 16);
    const compressedSize = u32(b, p + 20);
    const size = u32(b, p + 24);
    const nameLen = u16(b, p + 28);
    const extraLen = u16(b, p + 30);
    const commentLen = u16(b, p + 32);
    const localOffset = u32(b, p + 42);
    const name = dec.decode(b.subarray(p + 46, p + 46 + nameLen));
    if (u32(b, localOffset) !== 0x04034b50) throw new Error(`Corrupt ${what} (bad local header for ${name}).`);
    const start = localOffset + 30 + u16(b, localOffset + 26) + u16(b, localOffset + 28);
    out.push({ name, method, crc, compressedSize, size, raw: b.subarray(start, start + compressedSize) });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

async function pipe(data: Uint8Array, stream: { readable: ReadableStream<Uint8Array>; writable: WritableStream<Uint8Array> }): Promise<Uint8Array> {
  const source = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  });
  return new Uint8Array(await new Response(source.pipeThrough(stream)).arrayBuffer());
}

/** Decompressed bytes of an entry. */
export async function entryBytes(e: ZipEntry): Promise<Uint8Array> {
  if (e.method === 0) return e.raw;
  if (e.method === 8) return pipe(e.raw, new DecompressionStream('deflate-raw'));
  throw new Error(`Unsupported compression method ${e.method} in ${e.name}.`);
}

export async function entryText(e: ZipEntry): Promise<string> {
  return new TextDecoder().decode(await entryBytes(e));
}

let CRC_TABLE: Uint32Array | null = null;
export function crc32(data: Uint8Array): number {
  if (!CRC_TABLE) {
    CRC_TABLE = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      CRC_TABLE[n] = c >>> 0;
    }
  }
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) c = CRC_TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** A file to write: new content (`data`), or an entry copied unchanged from a zip that was read. */
export type ZipInput = { name: string; data: Uint8Array | string } | ZipEntry;

/** Build a zip. New content is deflated when CompressionStream is available, otherwise stored. */
export async function writeZip(files: ZipInput[]): Promise<Uint8Array> {
  const enc = new TextEncoder();
  const canDeflate = typeof CompressionStream !== 'undefined';
  const prepared: { name: Uint8Array; method: number; crc: number; size: number; raw: Uint8Array }[] = [];
  for (const f of files) {
    if ('raw' in f) {
      prepared.push({ name: enc.encode(f.name), method: f.method, crc: f.crc, size: f.size, raw: f.raw });
      continue;
    }
    const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
    const raw = canDeflate ? await pipe(data, new CompressionStream('deflate-raw')) : data;
    prepared.push({ name: enc.encode(f.name), method: canDeflate ? 8 : 0, crc: crc32(data), size: data.length, raw });
  }
  const chunks: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const header = (len: number) => {
    const a = new Uint8Array(len);
    return { a, v: new DataView(a.buffer) };
  };
  for (const p of prepared) {
    const { a, v } = header(30);
    v.setUint32(0, 0x04034b50, true);
    v.setUint16(4, 20, true);
    v.setUint16(6, 0x0800, true); // UTF-8 names
    v.setUint16(8, p.method, true);
    v.setUint16(10, 0, true);
    v.setUint16(12, 0x21, true); // 1980-01-01
    v.setUint32(14, p.crc, true);
    v.setUint32(18, p.raw.length, true);
    v.setUint32(22, p.size, true);
    v.setUint16(26, p.name.length, true);
    v.setUint16(28, 0, true);
    chunks.push(a, p.name, p.raw);
    const c = header(46);
    c.v.setUint32(0, 0x02014b50, true);
    c.v.setUint16(4, 20, true);
    c.v.setUint16(6, 20, true);
    c.v.setUint16(8, 0x0800, true);
    c.v.setUint16(10, p.method, true);
    c.v.setUint16(14, 0x21, true);
    c.v.setUint32(16, p.crc, true);
    c.v.setUint32(20, p.raw.length, true);
    c.v.setUint32(24, p.size, true);
    c.v.setUint16(28, p.name.length, true);
    c.v.setUint32(42, offset, true);
    central.push(c.a, p.name);
    offset += 30 + p.name.length + p.raw.length;
  }
  const centralSize = central.reduce((s, x) => s + x.length, 0);
  const end = header(22);
  end.v.setUint32(0, 0x06054b50, true);
  end.v.setUint16(8, prepared.length, true);
  end.v.setUint16(10, prepared.length, true);
  end.v.setUint32(12, centralSize, true);
  end.v.setUint32(16, offset, true);
  const all = [...chunks, ...central, end.a];
  const out = new Uint8Array(all.reduce((s, x) => s + x.length, 0));
  let p = 0;
  for (const x of all) {
    out.set(x, p);
    p += x.length;
  }
  return out;
}
