/**
 * Minimal, dependency-free .xlsx / .xlsm reader: returns the cell text of one worksheet as a
 * row-major string grid. Only what AllocationLab needs — values, not formatting or formulas.
 * Uses the platform's DecompressionStream('deflate-raw') and DOMParser.
 * Macros in .xlsm files are never read or executed.
 */

interface ZipEntry {
  name: string;
  method: number;
  compressedSize: number;
  localOffset: number;
}

function u16(b: Uint8Array, o: number): number {
  return b[o]! | (b[o + 1]! << 8);
}
function u32(b: Uint8Array, o: number): number {
  return (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0;
}

function readZipDirectory(b: Uint8Array): ZipEntry[] {
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 65557); i--) {
    if (u32(b, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('Not a valid .xlsx file (zip directory not found).');
  const count = u16(b, eocd + 10);
  let p = u32(b, eocd + 16);
  const entries: ZipEntry[] = [];
  const dec = new TextDecoder();
  for (let i = 0; i < count; i++) {
    if (u32(b, p) !== 0x02014b50) throw new Error('Corrupt .xlsx file (bad central directory).');
    const method = u16(b, p + 10);
    const compressedSize = u32(b, p + 20);
    const nameLen = u16(b, p + 28);
    const extraLen = u16(b, p + 30);
    const commentLen = u16(b, p + 32);
    const localOffset = u32(b, p + 42);
    const name = dec.decode(b.subarray(p + 46, p + 46 + nameLen));
    entries.push({ name, method, compressedSize, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const ds = new DecompressionStream('deflate-raw');
  const source = new ReadableStream<Uint8Array>({
    start(controller) {
      controller.enqueue(data);
      controller.close();
    },
  });
  const stream = source.pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

async function readEntry(b: Uint8Array, e: ZipEntry): Promise<string> {
  const p = e.localOffset;
  if (u32(b, p) !== 0x04034b50) throw new Error(`Corrupt .xlsx file (bad local header for ${e.name}).`);
  const start = p + 30 + u16(b, p + 26) + u16(b, p + 28);
  const raw = b.subarray(start, start + e.compressedSize);
  let bytes: Uint8Array;
  if (e.method === 0) bytes = raw;
  else if (e.method === 8) bytes = await inflateRaw(raw);
  else throw new Error(`Unsupported compression method ${e.method} in ${e.name}.`);
  return new TextDecoder().decode(bytes);
}

function colIndex(ref: string): number {
  const letters = /^[A-Z]+/.exec(ref)?.[0] ?? 'A';
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Concatenate <t> text of an element, ignoring phonetic runs (<rPh>). */
function textOf(el: Element): string {
  let out = '';
  const ts = el.getElementsByTagName('t');
  for (let i = 0; i < ts.length; i++) {
    const t = ts[i]!;
    if (t.parentElement?.localName === 'rPh') continue;
    out += t.textContent ?? '';
  }
  return out;
}

export interface XlsxSheet {
  sheetName: string;
  sheetNames: string[];
  rows: string[][];
}

/**
 * Read one worksheet. `preferSheet` (case-insensitive) is used when present, otherwise the first sheet.
 */
export async function readXlsxSheet(buffer: ArrayBuffer, preferSheet?: string): Promise<XlsxSheet> {
  const b = new Uint8Array(buffer);
  const entries = readZipDirectory(b);
  const byName = new Map(entries.map((e) => [e.name, e]));
  const get = async (n: string) => {
    const e = byName.get(n);
    return e ? readEntry(b, e) : null;
  };
  const parser = new DOMParser();
  const xml = (s: string) => parser.parseFromString(s, 'application/xml');

  const wbXml = await get('xl/workbook.xml');
  if (!wbXml) throw new Error('Not an Excel workbook (xl/workbook.xml missing).');
  const relsXml = await get('xl/_rels/workbook.xml.rels');
  const rels = new Map<string, string>();
  if (relsXml) {
    const rs = xml(relsXml).getElementsByTagName('Relationship');
    for (let i = 0; i < rs.length; i++) rels.set(rs[i]!.getAttribute('Id') ?? '', rs[i]!.getAttribute('Target') ?? '');
  }
  const sheetEls = xml(wbXml).getElementsByTagName('sheet');
  const sheets: { name: string; target: string }[] = [];
  for (let i = 0; i < sheetEls.length; i++) {
    const s = sheetEls[i]!;
    const rid = s.getAttribute('r:id') ?? s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') ?? '';
    let target = rels.get(rid) ?? `worksheets/sheet${i + 1}.xml`;
    target = target.startsWith('/') ? target.slice(1) : `xl/${target}`;
    sheets.push({ name: s.getAttribute('name') ?? `Sheet${i + 1}`, target });
  }
  if (sheets.length === 0) throw new Error('The workbook has no worksheets.');
  const chosen = (preferSheet && sheets.find((s) => s.name.toLowerCase() === preferSheet.toLowerCase())) || sheets[0]!;

  const shared: string[] = [];
  const ssXml = await get('xl/sharedStrings.xml');
  if (ssXml) {
    const sis = xml(ssXml).getElementsByTagName('si');
    for (let i = 0; i < sis.length; i++) shared.push(textOf(sis[i]!));
  }

  const sheetXml = await get(chosen.target);
  if (!sheetXml) throw new Error(`Worksheet "${chosen.name}" could not be read.`);
  const rowEls = xml(sheetXml).getElementsByTagName('row');
  const rows: string[][] = [];
  for (let i = 0; i < rowEls.length; i++) {
    const r = rowEls[i]!;
    const rIdx = Number(r.getAttribute('r') ?? rows.length + 1) - 1;
    const row: string[] = [];
    const cells = r.getElementsByTagName('c');
    for (let j = 0; j < cells.length; j++) {
      const c = cells[j]!;
      const t = c.getAttribute('t');
      const ci = colIndex(c.getAttribute('r') ?? '');
      const v = c.getElementsByTagName('v')[0]?.textContent ?? '';
      let val: string;
      if (t === 's') val = shared[Number(v)] ?? '';
      else if (t === 'inlineStr') {
        const is = c.getElementsByTagName('is')[0];
        val = is ? textOf(is) : '';
      } else val = v;
      while (row.length < ci) row.push('');
      row[ci] = val;
    }
    while (rows.length < rIdx) rows.push([]);
    rows[rIdx] = row;
  }
  return { sheetName: chosen.name, sheetNames: sheets.map((s) => s.name), rows };
}
