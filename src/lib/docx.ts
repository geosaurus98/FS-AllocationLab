/**
 * Fill Word (.docx) letter templates in the browser, the way the workbook's Word macros do:
 * replace text (across runs, whole words, Word's case matching), put tables and bulleted
 * paragraphs where placeholder text is, build the adviser signature table from bookmarks and
 * delete bookmarked blocks. Only word/document.xml (the main story) is changed, as in the macros.
 *
 * Uses DOMParser / XMLSerializer and the zip helpers; nothing leaves the browser.
 */

import type { DocCell, DocTable, RationaleItem } from '../engine/workbook/documentTables';
import { entryText, readZip, writeZip, type ZipEntry, type ZipInput } from './zip';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PR = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const REL_HYPERLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
const REL_NUMBERING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering';

const FONT = 'Verdana';
const twips = (pt: number) => String(Math.round(pt * 20));

const WORD_CHAR = /[\p{L}\p{N}_]/u;
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** Word's case matching when "Match case" is off: ALL CAPS stays all caps, Initial Cap stays initial cap. */
function matchCase(found: string, replacement: string): string {
  if (found.length > 1 && /\p{L}/u.test(found) && found === found.toUpperCase()) return replacement.toUpperCase();
  const f = found.charAt(0);
  if (f !== f.toLowerCase() && f === f.toUpperCase()) return replacement.charAt(0).toUpperCase() + replacement.slice(1);
  return replacement;
}

export interface FillReport {
  found: string[];
  missing: string[];
}

export class DocxDocument {
  private entries: ZipEntry[];
  private changed = new Map<string, string>();
  doc: Document;
  private rels: Document;
  private numbering: Document | null;
  private contentTypes: Document;
  private bulletNumId: string | null = null;
  private relSeq = 0;
  private settings: Document | null;
  private updateFieldsOnOpen = false;

  private constructor(entries: ZipEntry[], parts: { doc: string; rels: string; numbering: string | null; ct: string; settings?: string | null }) {
    this.entries = entries;
    const p = new DOMParser();
    this.doc = p.parseFromString(parts.doc, 'application/xml');
    this.rels = p.parseFromString(parts.rels, 'application/xml');
    this.numbering = parts.numbering ? p.parseFromString(parts.numbering, 'application/xml') : null;
    this.contentTypes = p.parseFromString(parts.ct, 'application/xml');
    this.settings = parts.settings ? p.parseFromString(parts.settings, 'application/xml') : null;
    if (this.doc.getElementsByTagName('parsererror').length) throw new Error('The Word document could not be read (document.xml is not valid XML).');
  }

  static async open(bytes: Uint8Array | ArrayBuffer): Promise<DocxDocument> {
    const entries = readZip(bytes, 'Word .docx file');
    const get = async (n: string) => {
      const e = entries.find((x) => x.name === n);
      return e ? entryText(e) : null;
    };
    const doc = await get('word/document.xml');
    if (!doc) throw new Error('This is not a Word .docx document (word/document.xml is missing).');
    const rels = (await get('word/_rels/document.xml.rels')) ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}"/>`;
    const ct = await get('[Content_Types].xml');
    if (!ct) throw new Error('This is not a Word .docx document ([Content_Types].xml is missing).');
    return new DocxDocument(entries, { doc, rels, numbering: await get('word/numbering.xml'), ct, settings: await get('word/settings.xml') });
  }

  /** A new, empty document (used when no firm template is loaded). */
  static blank(paragraphs: { text: string; style?: 'title' | 'heading' | 'note' }[]): DocxDocument {
    const body = paragraphs
      .map((p) => {
        const size = p.style === 'title' ? 28 : p.style === 'heading' ? 22 : p.style === 'note' ? 16 : 20;
        const b = p.style === 'title' || p.style === 'heading' ? '<w:b/>' : '';
        const color = p.style === 'note' ? '<w:color w:val="7F7F7F"/>' : '';
        const space = p.style === 'heading' ? '<w:spacing w:before="240" w:after="120"/>' : '<w:spacing w:before="0" w:after="120"/>';
        return `<w:p><w:pPr>${space}</w:pPr><w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>${b}${color}<w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t xml:space="preserve">${escapeXml(p.text)}</w:t></w:r></w:p>`;
      })
      .join('');
    const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
    const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
    const d = new DocxDocument([], { doc, rels: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PR}"/>`, numbering: null, ct });
    d.changed.set('_rels/.rels', rootRels);
    return d;
  }

  // ------------------------------------------------------------ basics

  private el(name: string, attrs: Record<string, string> = {}, ...children: (Node | null | undefined)[]): Element {
    const e = this.doc.createElementNS(W, `w:${name}`);
    for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, `w:${k}`, v);
    for (const c of children) if (c) e.appendChild(c);
    return e;
  }

  private get body(): Element {
    const b = this.doc.getElementsByTagNameNS(W, 'body')[0];
    if (!b) throw new Error('The Word document has no body.');
    return b;
  }

  private paragraphs(): Element[] {
    return Array.from(this.doc.getElementsByTagNameNS(W, 'p'));
  }

  /** w:t elements belonging to this paragraph (not to a text box nested inside it). */
  private textNodes(p: Element): Element[] {
    return Array.from(p.getElementsByTagNameNS(W, 't')).filter((t) => closest(t, 'p') === p);
  }

  paragraphText(p: Element): string {
    return this.textNodes(p)
      .map((t) => t.textContent ?? '')
      .join('');
  }

  /** All text of the main story, paragraphs separated by newlines. */
  text(): string {
    return this.paragraphs()
      .map((p) => this.paragraphText(p))
      .join('\n');
  }

  // ------------------------------------------------------------ find and replace

  /** ReplaceTextInWord: replace all, whole words, case-insensitive with Word's case matching. Returns the count. */
  replaceText(find: string, replacement: string, opts: { wholeWord?: boolean; matchCase?: boolean; adjustCase?: boolean } = {}): number {
    const whole = opts.wholeWord ?? true;
    const exact = opts.matchCase ?? false;
    const adjust = !exact && (opts.adjustCase ?? true);
    const re = new RegExp(escapeRe(find), exact ? 'g' : 'gi');
    let count = 0;
    for (const p of this.paragraphs()) {
      const ts = this.textNodes(p);
      if (ts.length === 0) continue;
      const text = ts.map((t) => t.textContent ?? '').join('');
      const matches: { s: number; e: number; found: string }[] = [];
      for (const m of text.matchAll(re)) {
        const s = m.index!;
        const e = s + m[0].length;
        if (whole && ((s > 0 && WORD_CHAR.test(text[s - 1]!)) || (e < text.length && WORD_CHAR.test(text[e]!)))) continue;
        matches.push({ s, e, found: m[0] });
      }
      for (const m of matches.reverse()) {
        this.spliceText(ts, m.s, m.e, adjust ? matchCase(m.found, replacement) : replacement);
        count++;
      }
    }
    return count;
  }

  /** Replace characters [s, e) of a paragraph's concatenated text; the new text takes the first run's format. */
  private spliceText(ts: Element[], s: number, e: number, replacement: string) {
    let pos = 0;
    let first = true;
    for (const t of ts) {
      const txt = t.textContent ?? '';
      const a = pos;
      const b = pos + txt.length;
      pos = b;
      if (b <= s || a >= e) continue;
      const from = Math.max(s, a) - a;
      const to = Math.min(e, b) - a;
      const next = txt.slice(0, from) + (first ? replacement : '') + txt.slice(to);
      first = false;
      t.textContent = next;
      t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    }
  }

  /** First paragraph containing `text` (case-insensitive, like Word's Find). */
  findParagraph(text: string): Element | null {
    const t = text.toLowerCase();
    return this.paragraphs().find((p) => this.paragraphText(p).toLowerCase().includes(t)) ?? null;
  }

  /** Remove the first occurrence of `text` from paragraph p. */
  private removeText(p: Element, text: string) {
    const ts = this.textNodes(p);
    const all = ts.map((t) => t.textContent ?? '').join('');
    const i = all.toLowerCase().indexOf(text.toLowerCase());
    if (i >= 0) this.spliceText(ts, i, i + text.length, '');
  }

  // ------------------------------------------------------------ tables

  private cellXml(c: DocCell, width: number, fontSize: number, borders: Partial<Record<'top' | 'left' | 'bottom' | 'right', number>>): Element {
    const b = Object.entries(borders);
    const tcBorders = b.length
      ? this.el(
          'tcBorders',
          {},
          ...(['top', 'left', 'bottom', 'right'] as const).filter((k) => borders[k]).map((k) => this.el(k, { val: 'single', sz: String(borders[k]), space: '0', color: '000000' })),
        )
      : null;
    const tcPr = this.el('tcPr', {}, this.el('tcW', { w: twips(width), type: 'dxa' }), tcBorders, c.fill ? this.el('shd', { val: 'clear', color: 'auto', fill: c.fill }) : null, this.el('vAlign', { val: 'center' }));
    const p = this.el('p', {}, this.el('pPr', {}, this.el('spacing', { before: '0', after: '0', line: '240', lineRule: 'auto' }), this.el('jc', { val: c.align === 'center' ? 'center' : 'left' })));
    if (c.text) {
      const run = this.run(c.text, { size: fontSize, bold: c.bold, color: c.link ? '0563C1' : c.color, underline: !!c.link });
      p.appendChild(c.link ? this.hyperlink(c.link, run) : run);
    }
    return this.el('tc', {}, tcPr, p);
  }

  private run(text: string, f: { size: number; bold?: boolean; color?: string; underline?: boolean }): Element {
    const rPr = this.el(
      'rPr',
      {},
      this.el('rFonts', { ascii: FONT, hAnsi: FONT, cs: FONT }),
      f.bold ? this.el('b') : this.el('b', { val: '0' }),
      f.color ? this.el('color', { val: f.color }) : null,
      this.el('sz', { val: String(f.size * 2) }),
      this.el('szCs', { val: String(f.size * 2) }),
      this.el('u', { val: f.underline ? 'single' : 'none' }),
    );
    const t = this.el('t');
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    t.textContent = text;
    return this.el('r', {}, rPr, t);
  }

  private hyperlink(address: string, ...runs: Element[]): Element {
    const id = this.addRelationship(REL_HYPERLINK, address, true);
    const h = this.el('hyperlink', { history: '1' }, ...runs);
    h.setAttributeNS(R, 'r:id', id);
    return h;
  }

  private addRelationship(type: string, target: string, external: boolean): string {
    const root = this.rels.documentElement;
    const ids = new Set(Array.from(root.getElementsByTagNameNS(PR, 'Relationship')).map((r) => r.getAttribute('Id')));
    let id: string;
    do id = `rIdAL${++this.relSeq}`;
    while (ids.has(id));
    const rel = this.rels.createElementNS(PR, 'Relationship');
    rel.setAttribute('Id', id);
    rel.setAttribute('Type', type);
    rel.setAttribute('Target', target);
    if (external) rel.setAttribute('TargetMode', 'External');
    root.appendChild(rel);
    return id;
  }

  buildTable(t: DocTable): Element {
    const total = t.widths.reduce((a, b) => a + b, 0);
    const tblPr = this.el(
      'tblPr',
      {},
      this.el('tblW', { w: twips(total), type: 'dxa' }),
      t.center ? this.el('jc', { val: 'center' }) : null,
      this.el('tblBorders', {}, ...['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => this.el(k, { val: 'nil' }))),
      this.el('tblLayout', { type: 'fixed' }),
      t.padding
        ? this.el('tblCellMar', {}, this.el('top', { w: twips(t.padding.top), type: 'dxa' }), this.el('left', { w: twips(t.padding.left), type: 'dxa' }), this.el('bottom', { w: twips(t.padding.bottom), type: 'dxa' }), this.el('right', { w: twips(t.padding.right), type: 'dxa' }))
        : this.el('tblCellMar', {}, this.el('left', { w: '57', type: 'dxa' }), this.el('right', { w: '57', type: 'dxa' })),
    );
    const grid = this.el('tblGrid', {}, ...t.widths.map((w) => this.el('gridCol', { w: twips(w) })));
    const tbl = this.el('tbl', {}, tblPr, grid);
    const n = t.rows.length;
    t.rows.forEach((row, i) => {
      const tr = this.el('tr', {}, this.el('trPr', {}, this.el('cantSplit'), t.minRowHeight ? this.el('trHeight', { val: twips(t.minRowHeight), hRule: 'atLeast' }) : null));
      row.forEach((c, j) => {
        const borders: Partial<Record<'top' | 'left' | 'bottom' | 'right', number>> = {};
        if (t.edgeBorders && i === 0) borders.top = 4;
        if (t.edgeBorders && i === n - 1) borders.bottom = 4;
        if (t.summaryBorders) {
          if (j === 0) borders.left = 4;
          if (j === row.length - 1) borders.right = 4;
          if (i >= 1 && (j === 1 || j === 3)) borders.left = 4;
          if (i >= 1 && j === 1) borders.right = 4;
          if (i === 1) borders.bottom = 12;
          if (c.topBorder) borders.top = 4;
        }
        if (t.columnLines) {
          borders.left = 4;
          borders.right = 4;
        }
        if (t.headerRule && i === 0) {
          borders.top = 4;
          borders.bottom = 4;
        }
        if (t.headerRule && i === n - 1) borders.bottom = 4;
        tr.appendChild(this.cellXml(c, t.widths[j] ?? 60, t.fontSize, borders));
      });
      tbl.appendChild(tr);
    });
    return tbl;
  }

  /** PasteIntoWordAtPlaceholder: the table goes where the placeholder text was. */
  insertTableAt(placeholder: string, t: DocTable): boolean {
    const p = this.findParagraph(placeholder);
    if (!p) return false;
    this.removeText(p, placeholder);
    p.parentNode!.insertBefore(this.buildTable(t), p);
    return true;
  }

  // ------------------------------------------------------------ investment rationale

  private ensureBulletNumbering(): string {
    if (this.bulletNumId) return this.bulletNumId;
    if (!this.numbering) {
      this.numbering = new DOMParser().parseFromString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"/>`, 'application/xml');
      this.addRelationship(REL_NUMBERING, 'numbering.xml', false);
      const types = this.contentTypes.documentElement;
      if (!Array.from(types.getElementsByTagNameNS(CT, 'Override')).some((o) => o.getAttribute('PartName') === '/word/numbering.xml')) {
        const o = this.contentTypes.createElementNS(CT, 'Override');
        o.setAttribute('PartName', '/word/numbering.xml');
        o.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml');
        types.appendChild(o);
      }
    }
    const nd = this.numbering;
    const root = nd.documentElement;
    const mk = (name: string, attrs: Record<string, string> = {}, ...kids: Element[]) => {
      const e = nd.createElementNS(W, `w:${name}`);
      for (const [k, v] of Object.entries(attrs)) e.setAttributeNS(W, `w:${k}`, v);
      kids.forEach((k) => e.appendChild(k));
      return e;
    };
    const maxOf = (tag: string, attr: string) => Math.max(0, ...Array.from(root.getElementsByTagNameNS(W, tag)).map((e) => Number(e.getAttributeNS(W, attr) ?? e.getAttribute(`w:${attr}`)) || 0));
    const absId = String(maxOf('abstractNum', 'abstractNumId') + 1);
    const numId = String(maxOf('num', 'numId') + 1);
    const abs = mk(
      'abstractNum',
      { abstractNumId: absId },
      mk('multiLevelType', { val: 'singleLevel' }),
      mk(
        'lvl',
        { ilvl: '0' },
        mk('start', { val: '1' }),
        mk('numFmt', { val: 'bullet' }),
        mk('lvlText', { val: '' }),
        mk('lvlJc', { val: 'left' }),
        mk('pPr', {}, mk('ind', { left: twips(54), hanging: twips(18) })),
        mk('rPr', {}, mk('rFonts', { ascii: 'Symbol', hAnsi: 'Symbol', hint: 'default' })),
      ),
    );
    const firstNum = root.getElementsByTagNameNS(W, 'num')[0];
    root.insertBefore(abs, firstNum ?? null);
    root.appendChild(mk('num', { numId }, mk('abstractNumId', { val: absId })));
    this.bulletNumId = numId;
    return numId;
  }

  private rationaleParagraph(item: RationaleItem, keepPPr: Element | null): Element {
    const p = this.el('p');
    let pPr: Element;
    if (item.kind === 'bullet') {
      pPr = this.el(
        'pPr',
        {},
        this.el('numPr', {}, this.el('ilvl', { val: '0' }), this.el('numId', { val: this.ensureBulletNumbering() })),
        this.el('spacing', { before: '120', after: '0' }),
        this.el('ind', { left: twips(54), hanging: twips(18) }),
      );
    } else {
      pPr = this.el('pPr', {}, this.el('numPr', {}, this.el('ilvl', { val: '0' }), this.el('numId', { val: '0' })), this.el('ind', { left: twips(14.25), hanging: '0' }));
    }
    const style = keepPPr?.getElementsByTagNameNS(W, 'pStyle')[0];
    if (style) pPr.insertBefore(style.cloneNode(true), pPr.firstChild);
    p.appendChild(pPr);
    const heading = item.kind !== 'bullet';
    p.appendChild(this.run(`${item.text} `, { size: 10, bold: heading, underline: item.kind === 'heading' }));
    if (item.factSheet) {
      p.appendChild(this.run('Link to the fact sheet ', { size: 10 }));
      p.appendChild(this.hyperlink(item.factSheet, this.run('here', { size: 10, color: '0563C1', underline: true })));
      p.appendChild(this.run('.', { size: 10 }));
    }
    return p;
  }

  /** ProcessInvestmentRationaleTable: headings and bullets replace the placeholder paragraph. */
  insertRationaleAt(placeholder: string, items: RationaleItem[]): boolean {
    const p = this.findParagraph(placeholder);
    if (!p) return false;
    this.removeText(p, placeholder);
    const pPr = p.getElementsByTagNameNS(W, 'pPr')[0] ?? null;
    const rest = this.paragraphText(p).trim();
    const parent = p.parentNode!;
    for (const item of items) parent.insertBefore(this.rationaleParagraph(item, pPr), p);
    if (!rest) parent.removeChild(p);
    return true;
  }

  // ------------------------------------------------------------ bookmarks

  bookmarkNames(): string[] {
    return Array.from(this.doc.getElementsByTagNameNS(W, 'bookmarkStart'))
      .map((b) => b.getAttributeNS(W, 'name') ?? b.getAttribute('w:name') ?? '')
      .filter((n) => n && !n.startsWith('_'));
  }

  private bookmark(name: string): { start: Element; end: Element | null } | null {
    const lower = name.toLowerCase();
    const start = Array.from(this.doc.getElementsByTagNameNS(W, 'bookmarkStart')).find((b) => (b.getAttributeNS(W, 'name') ?? b.getAttribute('w:name') ?? '').toLowerCase() === lower);
    if (!start) return null;
    const id = start.getAttributeNS(W, 'id') ?? start.getAttribute('w:id');
    const end = Array.from(this.doc.getElementsByTagNameNS(W, 'bookmarkEnd')).find((b) => (b.getAttributeNS(W, 'id') ?? b.getAttribute('w:id')) === id) ?? null;
    return { start, end };
  }

  hasBookmark(name: string): boolean {
    return this.bookmark(name) !== null;
  }

  /** Body-level element containing node (or the node itself if it is a body child). */
  private block(node: Node): Element {
    let n: Node = node;
    while (n.parentNode && n.parentNode !== this.body) n = n.parentNode;
    return n as Element;
  }

  private blocksOf(name: string): Element[] {
    const bm = this.bookmark(name);
    if (!bm) return [];
    const a = this.block(bm.start);
    const b = bm.end ? this.block(bm.end) : a;
    const out: Element[] = [];
    let n: Element | null = a;
    while (n) {
      out.push(n);
      if (n === b) break;
      n = n.nextElementSibling;
    }
    return out;
  }

  /** Range.Delete on a bookmark: removes what it marks, including whole tables. */
  deleteBookmark(name: string): boolean {
    const bm = this.bookmark(name);
    if (!bm) return false;
    const { start, end } = bm;
    const sb = this.block(start);
    const eb = end ? this.block(end) : sb;
    if (sb === eb && sb.localName === 'p' && end && start.parentNode === end.parentNode) {
      let n = start.nextSibling;
      while (n && n !== end) {
        const next = n.nextSibling;
        n.parentNode!.removeChild(n);
        n = next;
      }
    } else {
      const blocks = this.blocksOf(name);
      for (const b of blocks) {
        if (b === start || b === end) continue;
        if (b === sb && b.localName === 'p' && b !== eb) {
          // keep text before the bookmark in its first paragraph
          if (!this.textBefore(b, start)) b.parentNode!.removeChild(b);
          else this.removeAfter(b, start);
        } else if (b === eb && b.localName === 'p' && b !== sb && end) {
          this.removeBefore(b, end);
        } else {
          b.parentNode?.removeChild(b);
        }
      }
    }
    start.parentNode?.removeChild(start);
    end?.parentNode?.removeChild(end);
    return true;
  }

  private textBefore(p: Element, marker: Element): string {
    let out = '';
    for (const t of this.textNodes(p)) {
      if (marker.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_PRECEDING) out += t.textContent ?? '';
    }
    return out.trim();
  }

  private removeAfter(p: Element, marker: Element) {
    for (const child of Array.from(p.childNodes)) {
      if (child === marker || (child as Element).localName === 'pPr') continue;
      if (marker.compareDocumentPosition(child) & Node.DOCUMENT_POSITION_FOLLOWING && !(child.compareDocumentPosition(marker) & Node.DOCUMENT_POSITION_CONTAINED_BY)) p.removeChild(child);
    }
  }

  private removeBefore(p: Element, marker: Element) {
    for (const child of Array.from(p.childNodes)) {
      if (child === marker || (child as Element).localName === 'pPr') continue;
      if (marker.compareDocumentPosition(child) & Node.DOCUMENT_POSITION_PRECEDING && !(child.compareDocumentPosition(marker) & Node.DOCUMENT_POSITION_CONTAINED_BY)) p.removeChild(child);
    }
  }

  /** First table in (or containing) a bookmark: Bookmarks(name).Range.Tables(1). */
  private bookmarkTable(name: string): Element | null {
    const bm = this.bookmark(name);
    if (!bm) return null;
    const inTable = closest(bm.start, 'tbl');
    if (inTable) return inTable;
    for (const b of this.blocksOf(name)) {
      if (b.localName === 'tbl') return b;
      const t = b.getElementsByTagNameNS(W, 'tbl')[0];
      if (t) return t;
    }
    // a bookmark just before a table
    const next = this.block(bm.start).nextElementSibling;
    return next?.localName === 'tbl' ? next : null;
  }

  /**
   * Signature: a 3 × 2 borderless table at the AdvisorTable bookmark, filled from the first column of
   * the lead adviser's and the secondary adviser's bookmarked signature tables; the adviser
   * bookmarks (John, Hayden, Claudia, BKMRK) are then deleted.
   */
  signature(lead: string, secondary: string, deleteBookmarks: string[], layout: { widths: [number, number]; indent: number } = { widths: [380 * 0.75, 300 * 0.75], indent: 22 * 0.75 }): { built: boolean; missing: string[] } {
    const missing: string[] = [];
    const at = this.bookmark('AdvisorTable');
    if (!at) return { built: false, missing: ['AdvisorTable'] };
    const copyColumn = (name: string): Element[][] | null => {
      if (!name) return null;
      const t = this.bookmarkTable(name);
      if (!t) {
        missing.push(name);
        return null;
      }
      return Array.from(t.getElementsByTagNameNS(W, 'tr'))
        .filter((tr) => closest(tr, 'tbl') === t)
        .map((tr) => {
          const tc = Array.from(tr.childNodes).find((c) => (c as Element).localName === 'tc') as Element | undefined;
          return tc ? (Array.from(tc.childNodes).filter((c) => (c as Element).localName !== 'tcPr').map((c) => c.cloneNode(true)) as Element[]) : [];
        });
    };
    const cols = [copyColumn(lead), copyColumn(secondary)];
    const widths = layout.widths;
    const tbl = this.el(
      'tbl',
      {},
      this.el(
        'tblPr',
        {},
        this.el('tblW', { w: twips(widths[0]! + widths[1]!), type: 'dxa' }),
        this.el('jc', { val: 'left' }),
        this.el('tblInd', { w: twips(layout.indent), type: 'dxa' }),
        this.el('tblBorders', {}, ...['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((k) => this.el(k, { val: 'nil' }))),
        this.el('tblLayout', { type: 'fixed' }),
      ),
      this.el('tblGrid', {}, ...widths.map((w) => this.el('gridCol', { w: twips(w) }))),
    );
    for (let i = 0; i < 3; i++) {
      const tr = this.el('tr');
      for (let j = 0; j < 2; j++) {
        const content = cols[j]?.[i] ?? [];
        const tc = this.el('tc', {}, this.el('tcPr', {}, this.el('tcW', { w: twips(widths[j]!), type: 'dxa' })));
        const paras = content.length ? content : [this.el('p')];
        for (const node of paras) {
          if (node.localName === 'p') {
            let pPr = Array.from(node.childNodes).find((c) => (c as Element).localName === 'pPr') as Element | undefined;
            if (!pPr) node.insertBefore((pPr = this.el('pPr')), node.firstChild);
            for (const s of Array.from(pPr.childNodes)) if ((s as Element).localName === 'spacing') pPr.removeChild(s);
            const spacing = this.el('spacing', { before: '0', after: '0', line: '240', lineRule: 'auto' });
            const after = Array.from(pPr.childNodes).find((c) => ['ind', 'contextualSpacing', 'jc', 'rPr'].includes((c as Element).localName));
            pPr.insertBefore(spacing, after ?? null);
          }
          tc.appendChild(node);
        }
        if (!Array.from(tc.childNodes).some((c) => (c as Element).localName === 'p')) tc.appendChild(this.el('p'));
        tr.appendChild(tc);
      }
      tbl.appendChild(tr);
    }
    const p = this.block(at.start);
    if (p.localName === 'p') {
      for (const t of this.textNodes(p)) t.textContent = '';
      p.parentNode!.insertBefore(tbl, p);
    } else {
      p.parentNode!.insertBefore(tbl, p);
    }
    for (const n of deleteBookmarks) this.deleteBookmark(n);
    return { built: true, missing };
  }

  // ------------------------------------------------------------ bookmarked tables (annual review)

  /** The table holding a bookmark and the bookmarked cell: Bookmarks(name).Range.Tables(1) / .Cells(1). */
  tableAtBookmark(name: string): { tbl: Element; cell: Element | null } | null {
    const bm = this.bookmark(name);
    if (!bm) return null;
    const cell = closest(bm.start, 'tc');
    if (cell) return { tbl: closest(cell, 'tbl')!, cell };
    const tbl = this.bookmarkTable(name);
    if (!tbl) return null;
    return { tbl, cell: this.cellsOf(tbl)[0] ?? null };
  }

  tableRows(tbl: Element): Element[] {
    return Array.from(tbl.childNodes).filter((c) => (c as Element).localName === 'tr') as Element[];
  }

  rowCells(tr: Element): Element[] {
    return Array.from(tr.childNodes).filter((c) => (c as Element).localName === 'tc') as Element[];
  }

  /** Cells in reading order: Cell.Next walks them this way. */
  cellsOf(tbl: Element): Element[] {
    return this.tableRows(tbl).flatMap((tr) => this.rowCells(tr));
  }

  /** Cell (1-based row and column) like Table.Cell(r, c); null when out of range. */
  cellAt(tbl: Element, row: number, col: number): Element | null {
    const tr = this.tableRows(tbl)[row - 1];
    return tr ? (this.rowCells(tr)[col - 1] ?? null) : null;
  }

  cellText(tc: Element): string {
    return Array.from(tc.childNodes)
      .filter((c) => (c as Element).localName === 'p')
      .map((p) => this.paragraphText(p as Element))
      .join('\n');
  }

  /** Cell.Range.Text = text: one paragraph keeping the first paragraph's and first run's formatting. */
  setCellText(tc: Element, text: string, opts: { color?: string } = {}) {
    const paras = Array.from(tc.childNodes).filter((c) => (c as Element).localName === 'p') as Element[];
    const p = paras[0] ?? tc.appendChild(this.el('p'));
    for (const extra of paras.slice(1)) tc.removeChild(extra);
    const firstRun = p.getElementsByTagNameNS(W, 'r')[0];
    const rPr = firstRun ? (Array.from(firstRun.childNodes).find((c) => (c as Element).localName === 'rPr') as Element | undefined) : undefined;
    for (const c of Array.from(p.childNodes)) if ((c as Element).localName !== 'pPr') p.removeChild(c);
    const r = this.el('r');
    const props = rPr ? (rPr.cloneNode(true) as Element) : this.el('rPr');
    if (opts.color) setOrdered(props, this.el('color', { val: opts.color }), RPR_ORDER);
    if (props.childNodes.length) r.appendChild(props);
    const t = this.el('t');
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    p.appendChild(r);
  }

  setCellBorder(tc: Element, side: 'top' | 'bottom', size = 4) {
    const tcPr = this.tcPr(tc);
    let b = Array.from(tcPr.childNodes).find((c) => (c as Element).localName === 'tcBorders') as Element | undefined;
    if (!b) setOrdered(tcPr, (b = this.el('tcBorders')), TCPR_ORDER);
    setOrdered(b, this.el(side, { val: 'single', sz: String(size), space: '0', color: '000000' }), ['top', 'start', 'left', 'bottom', 'end', 'right', 'insideH', 'insideV', 'tl2br', 'tr2bl']);
  }

  private tcPr(tc: Element): Element {
    let pr = Array.from(tc.childNodes).find((c) => (c as Element).localName === 'tcPr') as Element | undefined;
    if (!pr) tc.insertBefore((pr = this.el('tcPr')), tc.firstChild);
    return pr;
  }

  deleteRow(tr: Element) {
    tr.parentNode?.removeChild(tr);
  }

  /** Cell(r1, col).Merge MergeTo:=Cell(r2, col): a vertical merge; the merged cell keeps the first cell's text. */
  mergeVertical(tbl: Element, col: number, r1: number, r2: number) {
    for (let r = r1; r <= r2; r++) {
      const tc = this.cellAt(tbl, r, col);
      if (!tc) continue;
      setOrdered(this.tcPr(tc), this.el('vMerge', r === r1 ? { val: 'restart' } : {}), TCPR_ORDER);
      if (r > r1) this.setCellText(tc, '');
    }
  }

  // ------------------------------------------------------------ more bookmark operations

  /** Bookmarks(name).Range.Text = text */
  setBookmarkText(name: string, text: string): boolean {
    const bm = this.bookmark(name);
    if (!bm) return false;
    const { start, end } = bm;
    let rPr: Element | null = null;
    if (end && start.parentNode === end.parentNode) {
      let n = start.nextSibling;
      while (n && n !== end) {
        const next = n.nextSibling;
        if (!rPr && (n as Element).localName === 'r') rPr = (Array.from(n.childNodes).find((c) => (c as Element).localName === 'rPr') as Element | undefined) ?? null;
        n.parentNode!.removeChild(n);
        n = next;
      }
    }
    if (!rPr) {
      let prev = start.previousSibling;
      while (prev && (prev as Element).localName !== 'r') prev = prev.previousSibling;
      if (prev) rPr = (Array.from(prev.childNodes).find((c) => (c as Element).localName === 'rPr') as Element | undefined) ?? null;
    }
    const r = this.el('r');
    if (rPr) r.appendChild(rPr.cloneNode(true));
    const t = this.el('t');
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    if (start.parentNode && (start.parentNode as Element).localName === 'p') start.parentNode.insertBefore(r, start.nextSibling);
    else {
      const p = this.block(start);
      if (p.localName === 'p') p.appendChild(r);
      else return false;
    }
    return true;
  }

  /** Expand the bookmark to its paragraph and the paragraph before it, and delete both (empty Sig2–Sig4). */
  deleteBookmarkParagraphWithPrevious(name: string): boolean {
    const bm = this.bookmark(name);
    if (!bm) return false;
    const p = (bm.start.parentNode as Element).localName === 'p' ? (bm.start.parentNode as Element) : closest(bm.start, 'p');
    if (!p) return false;
    let prev = p.previousElementSibling;
    while (prev && prev.localName !== 'p') prev = null;
    if (prev) prev.parentNode!.removeChild(prev);
    p.parentNode!.removeChild(p);
    bm.end?.parentNode?.removeChild(bm.end);
    return true;
  }

  /** HighlightTextInWord: highlight every whole-word, case-insensitive occurrence. Returns the count. */
  highlight(text: string, color = 'yellow'): number {
    const re = new RegExp(escapeRe(text), 'gi');
    let count = 0;
    for (const p of this.paragraphs()) {
      const full = this.paragraphText(p);
      const hits: [number, number][] = [];
      for (const m of full.matchAll(re)) {
        const s = m.index!;
        const e = s + m[0].length;
        if ((s > 0 && WORD_CHAR.test(full[s - 1]!)) || (e < full.length && WORD_CHAR.test(full[e]!) && WORD_CHAR.test(text.at(-1)!))) continue;
        hits.push([s, e]);
      }
      for (const [s, e] of hits.reverse()) {
        for (const r of this.isolate(p, s, e)) {
          let rPr = Array.from(r.childNodes).find((c) => (c as Element).localName === 'rPr') as Element | undefined;
          if (!rPr) r.insertBefore((rPr = this.el('rPr')), r.firstChild);
          setOrdered(rPr, this.el('highlight', { val: color }), RPR_ORDER);
        }
        count++;
      }
    }
    return count;
  }

  /** Split runs so that characters [s, e) of a paragraph sit in whole runs; returns those runs. */
  private isolate(p: Element, s: number, e: number): Element[] {
    const out: Element[] = [];
    let pos = 0;
    for (const t of this.textNodes(p)) {
      const txt = t.textContent ?? '';
      const a = pos;
      const b = pos + txt.length;
      pos = b;
      if (b <= s || a >= e) continue;
      const run = t.parentNode as Element;
      const from = Math.max(s, a) - a;
      const to = Math.min(e, b) - a;
      const ts = Array.from(run.childNodes).filter((c) => (c as Element).localName === 't');
      if (ts.length !== 1 || (from === 0 && to === txt.length)) {
        out.push(run);
        continue;
      }
      const mk = (str: string) => {
        const c = run.cloneNode(true) as Element;
        const ct = Array.from(c.childNodes).find((x) => (x as Element).localName === 't') as Element;
        ct.textContent = str;
        ct.setAttributeNS(XML_NS, 'xml:space', 'preserve');
        return c;
      };
      const parent = run.parentNode!;
      if (from > 0) parent.insertBefore(mk(txt.slice(0, from)), run);
      const mid = mk(txt.slice(from, to));
      parent.insertBefore(mid, run);
      if (to < txt.length) parent.insertBefore(mk(txt.slice(to)), run);
      parent.removeChild(run);
      out.push(mid);
    }
    return out;
  }

  /**
   * Paste another document's body at a bookmark (the annual review commentary insert). Text, tables
   * and formatting are copied; pictures, embedded objects and links that need the other file's parts are left out.
   */
  insertDocumentAt(name: string, other: DocxDocument): { inserted: boolean; droppedPictures: number } {
    const bm = this.bookmark(name);
    if (!bm) return { inserted: false, droppedPictures: 0 };
    const body = other.body;
    let dropped = 0;
    const nodes = Array.from(body.childNodes).filter((c) => (c as Element).localName !== 'sectPr');
    const imported = nodes.map((n) => this.doc.importNode(n, true) as Element);
    for (const n of imported) {
      if (n.nodeType !== 1) continue;
      for (const tag of ['drawing', 'pict', 'object']) {
        for (const d of Array.from(n.getElementsByTagNameNS(W, tag))) {
          d.parentNode?.removeChild(d);
          dropped++;
        }
      }
      for (const tag of ['bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'commentReference', 'footnoteReference', 'endnoteReference']) {
        for (const d of Array.from(n.getElementsByTagNameNS(W, tag))) d.parentNode?.removeChild(d);
      }
      for (const h of Array.from(n.getElementsByTagNameNS(W, 'hyperlink'))) {
        while (h.firstChild) h.parentNode!.insertBefore(h.firstChild, h);
        h.parentNode!.removeChild(h);
      }
      for (const pPr of Array.from(n.getElementsByTagNameNS(W, 'pPr'))) {
        for (const sp of Array.from(pPr.childNodes)) if ((sp as Element).localName === 'sectPr') pPr.removeChild(sp);
      }
    }
    const anchor = this.block(bm.start);
    const parent = anchor.parentNode!;
    for (const n of imported) parent.insertBefore(n, anchor);
    if (anchor.localName === 'p' && !this.paragraphText(anchor).trim()) {
      bm.end?.parentNode?.removeChild(bm.end);
      parent.removeChild(anchor);
    }
    return { inserted: true, droppedPictures: dropped };
  }

  /** Ask Word to update fields (the table of contents) when the document is opened. */
  updateFieldsOnOpenRequest() {
    this.updateFieldsOnOpen = true;
  }

  /**
   * The Statement of Advice appendix table: inserted at a bookmark in its own landscape section,
   * blank paragraphs around it removed, and the rest of the document portrait with 1-inch margins.
   */
  insertLandscapeTableAt(name: string, t: DocTable): boolean {
    const bm = this.bookmark(name);
    if (!bm) return false;
    const anchor = this.block(bm.start);
    const tbl = this.buildTable(t);
    anchor.parentNode!.insertBefore(tbl, anchor);
    if (anchor.localName === 'p') {
      for (const tn of this.textNodes(anchor)) tn.textContent = '';
    }
    // remove empty body paragraphs before and after the table (not section breaks)
    const isBlank = (el: Element) =>
      el.localName === 'p' && !this.paragraphText(el).trim() && !el.getElementsByTagNameNS(W, 'sectPr').length && !el.getElementsByTagNameNS(W, 'drawing').length && !el.getElementsByTagNameNS(W, 'pict').length;
    for (const el of Array.from(this.body.childNodes) as Element[]) if (el !== tbl && el.nodeType === 1 && isBlank(el)) this.body.removeChild(el);
    // the section the table is in: the next paragraph-level sectPr, else the body's
    let sectHolder: Element | null = null;
    for (let n = tbl.nextElementSibling; n; n = n.nextElementSibling) {
      const s = n.localName === 'p' ? Array.from(n.getElementsByTagNameNS(W, 'sectPr'))[0] : null;
      if (s) {
        sectHolder = s;
        break;
      }
    }
    const bodySect = sectHolder ?? (Array.from(this.body.childNodes).find((c) => (c as Element).localName === 'sectPr') as Element | undefined) ?? this.body.appendChild(this.el('sectPr'));
    const landscape = bodySect.cloneNode(true) as Element;
    setPage(this, landscape, { landscape: true, top: 1, bottom: 1, left: 0.01, right: 0.05 });
    setPage(this, bodySect, { landscape: false, top: 2.54, bottom: 2.54, left: 2.54, right: 2.54 });
    const breakPara = this.el('p', {}, this.el('pPr', {}, landscape));
    tbl.parentNode!.insertBefore(breakPara, tbl.nextSibling);
    return true;
  }

  // ------------------------------------------------------------ save

  async save(): Promise<Uint8Array> {
    const ser = new XMLSerializer();
    const out = (d: Document) => {
      const s = ser.serializeToString(d);
      return s.startsWith('<?xml') ? s : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`;
    };
    this.changed.set('word/document.xml', out(this.doc));
    this.changed.set('word/_rels/document.xml.rels', out(this.rels));
    this.changed.set('[Content_Types].xml', out(this.contentTypes));
    if (this.numbering) this.changed.set('word/numbering.xml', out(this.numbering));
    if (this.updateFieldsOnOpen && this.settings) {
      const root = this.settings.documentElement;
      if (!Array.from(root.childNodes).some((c) => (c as Element).localName === 'updateFields')) {
        const u = this.settings.createElementNS(W, 'w:updateFields');
        u.setAttributeNS(W, 'w:val', 'true');
        const before = Array.from(root.childNodes).find((c) => SETTINGS_AFTER_UPDATE_FIELDS.includes((c as Element).localName));
        root.insertBefore(u, before ?? null);
      }
      this.changed.set('word/settings.xml', out(this.settings));
    }
    const files: ZipInput[] = [];
    const seen = new Set<string>();
    const ct = this.changed.get('[Content_Types].xml')!;
    files.push({ name: '[Content_Types].xml', data: ct });
    seen.add('[Content_Types].xml');
    for (const e of this.entries) {
      if (seen.has(e.name)) continue;
      seen.add(e.name);
      const c = this.changed.get(e.name);
      files.push(c !== undefined ? { name: e.name, data: c } : e);
    }
    for (const [name, data] of this.changed) if (!seen.has(name)) files.push({ name, data });
    return writeZip(files);
  }
}

const RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
const TCPR_ORDER = ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark'];
const SETTINGS_AFTER_UPDATE_FIELDS = ['hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'];

/** Put `el` in `parent` in schema order, replacing an element of the same name. */
function setOrdered(parent: Element, el: Element, order: string[]) {
  const name = el.localName;
  for (const c of Array.from(parent.childNodes)) if ((c as Element).localName === name) parent.removeChild(c);
  const idx = order.indexOf(name);
  const after = Array.from(parent.childNodes).find((c) => order.indexOf((c as Element).localName) > idx);
  parent.insertBefore(el, after ?? null);
}

const CM = 567; // twips per centimetre
function setPage(d: DocxDocument, sect: Element, o: { landscape: boolean; top: number; bottom: number; left: number; right: number }) {
  const doc = sect.ownerDocument;
  const mk = (name: string) => doc.createElementNS(W, `w:${name}`);
  const order = ['footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols', 'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];
  const get = (name: string) => Array.from(sect.childNodes).find((c) => (c as Element).localName === name) as Element | undefined;
  const old = get('pgSz');
  const w = Number(old?.getAttributeNS(W, 'w') ?? old?.getAttribute('w:w') ?? 11906) || 11906;
  const h = Number(old?.getAttributeNS(W, 'h') ?? old?.getAttribute('w:h') ?? 16838) || 16838;
  const long = Math.max(w, h);
  const short = Math.min(w, h);
  const pgSz = mk('pgSz');
  pgSz.setAttributeNS(W, 'w:w', String(o.landscape ? long : short));
  pgSz.setAttributeNS(W, 'w:h', String(o.landscape ? short : long));
  if (o.landscape) pgSz.setAttributeNS(W, 'w:orient', 'landscape');
  setOrdered(sect, pgSz, order);
  const oldMar = get('pgMar');
  const pgMar = mk('pgMar');
  const keep = (k: string, dflt: string) => oldMar?.getAttributeNS(W, k) ?? oldMar?.getAttribute(`w:${k}`) ?? dflt;
  pgMar.setAttributeNS(W, 'w:top', String(Math.round(o.top * CM)));
  pgMar.setAttributeNS(W, 'w:right', String(Math.round(o.right * CM)));
  pgMar.setAttributeNS(W, 'w:bottom', String(Math.round(o.bottom * CM)));
  pgMar.setAttributeNS(W, 'w:left', String(Math.round(o.left * CM)));
  pgMar.setAttributeNS(W, 'w:header', keep('header', '709'));
  pgMar.setAttributeNS(W, 'w:footer', keep('footer', '709'));
  pgMar.setAttributeNS(W, 'w:gutter', keep('gutter', '0'));
  setOrdered(sect, pgMar, order);
  void d;
}

function closest(node: Node, localName: string): Element | null {
  let n: Node | null = node.parentNode;
  while (n) {
    if ((n as Element).localName === localName && (n as Element).namespaceURI === W) return n as Element;
    n = n.parentNode;
  }
  return null;
}

function escapeXml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
