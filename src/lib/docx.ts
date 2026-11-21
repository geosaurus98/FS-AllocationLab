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

  private constructor(entries: ZipEntry[], parts: { doc: string; rels: string; numbering: string | null; ct: string }) {
    this.entries = entries;
    const p = new DOMParser();
    this.doc = p.parseFromString(parts.doc, 'application/xml');
    this.rels = p.parseFromString(parts.rels, 'application/xml');
    this.numbering = parts.numbering ? p.parseFromString(parts.numbering, 'application/xml') : null;
    this.contentTypes = p.parseFromString(parts.ct, 'application/xml');
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
    return new DocxDocument(entries, { doc, rels, numbering: await get('word/numbering.xml'), ct });
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
  replaceText(find: string, replacement: string, opts: { wholeWord?: boolean; matchCase?: boolean } = {}): number {
    const whole = opts.wholeWord ?? true;
    const exact = opts.matchCase ?? false;
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
        this.spliceText(ts, m.s, m.e, exact ? replacement : matchCase(m.found, replacement));
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
      this.el('tblCellMar', {}, this.el('left', { w: '57', type: 'dxa' }), this.el('right', { w: '57', type: 'dxa' })),
    );
    const grid = this.el('tblGrid', {}, ...t.widths.map((w) => this.el('gridCol', { w: twips(w) })));
    const tbl = this.el('tbl', {}, tblPr, grid);
    const n = t.rows.length;
    t.rows.forEach((row, i) => {
      const tr = this.el('tr', {}, this.el('trPr', {}, this.el('cantSplit')));
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
  signature(lead: string, secondary: string, deleteBookmarks: string[]): { built: boolean; missing: string[] } {
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
    const widths = [380 * 0.75, 300 * 0.75];
    const tbl = this.el(
      'tbl',
      {},
      this.el(
        'tblPr',
        {},
        this.el('tblW', { w: twips(widths[0]! + widths[1]!), type: 'dxa' }),
        this.el('jc', { val: 'left' }),
        this.el('tblInd', { w: twips(22 * 0.75), type: 'dxa' }),
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
