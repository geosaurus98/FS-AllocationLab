/**
 * Annual Review (ports of AR_Part1 and AR_Part2) and Statement of Advice appendices
 * (Statement_Of_Advice) built from the firm's Word templates.
 *
 * AR_Part1 fills the Annual Review template: client wording, signatories, the adviser signature table,
 * the performance tables from the custodian report (Table1, Table2, Table3 bookmarks), the period
 * dates and the commentary insert. AR_Part2 then adds the advice tables, as in the letters.
 * Both parts run in one go here.
 */

import type { AdviceTables, DocTable } from '../engine/workbook/documentTables';
import { DocxDocument } from './docx';
import type { LetterClient } from './letters';
import { ALLOCATION_CLASSES, ddMmmYy, pct2, type AllocationClass, type PerformanceReport } from './performanceReport';

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const monthYear = (d: Date) => `${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
const ddmmyy = (d: Date) => `${String(d.getDate()).padStart(2, '0')}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getFullYear()).slice(-2)}`;
const RED = 'FF0000';

/** Details the annual review macro reads from the Master Portfolio Risk Profiles workbook. */
export interface ReviewDetails {
  /** Risk_Profile: replaces "Strategy_Name". */
  strategyName: string;
  /** AssetBlurb: replaces "Insert previous Blurb". */
  blurb: string;
  /** primary_objective: replaces "PrimaryObjective". */
  primaryObjective: string;
  /** Signature: up to four names, comma separated; Sig2–Sig4 paragraphs are removed when empty. */
  signatories: string;
  iOrWe: 'I' | 'We';
  /** RWT and PIR rates; blank leaves "XXX" highlighted. */
  rwt: string;
  pir: string;
  /** Blank leaves "accountant" highlighted. */
  accountant: string;
  /** AuthorityAct: when the client has given authority to act, the "AuthorityToAct" section is removed. */
  authorityToAct: boolean;
}

export const emptyReviewDetails = (): ReviewDetails => ({ strategyName: '', blurb: '', primaryObjective: '', signatories: '', iOrWe: 'We', rwt: '', pir: '', accountant: '', authorityToAct: false });

const HOLDERS_TRUST: [string, string][] = [
  ['holder1', "the Trust's"],
  ['holder2', "the beneficiaries' lives"],
  ['holder3', "your and the Trust's"],
  ['holder4', "the Trustees'"],
  ['holder5', 'the beneficiaries'],
  ['holder6', "The Trust's"],
  ['holder7', 'Current'],
  ['holder8', 'Trust'],
  ['holder9', "beneficiaries'"],
  ['holderA', "The Trustees'"],
  ['holderB', "The Trust's planning"],
];
const HOLDERS_PERSONAL: [string, string][] = [
  ['holder1', 'your'],
  ['holder2', 'your life'],
  ['holder3', 'your'],
  ['holder4', 'your'],
  ['holder5', 'your'],
  ['holder6', 'Your'],
  ['holder7', 'Personal'],
  ['holder8', 'Estate'],
  ['holder9', 'personal'],
  ['holderA', 'Your'],
  ['holderB', 'Your estate planning'],
];

/** AR_Part2 table layout (PasteIntoWordAtPlaceholder in the ARtwo module). */
const AR_LAYOUT: Record<'portfolio' | 'costs' | 'summary', Partial<DocTable>> = {
  portfolio: { widths: [170, 81, 71, 81, 71], fontSize: 10, center: false },
  costs: { widths: [163, 68, 69, 50, 50, 50, 75], fontSize: 8, center: false },
  summary: { widths: [320, 50, 70, 80], fontSize: 10, center: false },
};

const toDecimal = (s: string) => {
  const t = s.trim().split('%')[0]!.trim().replace(',', '.');
  const n = Number(t);
  return t !== '' && Number.isFinite(n) ? n / 100 : 0;
};

/** Table1: period figures down the bookmarked column, inception averages in the last two rows. */
function fillTable1(doc: DocxDocument, r: PerformanceReport, notes: string[]) {
  const at = doc.tableAtBookmark('Table1');
  if (!at) {
    notes.push('Bookmark "Table1" (performance summary) not found in the template.');
    return;
  }
  const { tbl } = at;
  const rows = doc.tableRows(tbl);
  const n = rows.length;
  if (r.inception && n >= 2) {
    const label = (kind: string) => `Net Portfolio Return p.a. from ${ddMmmYy(r.inception!.start)} to ${ddMmmYy(r.inception!.end)} (${kind} average)`;
    const set = (row: Element, a: string, b: string) => {
      const cells = doc.rowCells(row);
      if (cells[0]) doc.setCellText(cells[0], a);
      if (cells[1]) doc.setCellText(cells[1], b);
    };
    set(rows[n - 2]!, label('arithmetic'), pct2(r.inception.arithmetic));
    set(rows[n - 1]!, label('geometric'), pct2(r.inception.geometric));
  }
  if (!r.inception || r.inception.days < 1825) {
    if (n > 2) {
      doc.deleteRow(rows[n - 1]!);
      doc.deleteRow(rows[n - 2]!);
    }
  }
  const cells = doc.cellsOf(tbl);
  const start = at.cell ? cells.indexOf(at.cell) : 0;
  r.table1.forEach((v, k) => {
    const c = cells[start + 2 * k];
    if (c) doc.setCellText(c, v.text, v.negative ? { color: RED } : {});
  });
}

/** Table2: return per sector; rows still showing the template's "%" and "$" are removed. */
function fillTable2(doc: DocxDocument, r: PerformanceReport, notes: string[]) {
  const at = doc.tableAtBookmark('Table2');
  if (!at) {
    notes.push('Bookmark "Table2" (sector returns) not found in the template.');
    return;
  }
  const { tbl } = at;
  if (r.custodian === 'Apex') {
    const cells = doc.cellsOf(tbl);
    const start = at.cell ? cells.indexOf(at.cell) : 0;
    r.sectors.forEach((s, k) => {
      const pc = cells[start + 1 + 3 * k];
      const dc = cells[start + 2 + 3 * k];
      if (s.percent !== null && pc) doc.setCellText(pc, s.percent);
      if (s.dollars !== null && dc) doc.setCellText(dc, s.dollars);
    });
  } else {
    const rows = doc.tableRows(tbl);
    const startRow = at.cell ? rows.indexOf(at.cell.parentNode as Element) : 1;
    for (let i = Math.max(0, startRow); i < rows.length; i++) {
      const cells = doc.rowCells(rows[i]!);
      const label = doc.cellText(cells[0]!).trim();
      if (!label) break;
      const key = label === 'New Zealand Property' ? 'Property' : label;
      const s = r.sectors.find((x) => x.heading === key);
      if (s?.dollars && cells[2]) doc.setCellText(cells[2], s.dollars);
      if (label === 'Multi Sector') break;
    }
    // column 2 of rows 2–7 is set to "%" (the macro never fills the OneAnswer sector percentages)
    for (let i = 2; i <= 7; i++) {
      const c = doc.cellAt(tbl, i, 2);
      if (c) doc.setCellText(c, '%');
    }
    const c41 = doc.cellAt(tbl, 4, 1);
    if (c41 && doc.cellText(c41).toLowerCase().includes('new zealand property')) doc.setCellText(c41, doc.cellText(c41).replace(/New Zealand Property/i, 'Property'));
  }
  for (const row of [...doc.tableRows(tbl)].reverse()) {
    const cells = doc.rowCells(row);
    if (cells.length >= 3 && doc.cellText(cells[1]!).trim() === '%' && doc.cellText(cells[2]!).trim() === '$') doc.deleteRow(row);
  }
}

/** Table3: target (column 2) and current (column 4) allocation, totals, and the income/growth labels. */
function fillTable3(doc: DocxDocument, r: PerformanceReport | null, targets: Record<AllocationClass, number | null>, notes: string[]) {
  const at = doc.tableAtBookmark('Table3');
  if (!at) {
    notes.push('Bookmark "Table3" (asset allocation) not found in the template.');
    return;
  }
  const { tbl } = at;
  const label = (row: Element) => doc.cellText(doc.rowCells(row)[0]!).trim();
  const textAt = (r: number, c: number) => {
    const x = doc.cellAt(tbl, r, c);
    return x ? doc.cellText(x) : '';
  };
  const classOf = (text: string): AllocationClass | null => ALLOCATION_CLASSES.find((c) => text.includes(c)) ?? (text.includes('Property') ? 'Property' : null);
  for (const row of doc.tableRows(tbl)) {
    const cells = doc.rowCells(row);
    const cls = classOf(label(row));
    if (!cls) continue;
    const t = targets[cls];
    if (t !== null && cells[1]) doc.setCellText(cells[1], pct2(t / 100));
    const cur = r?.current[cls] ?? null;
    // Apex: the property and fixed-interest sums are always written, other classes when in the report
    const apexSums = r?.custodian === 'Apex' && !r.missing.includes('Report Class Current % NZD Amount') && (cls === 'Property' || cls === 'Fixed Interest');
    if (cells[3] && (cur !== null || apexSums)) doc.setCellText(cells[3], pct2(cur ?? 0));
  }
  const empty = (s: string) => s === '0.00%' || s === '%';
  for (const row of [...doc.tableRows(tbl)].reverse()) {
    const cells = doc.rowCells(row);
    if (cells.length >= 4 && empty(doc.cellText(cells[1]!).trim()) && empty(doc.cellText(cells[3]!).trim())) doc.deleteRow(row);
  }
  const rows = doc.tableRows(tbl);
  const n = rows.length;
  let t2 = 0;
  let t4 = 0;
  for (let j = 2; j <= n - 1; j++) {
    t2 += toDecimal(textAt(j, 2));
    t4 += toDecimal(textAt(j, 4));
  }
  const fiRow = rows.findIndex((row) => label(row).includes('Fixed Interest')) + 1;
  for (const [col, total] of [
    [2, t2],
    [4, t4],
  ] as const) {
    const diff = 1 - total;
    let shown = total;
    if (Math.abs(diff) <= 0.0002 && fiRow > 0) {
      const c = doc.cellAt(tbl, fiRow, col)!;
      doc.setCellText(c, pct2(toDecimal(doc.cellText(c)) + diff));
      shown = total + diff;
    }
    const tc = doc.cellAt(tbl, n, col);
    if (tc) doc.setCellText(tc, pct2(shown), Math.abs(shown - 1) > 1e-9 ? { color: RED } : {});
  }
  if (n >= 3) doc.mergeVertical(tbl, 3, 2, 3);
  if (n >= 7) doc.mergeVertical(tbl, 3, 4, 6);
  else if (n >= 5) doc.mergeVertical(tbl, 3, 4, n - 1);
  const c23 = doc.cellAt(tbl, 2, 3);
  if (n >= 2 && c23) {
    doc.setCellText(c23, '+/- 10% Income Assets');
    doc.setCellBorder(c23, 'bottom');
  }
  const c43 = doc.cellAt(tbl, 4, 3);
  if (n >= 4 && c43) doc.setCellText(c43, '+/- 10% Growth Assets');
}

export interface ReviewResult {
  bytes: Uint8Array;
  fileName: string;
  notes: string[];
}

export async function buildAnnualReview(args: {
  template: Uint8Array | ArrayBuffer;
  report: PerformanceReport | null;
  /** Sector targets in percent (Portfolio Cash_Target … Multi_Sector_Target). */
  targets: Record<AllocationClass, number | null>;
  tables: AdviceTables;
  client: LetterClient;
  review: ReviewDetails;
  commentary: Uint8Array | ArrayBuffer | null;
  today: Date;
}): Promise<ReviewResult> {
  const { report, targets, tables, client, review, today } = args;
  const doc = await DocxDocument.open(args.template);
  const notes: string[] = [];
  // ReplaceTextInWord in the ARone module: every occurrence, case-insensitive, text as given
  const rep = (a: string, b: string) => doc.replaceText(a, b, { wholeWord: false, adjustCase: false });

  if (review.authorityToAct && doc.hasBookmark('AuthorityToAct')) doc.deleteBookmark('AuthorityToAct');
  rep('Client_Name', client.fullName);
  rep('Year_Date', monthYear(today));
  rep('Strategy_Name', review.strategyName);
  rep('Insert previous Blurb', review.blurb);
  rep('PrimaryObjective', review.primaryObjective);
  const sigs = review.signatories
    .split(', ')
    .map((s) => s.trim())
    .filter(Boolean);
  for (let i = 1; i <= 4; i++) rep(`Sig${i}`, sigs[i - 1] ?? '');
  for (let i = 2; i <= 4; i++) if (!sigs[i - 1]) doc.deleteBookmarkParagraphWithPrevious(`Sig${i}`);
  rep('I / We', review.iOrWe);
  rep('I am / We are', `${review.iOrWe} ${review.iOrWe === 'I' ? 'am' : 'are'}`);
  rep('my / our', review.iOrWe === 'I' ? 'my' : 'our');
  if (!review.rwt.trim()) {
    rep('rwtXXX', 'XXX');
    doc.highlight('XXX%.');
  } else rep('rwtXXX', review.rwt.trim());
  if (!review.pir.trim()) {
    rep('pirXXX', 'XXX');
    doc.highlight('XXX%.');
  } else rep('pirXXX', review.pir.trim());
  if (!review.accountant.trim()) doc.highlight('accountant');
  else rep('accountant', review.accountant.trim());
  for (const [a, b] of client.clientType === 'Trust' ? HOLDERS_TRUST : HOLDERS_PERSONAL) rep(a, b);

  const sig = doc.signature(client.leadAdviser, client.secondaryAdviser, ['John', 'Hayden', 'Claudia', 'BKMRK'], { widths: [340 * 0.75, 300 * 0.75], indent: 40 * 0.75 });
  if (!sig.built) notes.push('The template has no "AdvisorTable" bookmark, so no signature table was added.');
  for (const m of sig.missing) notes.push(`No signature bookmark "${m}" with a table in the template.`);

  if (report) {
    for (const m of report.missing) notes.push(`"${m}" was not found in the ${report.custodian === 'OA' ? 'OneAnswer' : 'Apex'} report.`);
    fillTable1(doc, report, notes);
    if (report.dollarGain) rep('dollarGain', report.dollarGain);
    if (report.percentGain) rep('percentGain', report.percentGain);
    if (report.closingDate) {
      if (report.openingDate) doc.setBookmarkText('OpeningPV', `Opening Portfolio Value - ${report.openingDate}`);
      doc.setBookmarkText('ClosingPV', `Closing Portfolio Value - ${report.closingDate}`);
      doc.setBookmarkText('EndDate', report.closingDate);
      doc.setBookmarkText('AlloDate', report.closingDate);
      rep('Short_Date', report.closingDate);
      const parts = report.closingDate.split(' ');
      if (parts.length === 3) {
        rep('curYear', parts[2]!);
        rep('Month_Year', `${parts[1]} ${parts[2]}`);
      }
    }
    fillTable2(doc, report, notes);
  } else {
    notes.push('No performance report was pasted, so Table1, Table2 and the period dates were not filled.');
  }
  fillTable3(doc, report, targets, notes);

  if (args.commentary) {
    const other = await DocxDocument.open(args.commentary);
    const ins = doc.insertDocumentAt('OpeningComments', other);
    if (!ins.inserted) notes.push('Bookmark "OpeningComments" not found, so the commentary was not added.');
    else if (ins.droppedPictures) notes.push(`The commentary had ${ins.droppedPictures} picture(s) or object(s), which were not copied; paste them in Word.`);
  }
  doc.updateFieldsOnOpenRequest();

  // ---- AR_Part2: the advice tables
  if (!doc.insertRationaleAt('Investmentrationale', tables.rationale)) notes.push('"Investmentrationale" was not found, so the investment rationale was not added.');
  if (!doc.insertTableAt('Portfolioforadvice', { ...tables.portfolio, ...AR_LAYOUT.portfolio })) notes.push('"Portfolioforadvice" was not found, so the portfolio table was not added.');
  if (tables.costs) {
    if (!doc.insertTableAt('InsertTable', { ...tables.costs, ...AR_LAYOUT.costs })) notes.push('"InsertTable" was not found, so the costs table was not added.');
  } else for (const b of ['Cost1', 'Cost2', 'Cost3']) doc.deleteBookmark(b);
  if (!doc.insertTableAt('ChangeTable', { ...tables.summary, ...AR_LAYOUT.summary })) notes.push('"ChangeTable" was not found, so the summary table was not added.');

  return { bytes: await doc.save(), fileName: `${(client.fileName || client.fullName || 'Client').trim()} Annual Review ${ddmmyy(today)}.docx`, notes };
}

// ---------------------------------------------------------------- Statement of Advice

export async function buildStatementOfAdvice(args: { template: Uint8Array | ArrayBuffer | null; table: DocTable; strategyName: string; clientName: string; today: Date }): Promise<ReviewResult> {
  const notes: string[] = [];
  const doc = args.template ? await DocxDocument.open(args.template) : DocxDocument.blank([{ text: 'Statement of Advice appendix: strategy_name', style: 'title' }, { text: 'SOA_Table' }]);
  doc.replaceText('strategy_name', args.strategyName, { wholeWord: false, adjustCase: false });
  if (!args.template) {
    doc.insertTableAt('SOA_Table', args.table);
  } else if (!doc.insertLandscapeTableAt('SOA_Table', args.table)) {
    notes.push('Bookmark "SOA_Table" not found in the template, so the product table was not added.');
  }
  if (args.table.rows.length <= 1) notes.push('No held or recommended holding was found in the Data sheet, so the table has only its header.');
  return { bytes: await doc.save(), fileName: `SOA ${(args.clientName || 'Client').trim()} Appendices ${ddmmyy(args.today)}.docx`, notes };
}
