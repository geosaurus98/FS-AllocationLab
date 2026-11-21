/**
 * The tables and paragraphs the advice letters are built from, read from a calculated workbook.
 * Ports of the Portfolio sheet macros that prepare what is pasted into Word:
 *
 *  - Sort_Portfolio            holdings in each section (not Cash) sorted by name
 *  - Portfolio_For_Advice      columns A, B, E, F, H; changed holdings only     → "Portfolioforadvice"
 *  - Costs_For_Advice          column A + the custodian's cost columns           → "InsertTable"
 *  - InvestmentRationale       column T as headings and bullets                  → "Investmentrationale"
 *  - Simple_AR_table           the "Summary Table" sheet                         → "ChangeTable"
 *
 * Widths, font sizes and borders are those the letter macro applies after pasting
 * (PasteIntoWordAtPlaceholder / FormatTableForA4Page).
 *
 * Deliberate differences from the macros (all avoid wrong text in a client document):
 *  - Summary table: the running "section total" is reset after Cash, so the first "All Other …
 *    (No Changes)" row no longer subtracts cash funds from the wrong section; 3-letter names are
 *    treated as cash currency rows only inside the Cash section.
 *  - Summary table: a bond/term deposit whose name does not end in a date is shown as "Sell"
 *    (the macro stops with a type-mismatch error).
 *  - Investment rationale: the stray column-T formula on the "Multi Sector Total" row is ignored.
 *  - Costs: when no holding has a cost line, the Cost1–Cost3 bookmarks are removed. The macro
 *    means to, but its test (last used row in column A = 2) can never be true.
 */

import type { Workbook } from './evaluator';
import { SECTIONS, type SectionKey } from './scenario';
import { BLANK, isErr, parseDateText, type Scalar } from './values';

export interface DocCell {
  text: string;
  bold?: boolean;
  align?: 'left' | 'center';
  /** Font colour, hex RRGGBB. */
  color?: string;
  /** Cell shading, hex RRGGBB. */
  fill?: string;
  /** Hyperlink target for the cell text. */
  link?: string;
  /** Draw a top border on this cell (section breaks in the summary table). */
  topBorder?: boolean;
}

export interface DocTable {
  /** Column widths in points. */
  widths: number[];
  fontSize: number;
  /** Centre the table on the page. */
  center: boolean;
  /** Top border on the first row and bottom border on the last (ApplyBottomBorderToTable). */
  edgeBorders: boolean;
  /** Box and column lines drawn on the Summary Table sheet. */
  summaryBorders?: boolean;
  rows: DocCell[][];
}

export interface RationaleItem {
  kind: 'title' | 'heading' | 'bullet';
  text: string;
  /** Fact sheet address: the letter adds "Link to the fact sheet here." with "here" linked. */
  factSheet?: string;
}

export interface AdviceTables {
  portfolio: DocTable;
  /** `null` when no holding has a cost line (the Cost1–Cost3 bookmarks are then removed). */
  costs: DocTable | null;
  summary: DocTable;
  rationale: RationaleItem[];
}

// ---------------------------------------------------------------- formatting (Excel number formats)

const roundHalfAway = (x: number, dp = 0) => {
  const f = 10 ** dp;
  return (Math.sign(x) * Math.round(Math.abs(x) * f + 1e-9)) / f;
};
const groupDigits = (n: number, dp = 0) => {
  const [i, d] = Math.abs(roundHalfAway(n, dp)).toFixed(dp).split('.');
  return i!.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d ? `.${d}` : '');
};

function asText(v: Scalar): string {
  if (v === BLANK) return '';
  if (isErr(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

/** "$"#,##0 */
export function money(v: Scalar): string {
  if (typeof v !== 'number') return asText(v);
  const r = roundHalfAway(v);
  return `${r < 0 ? '-' : ''}$${groupDigits(r)}`;
}
/** "+$"#,##0;"-$"#,##0;"$"0 */
export function signedMoney(v: Scalar): string {
  if (typeof v !== 'number') return asText(v);
  const r = roundHalfAway(v);
  return r === 0 ? '$0' : `${r < 0 ? '-' : '+'}$${groupDigits(r)}`;
}
/** 0.00 */
export function fixed2(v: Scalar): string {
  if (typeof v !== 'number') return asText(v);
  const r = roundHalfAway(v, 2);
  return `${r < 0 ? '-' : ''}${Math.abs(r).toFixed(2)}`;
}
/** 0.00% */
export function percent2(v: Scalar): string {
  if (typeof v !== 'number') return asText(v);
  return `${fixed2(v * 100)}%`;
}

// ---------------------------------------------------------------- the sorted Portfolio sheet

export interface ViewRow {
  row: number;
  kind: 'header' | 'holding' | 'total' | 'grand';
  section: SectionKey | null;
  name: string;
}

/** Excel's ascending text sort: case-insensitive. */
const excelCompare = (a: string, b: string) => a.localeCompare(b, 'en', { sensitivity: 'base' });

/** Portfolio rows in the order the macros see them after Sort_Portfolio. Blank rows are left out. */
export function portfolioView(wb: Workbook): ViewRow[] {
  const out: ViewRow[] = [];
  for (const s of SECTIONS) {
    const header = s.first - 1;
    out.push({ row: header, kind: 'header', section: s.key, name: asText(wb.get('Portfolio', header, 1)) });
    const rows: ViewRow[] = [];
    for (let r = s.first; r <= s.last; r++) {
      const a = wb.get('Portfolio', r, 1);
      if (typeof a !== 'string' || a === '') continue;
      rows.push({ row: r, kind: 'holding', section: s.key, name: a });
    }
    if (s.key !== 'cash') rows.sort((x, y) => excelCompare(x.name, y.name));
    out.push(...rows);
    out.push({ row: s.total, kind: 'total', section: s.key, name: asText(wb.get('Portfolio', s.total, 1)) });
  }
  out.push({ row: 202, kind: 'grand', section: null, name: asText(wb.get('Portfolio', 202, 1)) });
  return out;
}

const num = (v: Scalar): number | null => (typeof v === 'number' ? v : null);
/** VBA `cell.Value = 0` / `<> 0`: blank counts as 0; text and errors never equal a number. */
const isZero = (v: Scalar) => v === BLANK || v === 0;

function dataLookup(wb: Workbook): Map<string, number> {
  const m = new Map<string, number>();
  for (let r = 2; r <= 1100; r++) {
    const fs = wb.get('Data', r, 4);
    if (typeof fs === 'string' && fs && !m.has(fs.toLowerCase())) m.set(fs.toLowerCase(), r);
  }
  return m;
}

// ---------------------------------------------------------------- Portfolio for Advice

export function portfolioForAdvice(wb: Workbook, view = portfolioView(wb)): DocTable {
  const P = (r: number, c: number) => wb.get('Portfolio', r, c);
  const line = (r: number, bold: boolean): DocCell[] => [
    { text: asText(P(r, 1)), bold, align: 'left' },
    { text: money(P(r, 2)), bold, align: 'center' },
    { text: money(P(r, 5)), bold, align: 'center' },
    { text: money(P(r, 6)), bold, align: 'center' },
    { text: fixed2(P(r, 8)), bold, align: 'center' },
  ];
  const rows: DocCell[][] = [[1, 2, 5, 6, 8].map((c, i) => ({ text: asText(P(1, c)).trim(), bold: true, align: i === 0 ? 'left' : 'center' }))];
  for (const s of SECTIONS) {
    const sec = view.filter((v) => v.section === s.key);
    const shown = sec.filter((v) => {
      if (v.kind !== 'holding') return false;
      const e = P(v.row, 5);
      if (isErr(e) || e === 'N/A' || isZero(e) || e === '') return false;
      return v.name.length !== 3;
    });
    const total = sec.find((v) => v.kind === 'total')!;
    if (shown.length > 0) {
      rows.push(line(s.first - 1, true));
      for (const v of shown) rows.push(line(v.row, false));
      rows.push(line(total.row, true));
    } else if (!(isZero(P(total.row, 2)) && isZero(P(total.row, 6)))) {
      rows.push(line(total.row, true));
    }
  }
  rows.push(line(202, true));
  return { widths: [170, 81, 81, 81, 71], fontSize: 10, center: true, edgeBorders: false, rows };
}

// ---------------------------------------------------------------- Costs for Advice

export function costsForAdvice(wb: Workbook, custodian: 'Apex' | 'OA', view = portfolioView(wb)): DocTable | null {
  const P = (r: number, c: number) => wb.get('Portfolio', r, c);
  const data = dataLookup(wb);
  const cols = custodian === 'Apex' ? [14, 15, 16, 17, 18, 19] : [12, 13, 16, 17, 18, 19];
  const fmt = (c: number, v: Scalar) => (c === 12 || c === 14 ? fixed2(v) : c === 19 ? asText(v) : percent2(v));
  const header: DocCell[] = [{ text: '', bold: true, align: 'left' }, ...cols.map((c): DocCell => ({ text: asText(P(2, c)).trim(), bold: true, align: 'center' }))];
  const rows: DocCell[][] = [header];
  for (const v of view) {
    if (v.kind !== 'holding') continue;
    const e = P(v.row, 5);
    const type = P(v.row, 10);
    if (isErr(e) || isZero(e) || e === '' || v.name.length === 3) continue;
    if (type === 'Term Deposit' || (type === 'Bond' && typeof e === 'number' && e < 0)) continue;
    const dr = data.get(v.name.toLowerCase());
    const pds = dr ? wb.get('Data', dr, 20) : BLANK;
    rows.push([
      { text: v.name, align: 'left' },
      ...cols.map((c): DocCell => {
        const val = P(v.row, c);
        const cell: DocCell = { text: fmt(c, val), align: 'center' };
        if (c === 19 && val === 'Click here' && typeof pds === 'string' && pds && pds !== '-') cell.link = pds;
        return cell;
      }),
    ]);
  }
  if (rows.length === 1) return null;
  return { widths: [163, 68, 69, 50, 50, 50, 75], fontSize: 8, center: true, edgeBorders: true, rows };
}

// ---------------------------------------------------------------- Investment rationale

export function investmentRationale(wb: Workbook, view = portfolioView(wb)): RationaleItem[] {
  const P = (r: number, c: number) => wb.get('Portfolio', r, c);
  const data = dataLookup(wb);
  const out: RationaleItem[] = [];
  const title = asText(P(11, 20)).trim();
  if (title) out.push({ kind: 'title', text: title });
  for (const s of SECTIONS) {
    if (s.key === 'cash') continue; // rows 1 to the end of the cash holdings are hidden
    const items: RationaleItem[] = [];
    for (const v of view) {
      if (v.section !== s.key || v.kind !== 'holding') continue;
      const t = P(v.row, 20);
      if (isErr(t)) continue;
      const text = asText(t).trim();
      if (!text) continue;
      const dr = data.get(v.name.toLowerCase());
      const fact = dr ? wb.get('Data', dr, 29) : BLANK;
      const item: RationaleItem = { kind: 'bullet', text };
      if (typeof fact === 'string' && fact !== '' && fact !== '-') item.factSheet = fact;
      items.push(item);
    }
    if (items.length === 0) continue;
    out.push({ kind: 'heading', text: asText(P(s.first - 1, 20)).trim() || s.label }, ...items);
  }
  return out;
}

// ---------------------------------------------------------------- Summary table (Simple_AR_table)

const NAVY = '002060';
const BLUE = 'C5D9F1';
const PALE = 'DCE6F1';
const WHITE = 'FFFFFF';
const GREEN = '00B050';
const RED = 'FF0000';

interface SumRow {
  a: string;
  b: string;
  c: number | null;
  d: number | null;
  kind: 'header' | 'holding' | 'other' | 'total';
}

function maturityAction(wb: Workbook, name: string): 'Maturity' | 'Sell' {
  const serial = parseDateText(name.slice(-10));
  if (serial === null) return 'Sell';
  const limit = num(wb.get('Portfolio', 'BD14')) ?? 0;
  return Math.abs(serial - wb.today) <= limit ? 'Maturity' : 'Sell';
}

export function summaryTable(wb: Workbook, clientName: string, view = portfolioView(wb)): DocTable {
  const P = (r: number, c: number) => wb.get('Portfolio', r, c);
  const E = (r: number) => num(P(r, 5)) ?? 0;
  const F = (r: number) => num(P(r, 6)) ?? 0;
  const out: SumRow[] = [];
  let totalSum = 0;
  let totalRec = 0;

  for (const s of SECTIONS) {
    const sec = view.filter((v) => v.section === s.key);
    const holdings = sec.filter((v) => v.kind === 'holding');
    const total = sec.find((v) => v.kind === 'total')!;
    const listed: SumRow[] = [];
    let sectionTotal = 0;
    const action = (v: ViewRow, c: number, d: number): string => {
      if (c > 0) return d === c ? 'Buy' : 'Add';
      if (c < 0) {
        if (d !== 0) return 'Trim';
        const type = P(v.row, 10);
        const k = P(v.row, 11);
        if ((type === 'Bond' || type === 'Term Deposit') && k !== 'Perpetual' && !v.name.includes('subject to liquidity')) return maturityAction(wb, v.name);
        return 'Sell';
      }
      return '';
    };

    if (s.key === 'cash') {
      let cashVal = 0;
      let cashTotal = 0;
      let assetTot = 0;
      let currencyRows = false;
      let cashAssets = false;
      for (const v of holdings) {
        const e = P(v.row, 5);
        if (e === BLANK || typeof e !== 'number') continue;
        if (v.name.length === 3) {
          cashVal += e;
          cashTotal += e;
          currencyRows = true;
          continue;
        }
        if (e === 0) continue;
        if (P(v.row, 10) === 'Cash') {
          cashTotal += e;
          assetTot += F(v.row);
          cashAssets = true;
        }
        out.push({ a: v.name, b: action(v, e, F(v.row)), c: e, d: F(v.row), kind: 'holding' });
      }
      if (currencyRows) {
        if (cashAssets) out.push({ a: 'All Other Cash', b: cashVal > 0 ? 'Add' : 'Trim', c: cashVal, d: F(total.row) - assetTot, kind: 'other' });
        out.push({ a: 'Cash Total', b: cashVal > 0 ? 'Increase' : cashVal === 0 ? 'Hold' : 'Reduce', c: cashTotal, d: F(total.row), kind: 'total' });
        totalSum += cashTotal;
        totalRec += F(total.row);
      }
      continue;
    }

    const hasData = [...holdings, total].some((v) => E(v.row) !== 0);
    const hasVals = !hasData && [...holdings, total].some((v) => F(v.row) !== 0);
    for (const v of holdings) {
      const e = E(v.row);
      if (e === 0) continue;
      sectionTotal += F(v.row);
      listed.push({ a: v.name, b: action(v, e, F(v.row)), c: e, d: F(v.row), kind: 'holding' });
    }
    if (hasData) {
      // largest increase first, then buys and sells each in name order
      const buys = listed.filter((x) => x.c! > 0).sort((x, y) => excelCompare(x.a, y.a));
      const sells = listed.filter((x) => x.c! < 0).sort((x, y) => excelCompare(x.a, y.a));
      out.push({ a: asText(P(s.first - 1, 1)), b: '', c: null, d: null, kind: 'header' });
      out.push(...buys, ...sells);
      out.push({ a: `All Other ${asText(P(s.first - 1, 1))} (No Changes)`, b: '', c: 0, d: F(total.row) - sectionTotal, kind: 'other' });
      out.push({ a: total.name, b: '', c: E(total.row), d: F(total.row), kind: 'total' });
      totalRec += F(total.row);
      totalSum += E(total.row);
    } else if (hasVals) {
      out.push({ a: total.name, b: '', c: E(total.row), d: F(total.row), kind: 'total' });
      totalRec += F(total.row);
      totalSum += E(total.row);
    }
  }

  const addRemove = P(11, 56); // BD11 Add_remove_to_Portfolio
  const rows: DocCell[][] = [
    [clientName, 'Action', 'Changes', 'Recomm. Portfolio'].map((t) => ({ text: t, bold: true, align: 'center' as const, color: WHITE, fill: NAVY })),
    [
      { text: '', fill: NAVY },
      { text: '', fill: NAVY },
      { text: '', fill: NAVY },
      addRemove === BLANK || addRemove === '' ? { text: '', fill: NAVY } : { text: signedMoney(addRemove), bold: true, align: 'center', color: WHITE, fill: NAVY },
    ],
  ];
  const actionColor = (b: string) => (['Buy', 'Add', 'Increase', 'Hold'].includes(b) ? GREEN : RED);
  for (const r of out) {
    const bold = r.kind === 'header' || r.kind === 'total';
    const top = r.kind === 'header' || r.kind === 'total';
    rows.push([
      { text: r.a, bold, align: 'left', fill: r.kind === 'holding' || r.kind === 'other' ? PALE : WHITE, topBorder: top },
      { text: r.b, align: 'center', color: r.b ? actionColor(r.b) : undefined, fill: WHITE, topBorder: top },
      { text: r.c === null ? '' : money(r.c), bold, align: 'center', fill: BLUE, topBorder: top },
      { text: r.d === null ? '' : money(r.d), bold, align: 'center', fill: WHITE, topBorder: top },
    ]);
  }
  rows.push([
    { text: 'Totals', bold: true, align: 'left', fill: WHITE, topBorder: true },
    { text: '', fill: WHITE, topBorder: true },
    { text: money(totalSum), bold: true, align: 'center', fill: BLUE, topBorder: true },
    { text: money(totalRec), bold: true, align: 'center', fill: WHITE, topBorder: true },
  ]);
  return { widths: [300, 50, 70, 80], fontSize: 10, center: false, edgeBorders: true, summaryBorders: true, rows };
}

export function adviceTables(wb: Workbook, opts: { clientName: string; custodian: 'Apex' | 'OA' }): AdviceTables {
  const view = portfolioView(wb);
  return {
    portfolio: portfolioForAdvice(wb, view),
    costs: costsForAdvice(wb, opts.custodian, view),
    summary: summaryTable(wb, opts.clientName, view),
    rationale: investmentRationale(wb, view),
  };
}
