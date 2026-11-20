import { Workbook } from './evaluator';
import { BLANK, isErr, type Scalar } from './values';

/**
 * Drive the workbook model the way an adviser (and the workbook's macros) would:
 *   1. import holdings into the Portfolio sections and set Portfolio_Total (custodian import macros);
 *   2. select the risk profile, which copies the profile's percentages into column D (ComboBox1_Change);
 *   3. press "Update Sheet", which adds model assets with a non-zero amount (Trigger_Recalculate).
 */

export type SectionKey = 'cash' | 'fi' | 'ae' | 'ge' | 'prop' | 'ms';

export const SECTIONS: { key: SectionKey; label: string; first: number; last: number; total: number }[] = [
  { key: 'cash', label: 'Cash', first: 4, last: 10, total: 11 },
  { key: 'fi', label: 'Fixed Interest', first: 13, last: 71, total: 72 },
  { key: 'ae', label: 'Australasian Equities', first: 74, last: 119, total: 120 },
  { key: 'ge', label: 'Global Equities', first: 122, last: 167, total: 168 },
  { key: 'prop', label: 'Property', first: 170, last: 191, total: 192 },
  { key: 'ms', label: 'Multi Sector', first: 194, last: 200, total: 201 },
];

const TARGET_ROWS: Record<SectionKey, number> = { cash: 4, fi: 5, ae: 7, ge: 8, prop: 9, ms: 10 }; // Asset Allocation rows
const TARGET_CELLS: Record<SectionKey, string> = { cash: 'D11', fi: 'D72', ae: 'D120', ge: 'D168', prop: 'D192', ms: 'D201' };
const SECTOR_OVERRIDE_CELLS = { fi: 'BD72', ae: 'BD120', ge: 'BD168', prop: 'BD192' } as const;

/** [name, value, override?] */
export type HoldingInput = [string, number, (number | null)?];

export interface Scenario {
  profile: string;
  custom?: { lowerRisk: number; cash: number; fi: number; growth: number };
  targetsPct?: Partial<Record<SectionKey, number | null>>;
  holdings: Partial<Record<SectionKey, HoldingInput[]>>;
  portfolioTotal?: number | null;
  feesAndWithdrawals?: number | null;
  addRemove?: number | null;
  cashTotalOverride?: number | null;
  sectorOverrides?: Partial<Record<'fi' | 'ae' | 'ge' | 'prop', number | null>>;
  portfolioDate?: string;
  insertModelAssets?: boolean;
  /** Run the import-time macros (AdjustPortfolio, Gold_distribution, ResetInvestmentsMaturing). Default true. */
  importMacros?: boolean;
  /** Also press the "Update Sheet" button (Trigger_Recalculate.Recalculate). */
  updateSheetButton?: boolean;
}

const PROFILE_COLS = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

export function applyScenario(wb: Workbook, sc: Scenario): { inserted: Record<SectionKey, string[]> } {
  const AA = 'Asset Allocation';
  const P = 'Portfolio';
  if (sc.custom) {
    wb.set(AA, 'H3', sc.custom.lowerRisk);
    wb.set(AA, 'H4', sc.custom.cash);
    wb.set(AA, 'H5', sc.custom.fi);
    wb.set(AA, 'H6', sc.custom.growth);
  }
  let total = 0;
  for (const s of SECTIONS) {
    const rows = sc.holdings[s.key] ?? [];
    if (rows.length > s.last - s.first + 1) throw new Error(`Too many ${s.label} holdings`);
    for (let i = 0; i <= s.last - s.first; i++) {
      const r = s.first + i;
      const h = rows[i];
      wb.set(P, r, 1, h ? h[0] : null);
      wb.set(P, r, 2, h ? h[1] : null);
      wb.set(P, r, 7, h && h[2] !== undefined ? h[2] : null);
      if (h) total += h[1];
    }
  }
  wb.set(P, 'B204', sc.portfolioTotal ?? total);
  wb.set(P, 'BH2', sc.feesAndWithdrawals ?? 0);
  wb.set(P, 'BD11', sc.addRemove ?? null);
  wb.set(P, 'BD12', sc.cashTotalOverride ?? null);
  for (const [k, cell] of Object.entries(SECTOR_OVERRIDE_CELLS)) wb.set(P, cell, sc.sectorOverrides?.[k as keyof typeof SECTOR_OVERRIDE_CELLS] ?? null);
  wb.set(P, 'B2', sc.portfolioDate ?? '2026-11-20');
  if (sc.importMacros !== false) importMacros(wb);

  // 2. ComboBox1_Change: profile column x 100 into column D.
  const headers = PROFILE_COLS.map((c) => wb.get(AA, `${c}2`));
  const idx = headers.findIndex((h) => h === sc.profile);
  if (idx < 0) throw new Error(`Unknown profile ${sc.profile}`);
  const col = PROFILE_COLS[idx]!;
  const targets = Object.fromEntries((Object.keys(TARGET_ROWS) as SectionKey[]).map((k) => [k, wb.get(AA, `${col}${TARGET_ROWS[k]}`)])) as Record<SectionKey, Scalar>;
  wb.set(P, 'BE1', sc.profile);
  for (const k of Object.keys(TARGET_CELLS) as SectionKey[]) {
    const override = sc.targetsPct && k in sc.targetsPct ? sc.targetsPct[k] : undefined;
    const t = targets[k];
    wb.set(P, TARGET_CELLS[k], override !== undefined ? override : typeof t === 'number' ? t * 100 : 0);
  }

  // 3. Asset Allocation Worksheet_Calculate event — add model rows, until nothing more is added.
  const inserted: Record<SectionKey, string[]> = { cash: [], fi: [], ae: [], ge: [], prop: [], ms: [] };
  if (sc.insertModelAssets !== false) {
    for (let i = 0; i < 10; i++) if (!worksheetCalculateInsert(wb, inserted)) break;
  }
  // 4. Optional "Update Sheet" button (Trigger_Recalculate.Recalculate).
  if (sc.updateSheetButton) updateSheetButton(wb, inserted);
  return { inserted };
}

export interface RowResult {
  row: number;
  name: string;
  value: Scalar;
  override: Scalar;
  recommended: Scalar;
  change: Scalar;
  recommendedPct: Scalar;
  info: Scalar;
  type: Scalar;
}

export function readSection(wb: Workbook, key: SectionKey): RowResult[] {
  const s = SECTIONS.find((x) => x.key === key)!;
  const out: RowResult[] = [];
  for (let r = s.first; r <= s.last; r++) {
    const a = wb.get('Portfolio', r, 1);
    if (a === BLANK || a === '' || a === 0) continue;
    out.push({
      row: r,
      name: String(a),
      value: wb.get('Portfolio', r, 2),
      override: wb.get('Portfolio', r, 7),
      recommended: wb.get('Portfolio', r, 6),
      change: wb.get('Portfolio', r, 5),
      recommendedPct: wb.get('Portfolio', r, 8),
      info: wb.get('Portfolio', r, 10),
      type: wb.get('Portfolio', r, 11),
    });
  }
  return out;
}

// ---------------------------------------------------------------- macro emulation

function insertRow(wb: Workbook, key: SectionKey, name: string, inserted: Record<SectionKey, string[]>): boolean {
  const s = SECTIONS.find((x) => x.key === key)!;
  for (let r = s.first; r <= s.last; r++) if (wb.get('Portfolio', r, 1) === name) return false;
  for (let r = s.first; r <= s.last; r++) {
    const a = wb.get('Portfolio', r, 1);
    if (a === BLANK) {
      wb.set('Portfolio', r, 1, name);
      wb.set('Portfolio', r, 2, 0);
      inserted[key].push(name);
      return true;
    }
  }
  return false;
}

class Abort extends Error {}

/** VBA Value2 used in a numeric comparison: error/text values raise (type mismatch ends the handler). */
function num(v: Scalar): number | '' {
  if (isErr(v)) throw new Abort();
  if (v === BLANK) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? -1 : 0;
  if (v === '') return '';
  throw new Abort();
}

/** Port of the Asset Allocation sheet's Worksheet_Calculate handler (one pass, live recalculation). */
export function worksheetCalculateInsert(wb: Workbook, inserted: Record<SectionKey, string[]>): boolean {
  const AA = 'Asset Allocation';
  let any = false;
  try {
    // ProcessBondsAndTDs — BondsTDs = H20:I27
    for (let r = 20; r <= 27; r++) {
      for (const [c, kind] of [[8, 'bonds'], [9, 'tds']] as const) {
        const v = wb.get(AA, r, c);
        if (isErr(v)) continue;
        const x = num(v);
        if (x === '' || x === 0) continue;
        const yearCell = wb.get(AA, r, 2);
        const year = typeof yearCell === 'number' ? String(yearCell) : String(yearCell === BLANK ? '' : yearCell);
        const desc = kind === 'bonds' ? `${year} Bonds (BBB+ or higher, subject to liquidity)` : `Term Deposits ${year}`;
        any = insertRow(wb, 'fi', desc, inserted) || any;
      }
    }
    // ProcessProperties — Properties = B119:D149
    for (let r = 119; r <= 149; r++) {
      const amt = wb.get(AA, r, 4);
      if (isErr(amt)) continue;
      const name = wb.get(AA, r, 2);
      if (typeof name === 'string' && name !== '') {
        const a = num(amt);
        if (a !== '' && a > 0) any = insertRow(wb, 'prop', name, inserted) || any;
      }
    }
    // ProcessGlobalEquities — Global_Equities = B79:G114
    const limit = wb.get(AA, 'Q75') as number;
    for (let r = 79; r <= 114; r++) {
      const d = wb.get(AA, r, 4);
      if (isErr(d)) continue;
      const dv = num(d);
      if (dv === '' || !(dv > 0)) continue;
      const gv = num(wb.get(AA, r, 7));
      if (gv === '') continue;
      const mv = num(wb.get(AA, r, 13));
      if ((mv === '' ? 0 : mv) < 2 && gv > limit) {
        const name = wb.get(AA, r, 2);
        const n = typeof name === 'string' ? name : '';
        if (n !== 'Active' && n !== 'Direct' && n !== 'Passive') any = insertRow(wb, 'ge', n, inserted) || any;
      }
    }
    // ProcessAustralasianEquities — AA_Australasian_Equities = B35:B70
    for (let r = 35; r <= 70; r++) {
      const t = wb.get(AA, r, 2);
      const txt = typeof t === 'string' ? t : t === BLANK ? '' : String(t);
      if (txt && !txt.toLowerCase().includes('direct') && txt.trim().split(' ').length > 1) {
        const dv = num(wb.get(AA, r, 4));
        if (dv !== '' && dv !== 0) any = insertRow(wb, 'ae', txt, inserted) || any;
      }
    }
  } catch (e) {
    if (!(e instanceof Abort)) throw e; // a run-time error ends the handler
  }
  return any;
}

/** Port of Trigger_Recalculate.Recalculate ("Update Sheet" button). */
export function updateSheetButton(wb: Workbook, inserted: Record<SectionKey, string[]>): void {
  const AA = 'Asset Allocation';
  const plan: [SectionKey, [number, number, number][]][] = [
    ['ae', [[2, 35, 70]]],
    ['ge', [[2, 79, 114], [10, 73, 77]]],
    ['prop', [[2, 119, 149]]],
  ];
  const toAdd: [SectionKey, string][] = [];
  for (const [key, ranges] of plan) {
    for (const [c, r1, r2] of ranges) {
      for (let r = r1; r <= r2; r++) {
        const v = wb.get(AA, r, c);
        if (typeof v !== 'string' || v === '' || v.includes('Direct')) continue;
        if (v.trim().split(' ').length <= 1) continue;
        const amt = wb.get(AA, r, c + 2);
        if (isErr(amt) || amt === 0 || amt === BLANK || amt === '' || amt === false) continue;
        toAdd.push([key, v.trim()]);
      }
    }
  }
  for (const [key, name] of toAdd) insertRow(wb, key, name, inserted);
}

/** Steps of PortfolioSetup that change inputs: AdjustPortfolio, Gold_distribution, ResetInvestmentsMaturing. */
export function importMacros(wb: Workbook): void {
  const P = 'Portfolio';
  const text = (v: Scalar) => (typeof v === 'string' ? v : v === BLANK ? '' : String(v));
  const lastPortfolioRow = () => {
    for (let r = 259; r >= 1; r--) if (text(wb.get(P, r, 1)) !== '') return r;
    return 1;
  };
  // AdjustPortfolio: Data rows flagged "Add" (column W) are added to their section with value 0.
  let lastData = 1;
  for (let r = 1; r < 1200; r++) if (text(wb.get('Data', r, 23)) !== '') lastData = r;
  for (let i = 2; i <= lastData; i++) {
    if (text(wb.get('Data', i, 23)) !== 'Add') continue;
    const name = text(wb.get('Data', i, 4));
    const section = text(wb.get('Data', i, 5));
    const lastp = lastPortfolioRow();
    let header: number | null = null;
    for (let r = 1; r <= lastp; r++) if (text(wb.get(P, r, 1)).toLowerCase() === section.toLowerCase()) { header = r; break; }
    if (header === null) continue;
    let exists = false;
    for (let r = header; r <= lastp; r++) if (text(wb.get(P, r, 1)).toLowerCase() === name.toLowerCase()) { exists = true; break; }
    if (exists) continue;
    let r = header;
    while (text(wb.get(P, r, 1)) !== '') r++;
    wb.set(P, r, 1, name);
    wb.set(P, r, 2, 0);
  }
  // Gold_distribution: K74:K77 from the gold assets held.
  const gold = [74, 75, 76, 77].map((r) => text(wb.get('Asset Allocation', r, 10)));
  const held = new Map<string, number>();
  let total = 0;
  const lastp = lastPortfolioRow();
  for (let r = 1; r <= lastp; r++) {
    const a = text(wb.get(P, r, 1));
    if (a && gold.includes(a) && !held.has(a)) {
      const b = wb.get(P, r, 2);
      const v = typeof b === 'number' ? b : 0;
      held.set(a, v);
      total += v;
    }
  }
  gold.forEach((g, i) => {
    let v = 0;
    if (held.has(g)) v = total !== 0 ? held.get(g)! / total : g === 'Sprott Physical Gold Trust ETV' ? 1 : 0;
    wb.set('Asset Allocation', 74 + i, 11, v);
  });
  // ResetInvestmentsMaturing
  wb.set(P, 'BD14', 90);
}
