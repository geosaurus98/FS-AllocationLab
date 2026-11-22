import type { FirmSettings } from '../settings';
import { Workbook, type WorkbookModel } from './evaluator';
import type { ImportResult } from './importHoldings';
import { applyScenario, readSection, SECTIONS, type RowResult, type SectionKey } from './scenario';
import { applyFirmSettings } from './settingsCells';
import { BLANK, isErr, type Scalar } from './values';
export type { SectionKey };

/** Run the recommendations workbook for one client. Pure apart from the time taken. */

export const RECOMMENDATION_PROFILES = ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive', 'Custom', 'Current', 'Master Risk Alloc'] as const;
export type RecommendationProfile = (typeof RECOMMENDATION_PROFILES)[number];

export interface ClientInputs {
  profile: RecommendationProfile;
  /** Custom profile: growth share (0–1) and cash share (0–1); fixed income = 1 − growth − cash. */
  custom: { growth: number; cash: number } | null;
  /** Portfolio date (yyyy-mm-dd). */
  portfolioDate: string;
  /** Annual fees and withdrawals cash must cover (Portfolio!BH2), in dollars. */
  feesAndWithdrawals: number;
  addRemove: number | null;
  cashTotalOverride: number | null;
  sectorOverrides: Partial<Record<'fi' | 'ae' | 'ge' | 'prop', number | null>>;
  /** Per-holding overrides (column G), keyed by Portfolio name. */
  overrides: Record<string, number | null>;
  updateSheetButton: boolean;
  /** Custodian (Portfolio!BD1): "Apex" or "OA" — affects broker columns and some lookups. */
  custodian?: 'Apex' | 'OA' | null;
  /** Holdings added during implementation (e.g. the actual bond bought for a "2027 Bonds" row), value 0. */
  addHoldings?: [SectionKey, string][];
}

export interface SectionResult {
  key: SectionKey;
  label: string;
  rows: RowResult[];
  current: Scalar;
  recommended: Scalar;
  target: Scalar;
  targetPct: Scalar;
}

export interface RecommendationResult {
  sections: SectionResult[];
  portfolioTotal: Scalar;
  recommendedTotal: Scalar;
  finalPortfolioTotal: Scalar;
  checks: string[];
  checkStatus: Scalar;
  inserted: Record<SectionKey, string[]>;
  ms: number;
  /** The live workbook, kept so implementation rounds can be updated without a full recalculation. */
  workbook: Workbook;
}

export function runRecommendation(model: WorkbookModel, settings: FirmSettings, imported: ImportResult, client: ClientInputs, today = new Date()): RecommendationResult {
  const t0 = performance.now();
  const wb = new Workbook(model, today);
  applyFirmSettings(wb, settings);
  const holdings = Object.fromEntries(
    Object.entries(imported.bySection).map(([k, rows]) => [k, rows.map(([name, value]) => [name, value, client.overrides[name] ?? null] as [string, number, number | null])]),
  );
  const custom = client.profile === 'Custom' && client.custom ? { lowerRisk: 1 - client.custom.growth, cash: client.custom.cash, fi: 1 - client.custom.growth - client.custom.cash, growth: client.custom.growth } : undefined;
  const { inserted } = applyScenario(wb, {
    profile: client.profile,
    custom,
    holdings,
    feesAndWithdrawals: client.feesAndWithdrawals,
    addRemove: client.addRemove,
    cashTotalOverride: client.cashTotalOverride,
    sectorOverrides: client.sectorOverrides,
    portfolioDate: client.portfolioDate,
    updateSheetButton: client.updateSheetButton,
  });
  if (client.custodian) wb.set('Portfolio', 'BD1', client.custodian);
  for (const [key, name] of client.addHoldings ?? []) {
    const sec = SECTIONS.find((x) => x.key === key)!;
    let exists = false;
    let empty: number | null = null;
    for (let r = sec.first; r <= sec.last; r++) {
      const a = wb.get('Portfolio', r, 1);
      if (a === name) exists = true;
      if (empty === null && a === BLANK) empty = r;
    }
    if (!exists && empty !== null) {
      wb.set('Portfolio', empty, 1, name);
      wb.set('Portfolio', empty, 2, 0);
    }
  }
  // Overrides on rows the workbook added (model assets) are applied after insertion.
  for (const s of SECTIONS) {
    for (let r = s.first; r <= s.last; r++) {
      const a = wb.get('Portfolio', r, 1);
      if (typeof a === 'string' && a in client.overrides && wb.get('Portfolio', r, 2) === 0) wb.set('Portfolio', r, 7, client.overrides[a] ?? null);
    }
  }
  const sections: SectionResult[] = SECTIONS.map((s) => ({
    key: s.key,
    label: s.label,
    rows: readSection(wb, s.key),
    current: wb.get('Portfolio', s.total, 2),
    recommended: wb.get('Portfolio', s.total, 6),
    target: wb.get('Portfolio', s.total, 7),
    targetPct: wb.get('Portfolio', s.total, 4),
  }));
  const checksText = wb.get('Portfolio', 'BD2');
  const checks = typeof checksText === 'string' && checksText ? checksText.split(', ').map((x) => x.trim()).filter(Boolean) : isErr(checksText) ? [`Check formula error ${checksText.code}`] : [];
  return {
    workbook: wb,
    sections,
    portfolioTotal: wb.get('Portfolio', 'B202'),
    recommendedTotal: wb.get('Portfolio', 'F202'),
    finalPortfolioTotal: wb.get('Portfolio', 'F204'),
    checks,
    checkStatus: wb.get('Portfolio', 'BC2'),
    inserted,
    ms: Math.round(performance.now() - t0),
  };
}

export const scalarNumber = (v: Scalar): number | null => (typeof v === 'number' ? v : v === BLANK || v === '' ? null : null);

// ---------------------------------------------------------------- implementation rounds

/** Portfolio round columns: Round 1..6 = AC, AG, AK, AO, AS, AW (each followed by three check columns). */
export const ROUND_COLUMNS = [29, 33, 37, 41, 45, 49];
export const MAX_ROUNDS = ROUND_COLUMNS.length;

/** Round amounts per Portfolio name: amounts[roundIndex]. */
export type RoundAmounts = Record<string, (number | null)[]>;

/** Write round amounts into the workbook and recalculate only what depends on them (Portfolio AC:BN). */
export function applyRounds(wb: Workbook, amounts: RoundAmounts): void {
  const cells: { sheet: string; row: number; col: number; value: number | null }[] = [];
  for (const s of SECTIONS) {
    for (let r = s.first; r <= s.last; r++) {
      const a = wb.get('Portfolio', r, 1);
      const v = typeof a === 'string' ? amounts[a] : undefined;
      ROUND_COLUMNS.forEach((c, i) => cells.push({ sheet: 'Portfolio', row: r, col: c, value: v?.[i] ?? null }));
    }
  }
  wb.setScoped(cells, { sheet: 'Portfolio', c1: 29, c2: 66 });
}

export interface ImplementationRow {
  section: SectionKey;
  row: number;
  name: string;
  current: Scalar;
  recommended: Scalar;
  change: Scalar;
  remaining: Scalar;
  oaBroker: Scalar;
  oaBrokerage: Scalar;
  apexBroker: Scalar;
  apexBrokerage: Scalar;
  buySpread: Scalar;
  sellSpread: Scalar;
  costPa: Scalar;
  rationale: Scalar;
  type: Scalar;
  /** Custodian asset names (Data sheet), used to find the asset when trading. */
  apexName: string;
  oaName: string;
}

export function readImplementation(wb: Workbook): { rows: ImplementationRow[]; checks: string[]; exposure: Record<string, Scalar> } {
  // FS name → Data row, for the custodian names (SafeNameChange)
  const data = new Map<string, { apex: string; oa: string }>();
  const maxData = 1100;
  for (let r = 2; r <= maxData; r++) {
    const fs = wb.get('Data', r, 4);
    if (typeof fs === 'string' && fs && !data.has(fs.toLowerCase())) {
      const ap = wb.get('Data', r, 3);
      const oa = wb.get('Data', r, 2);
      data.set(fs.toLowerCase(), { apex: typeof ap === 'string' ? ap : '', oa: typeof oa === 'string' ? oa : '' });
    }
  }
  const rows: ImplementationRow[] = [];
  for (const s of SECTIONS) {
    for (let r = s.first; r <= s.last; r++) {
      const a = wb.get('Portfolio', r, 1);
      if (typeof a !== 'string' || a === '') continue;
      const g = (c: number) => wb.get('Portfolio', r, c);
      const d = data.get(a.toLowerCase());
      rows.push({
        section: s.key,
        row: r,
        name: a,
        current: g(2),
        recommended: g(6),
        change: g(5),
        remaining: g(53),
        oaBroker: g(12),
        oaBrokerage: g(13),
        apexBroker: g(14),
        apexBrokerage: g(15),
        buySpread: g(16),
        sellSpread: g(17),
        costPa: g(18),
        rationale: g(20),
        type: g(11),
        apexName: d?.apex && d.apex !== 'X' ? d.apex : '',
        oaName: d?.oa && d.oa !== 'X' ? d.oa : '',
      });
    }
  }
  const checksText = wb.get('Portfolio', 'BD2');
  const checks = typeof checksText === 'string' && checksText ? checksText.split(', ').map((x) => x.trim()).filter(Boolean) : [];
  const exposure: Record<string, Scalar> = {};
  for (const [c, bank] of [[61, 'BNZ'], [62, 'ANZ'], [63, 'Westpac'], [64, 'ASB'], [65, 'Kiwibank']] as const) exposure[bank] = wb.get('Portfolio', 55, c);
  return { rows, checks, exposure };
}
