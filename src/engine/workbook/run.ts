import type { FirmSettings } from '../settings';
import { Workbook, type WorkbookModel } from './evaluator';
import type { ImportResult } from './importHoldings';
import { applyScenario, readSection, SECTIONS, type RowResult, type SectionKey } from './scenario';
import { applyFirmSettings } from './settingsCells';
import { BLANK, isErr, type Scalar } from './values';

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
