import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { isErr, type XValue } from './excel';
import { currentSectorPct, defaultTargetsPct, profileColumn, WORKBOOK_TACTICAL_TABLE, type WorkbookProfile } from './profiles';
import { sectorTargets } from './sectorTargets';

/**
 * Golden tests: expected values were produced by the real recommendations workbook, recalculated in
 * LibreOffice from synthetic scenarios (scripts/golden/gen_sector_targets.py + run_workbook.py).
 */

type Key = 'cash' | 'fi' | 'ae' | 'ge' | 'prop' | 'ms';
interface Fixture {
  scenario: {
    id: string;
    profile: WorkbookProfile;
    custom?: { lowerRisk: number; cash: number; fi: number; growth: number };
    targetsPct?: Partial<Record<Key, number | null>>;
    holdings: Partial<Record<Key, [string, number][]>>;
    feesAndWithdrawals?: number | null;
    addRemove?: number | null;
    sectorOverrides?: Partial<Record<'fi' | 'ae' | 'ge' | 'prop', number | null>>;
  };
  workbook: { sectors: Record<Key, { G: number | { error: string } }>; portfolio: { finalPortfolioTotal: number } };
}

const fixtures: Fixture[] = JSON.parse(readFileSync(resolve(__dirname, '__golden__/sector-targets.json'), 'utf8'));

function same(actual: XValue, expected: number | { error: string }) {
  if (typeof expected === 'object') {
    expect(isErr(actual) ? actual.code : actual).toBe(expected.error);
  } else {
    expect(isErr(actual)).toBe(false);
    expect(actual as number).toBeCloseTo(expected, 6);
  }
}

describe('sector targets match the workbook', () => {
  it('has fixtures', () => expect(fixtures.length).toBeGreaterThanOrEqual(30));

  for (const f of fixtures) {
    const s = f.scenario;
    it(`${s.id} (${s.profile})`, () => {
      const totals = Object.fromEntries((['cash', 'fi', 'ae', 'ge', 'prop', 'ms'] as Key[]).map((k) => [k, (s.holdings[k] ?? []).reduce((a, h) => a + h[1], 0)])) as Record<Key, number>;
      const P = Object.values(totals).reduce((a, b) => a + b, 0);
      const table = s.custom ? { ...WORKBOOK_TACTICAL_TABLE, Custom: { ...s.custom, ms: 0 } } : WORKBOOK_TACTICAL_TABLE;
      const t = profileColumn(table, s.profile, currentSectorPct(totals));
      const d = { ...defaultTargetsPct(t), ...(s.targetsPct ?? {}) };
      const ov = s.sectorOverrides ?? {};
      const r = sectorTargets({
        t,
        cashTargetPct: d.cash ?? null,
        fiTargetPct: d.fi ?? null,
        msTargetPct: d.ms ?? null,
        portfolioTotal: P,
        feesAndWithdrawals: s.feesAndWithdrawals ?? null,
        addRemove: s.addRemove ?? null,
        cashTotalOverride: null,
        fiOverride: ov.fi ?? null,
        aeOverride: ov.ae ?? null,
        geOverride: ov.ge ?? null,
        propOverride: ov.prop ?? null,
        current: { fi: totals.fi, ae: totals.ae, ge: totals.ge, prop: totals.prop },
      });
      for (const k of ['cash', 'fi', 'ae', 'ge', 'prop', 'ms'] as Key[]) same(r[k], f.workbook.sectors[k].G);
      expect(r.finalPortfolioTotal).toBeCloseTo(f.workbook.portfolio.finalPortfolioTotal, 6);
    });
  }
});
