import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Workbook, type WorkbookModel } from './evaluator';
import { applyRounds } from './run';
import { applyScenario, SECTIONS, type Scenario, type SectionKey } from './scenario';
import { BLANK, isErr, type Scalar } from './values';

/**
 * Implementation rounds against the real workbook (scripts/golden/gen_rounds.py + run_workbook.py):
 * Remaining (BA), broker/brokerage/spread/rationale columns, the bank-exposure table and the checks.
 * Round amounts are applied with the scoped recalculation the app uses.
 */

type Cell = number | string | null | { error: string };
interface Fixture {
  scenario: Scenario & { id: string; custodian?: 'Apex' | 'OA'; addHoldings?: [SectionKey, string][]; rounds?: Record<string, number[]> };
  rows: Record<string, Record<string, Cell>[]>;
  implementation: Record<string, Cell>;
  checks: string;
}

const file = resolve(__dirname, '__golden_rounds__/rounds.json');
const model: WorkbookModel = JSON.parse(readFileSync(resolve(__dirname, 'model.json'), 'utf8'));
const fixtures: Fixture[] = existsSync(file) ? JSON.parse(readFileSync(file, 'utf8')) : [];

function same(got: Scalar, exp: Cell): boolean {
  if (exp !== null && typeof exp === 'object') return isErr(got) && got.code === exp.error;
  if (typeof exp === 'number') return typeof got === 'number' && Math.abs(got - exp) <= 1e-6 * Math.max(1, Math.abs(exp));
  if (exp === null || exp === '') return got === BLANK || got === '';
  return got === exp;
}

const COLS: Record<string, number> = { E: 5, F: 6, L: 12, M: 13, N: 14, O: 15, P: 16, Q: 17, R: 18, T: 20, BA: 53 };

describe('implementation rounds match the workbook', () => {
  it('has fixtures', () => expect(fixtures.length).toBeGreaterThanOrEqual(10));
  for (const f of fixtures) {
    it(f.scenario.id, () => {
      const wb = new Workbook(model, new Date(2026, 10, 20));
      wb.libreOfficeCompat = true;
      applyScenario(wb, f.scenario);
      if (f.scenario.custodian) wb.set('Portfolio', 'BD1', f.scenario.custodian);
      for (const [key, name] of f.scenario.addHoldings ?? []) {
        const s = SECTIONS.find((x) => x.key === key)!;
        let exists = false;
        for (let r = s.first; r <= s.last; r++) if (wb.get('Portfolio', r, 1) === name) exists = true;
        if (exists) continue;
        for (let r = s.first; r <= s.last; r++) {
          if (wb.get('Portfolio', r, 1) === BLANK) {
            wb.set('Portfolio', r, 1, name);
            wb.set('Portfolio', r, 2, 0);
            break;
          }
        }
      }
      // warm the cache first so the scoped update is exercised
      for (const s of SECTIONS) wb.get('Portfolio', s.total, 6);
      applyRounds(wb, f.scenario.rounds ?? {});
      const diffs: string[] = [];
      for (const s of SECTIONS) {
        const exp = f.rows[s.key] ?? [];
        const names: string[] = [];
        for (let r = s.first; r <= s.last; r++) {
          const a = wb.get('Portfolio', r, 1);
          if (typeof a === 'string' && a) names.push(a);
        }
        if (names.join('|') !== exp.map((e) => e.A).join('|')) {
          diffs.push(`${s.key} rows differ`);
          continue;
        }
        exp.forEach((e, i) => {
          const r = s.first + i;
          for (const [col, c] of Object.entries(COLS)) {
            const got = wb.get('Portfolio', r, c);
            if (!same(got, e[col] ?? null)) diffs.push(`${s.key} ${e.A} ${col}: ${isErr(got) ? got.code : String(got)} vs ${JSON.stringify(e[col])}`);
          }
        });
      }
      for (const [a1, exp] of Object.entries(f.implementation ?? {})) {
        const got = wb.get('Portfolio', a1);
        if (!same(got, exp)) diffs.push(`${a1}: ${String(got)} vs ${JSON.stringify(exp)}`);
      }
      if (wb.get('Portfolio', 'BD2') !== f.checks) diffs.push(`checks: ${String(wb.get('Portfolio', 'BD2'))} vs ${f.checks}`);
      expect(diffs.slice(0, 12)).toEqual([]);
    });
  }
});
