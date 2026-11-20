import { readdirSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Workbook, type WorkbookModel } from './evaluator';
import { applyScenario, readSection, SECTIONS, type Scenario } from './scenario';
import { BLANK, isErr, type Scalar } from './values';

/**
 * The evaluator against the real workbook recalculated in LibreOffice (scripts/golden/run_workbook.py).
 * `mode` records how each fixture set was generated:
 *   targets — import + profile only;  button — "Update Sheet" button emulation only;
 *   full    — import macros + Worksheet_Calculate insertion (what the app does).
 */

const model: WorkbookModel = JSON.parse(readFileSync(resolve(__dirname, 'model.json'), 'utf8'));
type Cell = number | string | null | { error: string };
interface Fixture {
  mode: 'targets' | 'button' | 'full';
  scenarios: { scenario: Scenario & { id: string }; rows: Record<string, [string, Cell][]>; sectors: Record<string, [Cell, Cell]>; checks: string }[];
}

function same(got: Scalar, exp: Cell): boolean {
  if (exp !== null && typeof exp === 'object') return isErr(got) && got.code === exp.error;
  if (typeof exp === 'number') return typeof got === 'number' && Math.abs(got - exp) <= 1e-6 * Math.max(1, Math.abs(exp));
  if (exp === null || exp === '') return got === BLANK || got === '';
  return got === exp;
}

const dir = resolve(__dirname, '__golden__');
for (const file of readdirSync(dir).filter((f) => f.endsWith('.json'))) {
  const fx: Fixture = JSON.parse(readFileSync(resolve(dir, file), 'utf8'));
  describe(`workbook golden: ${file} (${fx.mode})`, () => {
    for (const s of fx.scenarios) {
      it(s.scenario.id, () => {
        const wb = new Workbook(model, new Date(2026, 10, 20));
        wb.libreOfficeCompat = true;
        const sc: Scenario =
          fx.mode === 'targets' ? { ...s.scenario, importMacros: false, insertModelAssets: false } : fx.mode === 'button' ? { ...s.scenario, importMacros: false, insertModelAssets: false, updateSheetButton: true } : s.scenario;
        applyScenario(wb, sc);
        const diffs: string[] = [];
        for (const sec of SECTIONS) {
          const exp = s.rows[sec.key] ?? [];
          const got = readSection(wb, sec.key);
          if (got.map((r) => r.name).join('|') !== exp.map((r) => r[0]).join('|')) diffs.push(`${sec.key} rows: ${got.map((r) => r.name).join('|')} vs ${exp.map((r) => r[0]).join('|')}`);
          else exp.forEach(([name, f], i) => !same(got[i]!.recommended, f) && diffs.push(`${sec.key} ${name}: ${String(got[i]!.recommended)} vs ${JSON.stringify(f)}`));
          const [ef, eg] = s.sectors[sec.key]!;
          if (!same(wb.get('Portfolio', sec.total, 6), ef)) diffs.push(`${sec.key} total F`);
          if (!same(wb.get('Portfolio', sec.total, 7), eg)) diffs.push(`${sec.key} target G`);
        }
        if (wb.get('Portfolio', 'BD2') !== s.checks) diffs.push(`checks: ${String(wb.get('Portfolio', 'BD2'))} vs ${s.checks}`);
        expect(diffs).toEqual([]);
      });
    }
  });
}
