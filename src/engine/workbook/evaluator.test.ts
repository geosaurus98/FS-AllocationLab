import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { Workbook, type WorkbookModel } from './evaluator';
import { parseFormula } from './parser';
import { BLANK, isErr } from './values';

const model: WorkbookModel = JSON.parse(readFileSync(resolve(__dirname, 'model.json'), 'utf8'));
const cached: Record<string, Record<string, [string, unknown]>> = JSON.parse(readFileSync(resolve(__dirname, 'cached-values.json'), 'utf8'));

describe('parser', () => {
  it('parses every formula in the workbook', () => {
    const failures: string[] = [];
    for (const [sheet, s] of Object.entries(model.sheets)) {
      for (const [r, c, k, p] of s.cells) {
        if (k !== 'f' && k !== 'a') continue;
        try {
          parseFormula(String(p));
        } catch (e) {
          failures.push(`${sheet}!R${r}C${c}: ${(e as Error).message.slice(0, 160)}`);
        }
      }
    }
    expect(failures.slice(0, 20)).toEqual([]);
  });
});

describe('evaluator reproduces Excel’s saved results on the template', () => {
  for (const sheet of ['Portfolio', 'Asset Allocation']) {
    it(sheet, () => {
      const wb = new Workbook(model, new Date(2026, 10, 20));
      const diffs: string[] = [];
      let checked = 0;
      for (const [rc, [kind, expected]] of Object.entries(cached[sheet]!)) {
        const [r, c] = rc.split(',').map(Number) as [number, number];
        if (c > 60) continue; // hyperlink/lookup helper columns far right are not part of the engine
        const got = wb.get(sheet, r, c);
        checked++;
        const ok =
          kind === 'n'
            ? typeof got === 'number' && Math.abs(got - (expected as number)) <= 1e-9 * Math.max(1, Math.abs(expected as number))
            : kind === 'b'
              ? got === expected
              : kind === 'e'
                ? isErr(got) && got.code === expected
                : (got === BLANK ? '' : typeof got === 'string' ? got : String(got)) === expected;
        if (!ok) diffs.push(`R${r}C${c} expected ${kind}:${JSON.stringify(expected)} got ${isErr(got) ? got.code : JSON.stringify(got === BLANK ? '<blank>' : got)} :: ${wb.formulaAt(sheet, r, c)?.slice(0, 100)}`);
      }
      expect(checked).toBeGreaterThan(400);
      expect(diffs.slice(0, 25)).toEqual([]);
    });
  }
});
