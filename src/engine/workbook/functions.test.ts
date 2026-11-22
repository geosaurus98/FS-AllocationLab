import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { WORKBOOK_SETTINGS } from '../settings';
import { Workbook, type WorkbookModel } from './evaluator';
import { applyFirmSettings } from './settingsCells';
import { BLANK } from './values';

/** Focused checks of Excel semantics the workbook relies on, using a tiny in-memory model. */
function wb(cells: Record<string, string | number>, arrays: string[] = []) {
  const model: WorkbookModel = {
    source: 'test.xlsx',
    names: {},
    iterate: true,
    iterateCount: 100,
    iterateDelta: 0.001,
    sheets: {
      S: {
        maxRow: 50,
        maxCol: 10,
        cells: Object.entries(cells).map(([a1, v]) => {
          const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
          const c = m[1]!.charCodeAt(0) - 64;
          const r = Number(m[2]);
          if (typeof v === 'string' && v.startsWith('=')) return [r, c, arrays.includes(a1) ? 'a' : 'f', v.slice(1)];
          return [r, c, typeof v === 'number' ? 'n' : 's', v];
        }),
      },
    },
  };
  return new Workbook(model, new Date(2026, 10, 20));
}

describe('Excel semantics', () => {
  it('COUNTIF "<>0" counts empty cells too (the multi-sector quirk)', () => {
    const w = wb({ A1: 5, A3: 0, B1: '=COUNTIF(A1:A8,"<>0")' });
    expect(w.get('S', 'B1')).toBe(7);
  });
  it('SUMIF with an array-constant criterion returns an array that SUM adds up', () => {
    const w = wb({ A1: 'ETF', A2: 'Gold', A3: 'Bond', B1: 10, B2: 20, B3: 40, C1: '=SUM(SUMIF(A1:A3,{"ETF","Gold"},B1:B3))' });
    expect(w.get('S', 'C1')).toBe(30);
  });
  it('implicit intersection in normal formulas, element-wise in array formulas', () => {
    const w = wb({ A1: 1, A2: 2, A3: 3, B2: '=A1:A3*10', C1: '=SUM(A1:A3*10)' }, ['C1']);
    expect(w.get('S', 'B2')).toBe(20);
    expect(w.get('S', 'C1')).toBe(60);
  });
  it('FILTER and TEXTJOIN (Excel 365) work, and fail as in LibreOffice in compat mode', () => {
    const cells = { A1: 'a', A2: 'b', A3: 'c', B1: 1, B2: 0, B3: 1, C1: '=TEXTJOIN(", ",TRUE,FILTER(A1:A3,B1:B3=1,""))' };
    expect(wb(cells).get('S', 'C1')).toBe('a, c');
    const w = wb(cells);
    w.libreOfficeCompat = true;
    expect(w.get('S', 'C1')).toMatchObject({ code: '#NAME?' });
  });
  it('DATEVALUE reads NZ day/month/year and MOD/ROUND follow Excel', () => {
    const w = wb({ A1: '=YEAR(DATEVALUE("24/12/2026"))', A2: '=MOD(-7,100)', A3: '=ROUND(1250,-2)', A4: '=ROUND(-2.5,0)' });
    expect(w.get('S', 'A1')).toBe(2026);
    expect(w.get('S', 'A2')).toBe(93);
    expect(w.get('S', 'A3')).toBe(1300);
    expect(w.get('S', 'A4')).toBe(-3);
  });
  it('a formula pointing at an empty cell shows 0, and blank compares equal to ""', () => {
    const w = wb({ A1: '=B9', A2: '=B9=""', A3: '=B9=0' });
    expect(w.get('S', 'A1')).toBe(0);
    expect(w.get('S', 'A2')).toBe(true);
    expect(w.get('S', 'A3')).toBe(true);
    expect(w.get('S', 'B9')).toBe(BLANK);
  });
  it('solves a self-referencing cell by iteration (the Custom property cell)', () => {
    const w = wb({ A1: 0.04, A2: 0.6, A3: 0.7, B1: '=(A1/(A2-A1))*(A3-B1)' });
    expect(w.get('S', 'B1') as number).toBeCloseTo((0.04 / 0.6) * 0.7, 12);
  });
});

describe('firm settings flow into the workbook', () => {
  const model: WorkbookModel = JSON.parse(readFileSync(resolve(__dirname, 'model.json'), 'utf8'));
  it('changing Growth cash/FI changes the Growth profile column; derived cells stay formulas', () => {
    const w = new Workbook(model);
    const s = JSON.parse(JSON.stringify(WORKBOOK_SETTINGS));
    s.tactical.Growth.cash = 0.03;
    s.tactical.Growth.fi = 0.17;
    applyFirmSettings(w, s);
    expect(w.get('Asset Allocation', 'F4')).toBe(0.03);
    expect(w.get('Asset Allocation', 'F5')).toBe(0.17);
    expect(w.formulaAt('Asset Allocation', 7, 6)).not.toBeNull(); // F7 still a formula
  });
  it('workbook defaults are a no-op', () => {
    const a = new Workbook(model);
    const b = new Workbook(model);
    applyFirmSettings(b, WORKBOOK_SETTINGS);
    for (const cell of ['C7', 'D8', 'H9', 'D119', 'D130']) expect(b.get('Asset Allocation', cell)).toEqual(a.get('Asset Allocation', cell));
  });
});
