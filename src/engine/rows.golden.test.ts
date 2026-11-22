import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { cashRecommendations, msRecommendations } from './cashMs';
import { buildDataLookup, dataRowsFromGrid } from './datasheet';
import { isErr, type XValue } from './excel';
import { propertyAssetsToInsert, propertyModel, propertyRecommendation, type HoldingRow } from './property';
import { WORKBOOK_SETTINGS } from './settings';

/** Property, multi-sector and cash rows vs the workbook (scripts/golden/gen_property_ms.py). */

type V = number | string | null | { error: string };
interface Row { A: string; B: number | null; F: V; G: number | null; K: V }
interface Fixture {
  scenario: { id: string };
  workbook: {
    sectors: Record<string, { B: V; E: V; F: V; G: V }>;
    rows: Record<'cash' | 'prop' | 'ms', Row[]>;
    property: { C117: V; amounts: Record<string, V> };
  };
}

const fixtures: Fixture[] = JSON.parse(readFileSync(resolve(__dirname, '__golden__/rows-property-ms.json'), 'utf8'));
const data = buildDataLookup(dataRowsFromGrid(JSON.parse(readFileSync(resolve(__dirname, '__golden__/data-sheet.json'), 'utf8'))));
const num = (v: V): number => (typeof v === 'number' ? v : 0);

function same(actual: XValue | null, expected: V) {
  if (expected !== null && typeof expected === 'object') {
    expect(isErr(actual) ? actual.code : actual).toBe(expected.error);
  } else if (expected === null || expected === '') {
    expect(actual).toBeNull();
  } else {
    expect(isErr(actual)).toBe(false);
    expect(actual as number).toBeCloseTo(expected as number, 6);
  }
}

const toRow = (r: Row): HoldingRow => ({ name: r.A, value: r.B, override: r.G });

describe('property rows match the workbook', () => {
  it('has fixtures', () => expect(fixtures.length).toBeGreaterThanOrEqual(30));
  for (const f of fixtures) {
    it(`${f.scenario.id}`, () => {
      const w = f.workbook;
      const target = num(w.sectors.prop!.G);
      const rows = w.rows.prop.map(toRow);
      const model = propertyModel(WORKBOOK_SETTINGS, target, rows);
      same(model.pot, w.property.C117);
      for (const line of model.lines) if (line.name) same(line.amount, w.property.amounts[line.name] ?? null);
      // every model asset the macro would insert is present in the workbook rows
      for (const name of propertyAssetsToInsert(model)) expect(w.rows.prop.some((r) => r.A === name)).toBe(true);
      for (const r of w.rows.prop) same(propertyRecommendation(WORKBOOK_SETTINGS, model, toRow(r), data), r.F);
    });
  }
});

describe('multi-sector and cash rows match the workbook', () => {
  for (const f of fixtures) {
    it(`${f.scenario.id}`, () => {
      const w = f.workbook;
      const ms = w.rows.ms.map(toRow);
      if (ms.length) {
        const bColumn: (number | string | null)[] = [null, ...Array.from({ length: 7 }, (_, i) => ms[i]?.value ?? null), num(w.sectors.ms!.B), null, null, null, null, null, null];
        const got = msRecommendations(ms, bColumn, [num(w.sectors.ms!.G), 0, 0, 0, 0, 0, 0]);
        w.rows.ms.forEach((r, i) => same(got[i]!, r.F));
      }
      const s = w.sectors;
      const changes = { fi: num(s.fi!.E), ae: num(s.ae!.E), ge: num(s.ge!.E), prop: num(s.prop!.E), ms: num(s.ms!.E) };
      const cash = cashRecommendations(w.rows.cash.map(toRow), changes, (f.scenario as { addRemove?: number }).addRemove ?? null);
      w.rows.cash.forEach((r, i) => same(cash[i]!, r.F));
    });
  }
});
