import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { beforeAll, describe, expect, it } from 'vitest';
import { adviceTables, money, percent2, portfolioView, signedMoney, fixed2, type AdviceTables } from './documentTables';
import { Workbook, type WorkbookModel } from './evaluator';
import { applyScenario, SECTIONS, type Scenario } from './scenario';
import { BLANK } from './values';

const model: WorkbookModel = JSON.parse(readFileSync(resolve(__dirname, 'model.json'), 'utf8'));
const fixtures: { scenario: Scenario & { custodian?: 'Apex' | 'OA' } }[] = JSON.parse(readFileSync(resolve(__dirname, '__golden_rounds__/rounds.json'), 'utf8'));

function calc(i: number): Workbook {
  const wb = new Workbook(model, new Date(2026, 10, 20));
  wb.libreOfficeCompat = true;
  const sc = fixtures[i]!.scenario;
  applyScenario(wb, sc);
  wb.set('Portfolio', 'BD1', sc.custodian ?? 'Apex');
  return wb;
}

const num = (s: string) => Number(s.replace(/[$,+]/g, ''));

describe('number formats', () => {
  it('formats like Excel', () => {
    expect(money(1234.5)).toBe('$1,235');
    expect(money(-1234.5)).toBe('-$1,235');
    expect(money(0)).toBe('$0');
    expect(money(BLANK)).toBe('');
    expect(signedMoney(5000)).toBe('+$5,000');
    expect(signedMoney(-5000)).toBe('-$5,000');
    expect(signedMoney(0)).toBe('$0');
    expect(fixed2(12.345)).toBe('12.35');
    expect(percent2(0.0125)).toBe('1.25%');
  });
});

describe('advice tables from a calculated portfolio', () => {
  let wb: Workbook;
  let t: AdviceTables;
  beforeAll(() => {
    wb = calc(0);
    t = adviceTables(wb, { clientName: 'Sample Client', custodian: 'Apex' });
  });

  it('sorts holdings by name within each section except cash', () => {
    const view = portfolioView(wb);
    for (const s of SECTIONS.filter((x) => x.key !== 'cash')) {
      const names = view.filter((v) => v.section === s.key && v.kind === 'holding').map((v) => v.name);
      expect(names).toEqual([...names].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' })));
    }
  });

  it('portfolio for advice: changed holdings bracketed by their section header and total', () => {
    const rows = t.portfolio.rows;
    expect(rows[0]!.map((c) => c.text)).toEqual(['', 'Portfolio', 'Changes', 'Recomm. Portfolio', '% Allocation']);
    expect(rows.at(-1)![0]!.text).toBe('Portfolio Total');
    expect(rows.at(-1)![1]!.text).toBe(money(wb.get('Portfolio', 202, 2)));
    let open: string | null = null;
    for (const r of rows.slice(1, -1)) {
      const name = r[0]!.text;
      if (r[0]!.bold && !name.endsWith('Total')) open = name;
      else if (r[0]!.bold) {
        if (open) expect(name).toBe(`${open} Total`);
        open = null;
      } else {
        expect(open).not.toBeNull();
        expect(num(r[2]!.text)).not.toBe(0);
        expect(name.length).not.toBe(3);
      }
    }
  });

  it('costs: header from row 2, no cash currencies or term deposits', () => {
    const c = t.costs!;
    expect(c.rows[0]!.map((x) => x.text)).toEqual(['', 'Broker', 'Brokerage', 'Buy', 'Sell', 'Cost p.a.', 'Product Disclosure']);
    expect(c.widths).toHaveLength(7);
    for (const r of c.rows.slice(1)) {
      expect(r[0]!.text.length).not.toBe(3);
      expect(r).toHaveLength(7);
      if (r[6]!.text === 'Click here') expect(r[6]!.link).toMatch(/^https?:/);
    }
  });

  it('summary table adds up', () => {
    const rows = t.summary.rows;
    expect(rows[0]![0]!.text).toBe('Sample Client');
    const body = rows.slice(2, -1);
    const totals = rows.at(-1)!;
    expect(totals[0]!.text).toBe('Totals');
    // section totals add to the Totals row
    const sectionTotals = body.filter((r) => r[0]!.bold && r[0]!.text.endsWith('Total'));
    const sumC = sectionTotals.reduce((a, r) => a + num(r[2]!.text), 0);
    expect(Math.abs(sumC - num(totals[2]!.text))).toBeLessThanOrEqual(sectionTotals.length);
    // within a section: listed + "All Other … (No Changes)" = section total (recommended)
    let listed = 0;
    let started = false;
    for (const r of body) {
      const a = r[0]!.text;
      if (r[0]!.bold && !a.endsWith('Total')) {
        listed = 0;
        started = true;
      } else if (a.startsWith('All Other') && a.endsWith('(No Changes)')) listed += num(r[3]!.text);
      else if (r[0]!.bold && started) {
        expect(Math.abs(listed - num(r[3]!.text))).toBeLessThanOrEqual(20);
        started = false;
      } else if (started) {
        listed += num(r[3]!.text);
        const c = num(r[2]!.text);
        const b = r[1]!.text;
        if (c > 0) expect(['Buy', 'Add']).toContain(b);
        if (c < 0) expect(['Trim', 'Sell', 'Maturity']).toContain(b);
      }
    }
  });

  it('summary table lists buys then sells, each by name', () => {
    const rows = t.summary.rows.slice(2, -1);
    let block: { name: string; c: number }[] = [];
    const check = () => {
      const buys = block.filter((x) => x.c > 0).map((x) => x.name);
      const sells = block.filter((x) => x.c < 0).map((x) => x.name);
      expect(block.map((x) => x.name)).toEqual([...buys, ...sells]);
      expect(buys).toEqual([...buys].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' })));
      expect(sells).toEqual([...sells].sort((a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' })));
      block = [];
    };
    let inSection = false;
    for (const r of rows) {
      const a = r[0]!.text;
      if (r[0]!.bold) {
        if (inSection) check();
        inSection = !a.endsWith('Total');
      } else if (inSection && !a.startsWith('All Other')) block.push({ name: a, c: num(r[2]!.text) });
    }
  });

  it('investment rationale: a title, then headings each followed by bullets', () => {
    const r = t.rationale;
    expect(r[0]).toEqual({ kind: 'title', text: 'Investment Rationale' });
    for (let i = 1; i < r.length; i++) if (r[i]!.kind === 'heading') expect(r[i + 1]?.kind).toBe('bullet');
    const bullets = r.filter((x) => x.kind === 'bullet');
    expect(bullets.length).toBeGreaterThan(0);
    for (const b of bullets) expect(b.text).toMatch(/investment in /);
    expect(bullets.some((b) => / Total\.$/.test(b.text))).toBe(false);
  });

  it('every fixture produces consistent tables', () => {
    for (let i = 0; i < fixtures.length; i++) {
      const w = calc(i);
      const a = adviceTables(w, { clientName: 'X', custodian: fixtures[i]!.scenario.custodian ?? 'Apex' });
      expect(a.portfolio.rows.at(-1)![0]!.text).toBe('Portfolio Total');
      expect(a.summary.rows.at(-1)![3]!.text).toBe(money(w.get('Portfolio', 202, 6)));
    }
  });
});
