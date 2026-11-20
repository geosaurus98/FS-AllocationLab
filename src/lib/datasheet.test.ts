// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildSecurityMaster, lookupSecurity, stripDetailSuffix } from './datasheet';
import { aggregateClient, analyseClient, classifyHoldings, parseClientsCsv, parseHoldingsCsv } from './holdings';
import { readXlsxSheet } from './xlsx';

const fixture = readFileSync(resolve(__dirname, '__fixtures__/datasheet-synthetic.xlsx'));
const buf = fixture.buffer.slice(fixture.byteOffset, fixture.byteOffset + fixture.byteLength);

describe('readXlsxSheet', () => {
  it('reads the preferred sheet with shared strings and numbers', async () => {
    const s = await readXlsxSheet(buf, 'Data');
    expect(s.sheetName).toBe('Data');
    expect(s.sheetNames).toEqual(['Notes', 'Data']);
    expect(s.rows[0]?.[2]).toBe('Apex Name');
    expect(s.rows[1]?.[3]).toBe('Demo Cash Fund');
  });
  it('rejects non-xlsx input', async () => {
    await expect(readXlsxSheet(new TextEncoder().encode('hello').buffer)).rejects.toThrow(/Not a valid .xlsx/);
  });
});

describe('security master', () => {
  const load = async () => buildSecurityMaster((await readXlsxSheet(buf, 'Data')).rows);

  it('strips only semicolon detail suffixes', () => {
    expect(stripDetailSuffix('Demo Bank NZ Ltd (188,000.00; 3.53%; 18 Months; 24/12/2026)')).toBe('Demo Bank NZ Ltd');
    expect(stripDetailSuffix('Demo Fund (PIE)')).toBe('Demo Fund (PIE)');
  });

  it('matches Apex, FS and OA names, case/whitespace-insensitively', async () => {
    const m = await load();
    const a = lookupSecurity(m, 'demo global share inc - usd  (DGS)');
    expect(a.status === 'matched' && a.record.assetClass).toBe('Global Equities');
    const f = lookupSecurity(m, 'Demo Cash Fund');
    expect(f.status === 'matched' && f.via).toBe('FS');
    const o = lookupSecurity(m, 'Demo Ampersand & Co');
    expect(o.status).toBe('matched');
  });

  it('matches term deposits after removing the detail suffix', async () => {
    const r = lookupSecurity(await load(), 'Demo Bank NZ Ltd (188,000.00; 3.53%; 18 Months; 24/12/2026)');
    expect(r.status === 'matched' && r.stripped).toBe(true);
    expect(r.status === 'matched' && r.record.type).toBe('Term Deposit');
  });

  it('reports conflicting classifications as ambiguous and unknown names as not found', async () => {
    const m = await load();
    expect(lookupSecurity(m, 'Demo Duplicate Fund').status).toBe('ambiguous');
    expect(lookupSecurity(m, 'Some Other Fund').status).toBe('not-found');
  });

  it('classifies holdings with precedence override > file > datasheet', async () => {
    const m = await load();
    const rows = parseHoldingsCsv('Portfolio Id\tAsset name\tAsset value NZD\nP1\tDemo Cash Fund (PIE)\t1000\nP1\tDemo Warrant - NZD (DWW)\t50\nP1\tMystery Fund\t200\nP1\tDemo Duplicate Fund\t300\n').rows;
    expect(rows).toHaveLength(4);
    const c = classifyHoldings(rows, m, { 'mystery fund': 'Property' });
    expect(c.map((x) => x.source)).toEqual(['datasheet', 'datasheet-no-class', 'override', 'ambiguous']);
    expect(c.map((x) => x.assetClass)).toEqual(['Cash', '', 'Property', '']);
    const agg = aggregateClient('P1', c, {});
    expect(agg.totalValue).toBe(1550);
    expect(agg.unmappedValue).toBe(350);
  });

  it('converts an annual fee % into monthly dollars from portfolio value', () => {
    const clients = parseClientsCsv('Client number,Risk Profile,Growth Target %,Annual Fee %,Monthly Withdrawals\nP1,Growth,60,1%,1000\n');
    expect(clients.errors).toEqual([]);
    const holdings = parseHoldingsCsv('Client number,Asset Name,Asset Value,Asset Class\nP1,Cash fund,"120,000",Cash\nP1,Shares,"1,080,000",Global Equities\n').rows;
    const a = analyseClient({ client: clients.rows[0]!, aggregate: aggregateClient('P1', holdings, {}), reference: null, referenceLabel: 'x', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(a.cash.monthlyFees).toBe(1000);
    expect(a.cash.monthlyOutflow).toBe(2000);
    expect(a.cash.monthsCovered).toBe(60);
  });
});
