import { describe, expect, it } from 'vitest';
import {
  aggregateClient,
  analyseClient,
  classLabels,
  parseClientsCsv,
  parseHoldingsCsv,
  parseMoney,
  resolveClass,
  type ClientRecord,
} from './holdings';
import { weightsFrom } from './reference';

const ref6040 = weightsFrom({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 39, property: 4, multiSector: 0 });
const H = 'Client number,Asset Name,Asset Value,Asset Class,Asset Sub-Class';

const client = (over: Partial<ClientRecord> = {}): ClientRecord => ({
  line: 2, clientNumber: 'C1', profileId: 'balanced6040', growthTarget: 60, monthlyFees: 500, annualFeePct: null, monthlyWithdrawals: 1500,
  cashCoverMonths: null, driftTolerance: null, notes: '', ...over,
});

describe('parseMoney', () => {
  it('handles $, commas, parentheses and blanks', () => {
    expect(parseMoney('$12,345.67')).toEqual({ ok: true, value: 12345.67 });
    expect(parseMoney('(1,000)')).toEqual({ ok: true, value: -1000 });
    expect(parseMoney('-$50')).toEqual({ ok: true, value: -50 });
    expect(parseMoney('')).toEqual({ ok: true, value: null });
    expect(parseMoney('ten').ok).toBe(false);
  });
});

describe('parseHoldingsCsv', () => {
  it('parses the five-column layout and reports bad rows by line', () => {
    const r = parseHoldingsCsv(`${H}\nC1,Fund A,"$1,000",Cash,NZ\nC1,Fund B,abc,Cash,NZ\n,Fund C,10,Cash,\nC2,Fund D,-5,Property,\n`);
    expect(r.fatal).toEqual([]);
    expect(r.rows.map((x) => x.value)).toEqual([1000, -5]);
    expect(r.errors.map((e) => e.line)).toEqual([3, 4]);
    expect(r.warnings[0]?.message).toMatch(/Negative/);
  });
  it('rejects files missing required columns', () => {
    expect(parseHoldingsCsv('Client number,Asset Value\nC1,5').fatal[0]).toMatch(/Asset Name/);
  });
});

describe('parseClientsCsv', () => {
  const CH = 'Client number,Risk Profile,Growth Target %,Monthly Fees,Monthly Withdrawals';
  it('parses profiles by label, csv key or id and rejects duplicates/unknowns', () => {
    const r = parseClientsCsv(`${CH}\nC1,Balanced 60/40,60,$400,"$2,000"\nC2,balanced_50_50,50,0,0\nC1,Growth,75,0,0\nC3,Aggressive,90,0,0\n`);
    expect(r.rows.map((c) => c.profileId)).toEqual(['balanced6040', 'balanced5050']);
    expect(r.rows[0]?.monthlyWithdrawals).toBe(2000);
    expect(r.errors.map((e) => e.message).join(' ')).toMatch(/duplicate.*C1/);
    expect(r.errors.map((e) => e.message).join(' ')).toMatch(/unknown Risk Profile "Aggressive"/);
  });
  it('rejects out-of-range growth targets and negative fees', () => {
    const r = parseClientsCsv(`${CH}\nC1,Growth,120,-1,0\n`);
    expect(r.rows).toHaveLength(0);
    expect(r.errors[0]?.message).toMatch(/between 0 and 100/);
    expect(r.errors[0]?.message).toMatch(/cannot be negative/);
  });
});

describe('class mapping', () => {
  it('maps exact synonyms by default and leaves others unmapped', () => {
    expect(resolveClass('Fixed Income', {})).toBe('fixedInterest');
    expect(resolveClass('  multi-sector ', {})).toBe('multiSector');
    expect(resolveClass('Diversified', {})).toBeNull();
    expect(resolveClass('Diversified', { diversified: 'multiSector' })).toBe('multiSector');
    expect(resolveClass('Cash', { cash: null })).toBeNull();
  });
  it('summarises labels with value', () => {
    const rows = parseHoldingsCsv(`${H}\nC1,a,10,Cash,\nC1,b,30,Gold,\nC2,c,5,cash,\n`).rows;
    const labels = classLabels(rows, {});
    expect(labels.find((l) => l.key === 'gold')).toMatchObject({ value: 30, mapped: null });
    expect(labels.find((l) => l.key === 'cash')).toMatchObject({ value: 15, count: 2, mapped: 'cash' });
  });
});

describe('aggregate & analyse', () => {
  const rows = parseHoldingsCsv(
    `${H}\nC1,a,"20,000",Cash,\nC1,b,"190,000",Fixed Interest,\nC1,c,"90,000",Australasian Equities,\nC1,d,"180,000",Global Equities,\nC1,e,"20,000",Property,\nC2,x,100,Gold,\n`,
  ).rows;

  it('turns values into weights over total value, keeping unmapped value explicit', () => {
    const agg = aggregateClient('C1', rows, {});
    expect(agg.totalValue).toBe(500000);
    expect(agg.weights).toEqual(weightsFrom({ cash: 4, fixedInterest: 38, ausEquities: 18, globalEquities: 36, property: 4, multiSector: 0 }));
    const agg2 = aggregateClient('C2', rows, {});
    expect(agg2.unmappedValue).toBe(100);
    expect(agg2.weights.cash).toBe(0);
  });

  it('measures drift against the profile and flags breaches beyond tolerance', () => {
    const a = analyseClient({ client: client(), aggregate: aggregateClient('C1', rows, {}), reference: ref6040, referenceLabel: 'Strategic', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(a.drift?.globalEquities).toBe(-3);
    expect(a.drift?.cash).toBe(2);
    expect(a.breaches).toEqual(['globalEquities']);
    expect(a.split.growth).toBe(58);
    expect(a.growthGap).toBe(-2);
  });

  it('uses per-client tolerance overrides', () => {
    const a = analyseClient({ client: client({ driftTolerance: 5 }), aggregate: aggregateClient('C1', rows, {}), reference: ref6040, referenceLabel: 'Strategic', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(a.breaches).toEqual([]);
  });

  it('assesses cash coverage of monthly fees and withdrawals', () => {
    const a = analyseClient({ client: client(), aggregate: aggregateClient('C1', rows, {}), reference: ref6040, referenceLabel: 'Strategic', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(a.cash.monthlyOutflow).toBe(2000);
    expect(a.cash.requiredCash).toBe(24000);
    expect(a.cash.monthsCovered).toBe(10);
    expect(a.cash.shortfall).toBe(4000);
    expect(a.cash.requiredCashPct).toBe(4.8);
    expect(a.cash.impliedCashPct).toBe(4.8);
    expect(a.issues.some((i) => /shortfall \$4,000\.00/.test(i.message))).toBe(true);
  });

  it('does not assess coverage when fees/withdrawals are missing', () => {
    const a = analyseClient({ client: client({ monthlyFees: null }), aggregate: aggregateClient('C1', rows, {}), reference: ref6040, referenceLabel: 'Strategic', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(a.cash.requiredCash).toBeNull();
    expect(a.issues.some((i) => /cannot be assessed/.test(i.message))).toBe(true);
  });

  it('flags missing client details, missing holdings and missing reference data', () => {
    const none = analyseClient({ client: null, aggregate: aggregateClient('ZZ', rows, {}), reference: null, referenceLabel: 'Strategic', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(none.issues.filter((i) => i.severity === 'error')).toHaveLength(2);
    const noHoldings = analyseClient({ client: client(), aggregate: aggregateClient('ZZ', rows, {}), reference: ref6040, referenceLabel: 'Strategic', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(noHoldings.cash.monthsCovered).toBeNull();
    expect(noHoldings.cash.shortfall).toBeNull();
    const noRef = analyseClient({ client: client(), aggregate: aggregateClient('C1', rows, {}), reference: null, referenceLabel: 'Target', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(noRef.drift).toBeNull();
    expect(noRef.issues.some((i) => /Target has no data/.test(i.message))).toBe(true);
  });

  it('notes when the client growth target differs from the profile model', () => {
    const a = analyseClient({ client: client({ growthTarget: 55 }), aggregate: aggregateClient('C1', rows, {}), reference: ref6040, referenceLabel: 'Strategic', multiSectorGrowthFraction: null, defaultCoverMonths: 12, defaultTolerance: 2 });
    expect(a.referenceGrowthGap).toBe(5);
  });

  it('never mutates inputs', () => {
    const frozen = Object.freeze(rows.map((r) => Object.freeze({ ...r })));
    expect(() => aggregateClient('C1', frozen as typeof rows, {})).not.toThrow();
  });
});
