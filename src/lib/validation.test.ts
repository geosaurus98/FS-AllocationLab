import { describe, expect, it } from 'vitest';
import { emptyBounds, profile, weightsFrom } from './reference';
import type { Settings } from './types';
import { validateAllocation, validateBounds } from './validation';

const settings: Settings = { sumTolerance: 0.01, nominalSplitTolerance: 0.5 };
const w6040 = weightsFrom({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 39, property: 4, multiSector: 0 });
const minCash2 = () => {
  const b = emptyBounds();
  b.cash = { min: 2, max: null };
  return b;
};

describe('validateAllocation', () => {
  it('accepts the illustrative Balanced 60/40 allocation with a 2% minimum Cash', () => {
    const r = validateAllocation({ weights: w6040, bounds: minCash2(), multiSectorGrowthFraction: null, settings, profile: profile('balanced6040') });
    expect(r.status).toBe('valid');
    expect(r.issues).toEqual([]);
  });

  it('rejects totals outside tolerance and explains the gap', () => {
    const w = { ...w6040, globalEquities: 40 };
    const r = validateAllocation({ weights: w, bounds: minCash2(), multiSectorGrowthFraction: null, settings });
    expect(r.status).toBe('invalid');
    const issue = r.issues.find((i) => i.code === 'SUM_NOT_100');
    expect(issue?.message).toContain('101.00%');
    expect(issue?.message).toContain('over');
  });

  it('accepts totals inside the decimal tolerance', () => {
    const w = { ...w6040, globalEquities: 39.005 };
    const r = validateAllocation({ weights: w, bounds: minCash2(), multiSectorGrowthFraction: null, settings });
    expect(r.issues.some((i) => i.code === 'SUM_NOT_100')).toBe(false);
  });

  it('enforces the configurable minimum Cash weight', () => {
    const w = { ...w6040, cash: 1.5, fixedInterest: 38.5 };
    const r = validateAllocation({ weights: w, bounds: minCash2(), multiSectorGrowthFraction: null, settings });
    expect(r.status).toBe('invalid');
    expect(r.issues.map((i) => i.code)).toContain('BELOW_MIN_CASH');

    const b = emptyBounds();
    b.cash = { min: 1, max: null };
    const r2 = validateAllocation({ weights: w, bounds: b, multiSectorGrowthFraction: null, settings });
    expect(r2.issues.map((i) => i.code)).not.toContain('BELOW_MIN_CASH');
  });

  it('checks min/max for other asset classes', () => {
    const b = minCash2();
    b.globalEquities = { min: 10, max: 35 };
    b.fixedInterest = { min: 40, max: null };
    const r = validateAllocation({ weights: w6040, bounds: b, multiSectorGrowthFraction: null, settings });
    const codes = r.issues.map((i) => `${i.code}:${i.assetClass}`);
    expect(codes).toContain('ABOVE_MAX:globalEquities');
    expect(codes).toContain('BELOW_MIN:fixedInterest');
  });

  it('warns when no minimum Cash weight is configured', () => {
    const r = validateAllocation({ weights: w6040, bounds: emptyBounds(), multiSectorGrowthFraction: null, settings });
    expect(r.status).toBe('warnings');
    expect(r.issues.map((i) => i.code)).toContain('MIN_CASH_NOT_SET');
  });

  it('marks allocations with missing weights as incomplete', () => {
    const r = validateAllocation({ weights: { ...w6040, property: null }, bounds: minCash2(), multiSectorGrowthFraction: null, settings });
    expect(r.status).toBe('incomplete');
    expect(r.issues.find((i) => i.code === 'MISSING_WEIGHT')?.assetClass).toBe('property');
  });

  it('reports a fully missing allocation as missing', () => {
    const r = validateAllocation({ weights: weightsFrom({}), bounds: null, multiSectorGrowthFraction: null, settings });
    expect(r.status).toBe('missing');
  });

  it('flags unclassified Multi-Sector exposure', () => {
    const w = weightsFrom({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 34, property: 4, multiSector: 5 });
    const r = validateAllocation({ weights: w, bounds: minCash2(), multiSectorGrowthFraction: null, settings });
    expect(r.issues.map((i) => i.code)).toContain('MULTI_SECTOR_UNCLASSIFIED');
    const r2 = validateAllocation({ weights: w, bounds: minCash2(), multiSectorGrowthFraction: 0.5, settings });
    expect(r2.issues.map((i) => i.code)).not.toContain('MULTI_SECTOR_UNCLASSIFIED');
  });

  it('distinguishes Balanced 50/50 from 60/40 via the nominal split check', () => {
    const r = validateAllocation({ weights: w6040, bounds: minCash2(), multiSectorGrowthFraction: null, settings, profile: profile('balanced5050') });
    expect(r.issues.map((i) => i.code)).toContain('NOMINAL_SPLIT_MISMATCH');
  });

  it('rejects negative weights and never mutates input', () => {
    const w = Object.freeze({ ...w6040, cash: -2, fixedInterest: 42 });
    const r = validateAllocation({ weights: w, bounds: null, multiSectorGrowthFraction: null, settings });
    expect(r.issues.map((i) => i.code)).toContain('NEGATIVE_WEIGHT');
    expect(w.cash).toBe(-2);
  });
});

describe('validateBounds', () => {
  it('rejects min > max and infeasible minimum totals', () => {
    const b = emptyBounds();
    b.cash = { min: 10, max: 5 };
    b.fixedInterest = { min: 95, max: null };
    const codes = validateBounds(b).map((i) => i.message);
    expect(codes.some((m) => m.includes('greater than maximum'))).toBe(true);
    expect(codes.some((m) => m.includes('no allocation can satisfy'))).toBe(true);
  });
});
