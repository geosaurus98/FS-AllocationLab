import { describe, expect, it } from 'vitest';
import { cleanFloat, diffWeights, growthDefensiveSplit, sumsTo100, sumWeights } from './allocation';
import { weightsFrom } from './reference';

const illustrative6040 = weightsFrom({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 39, property: 4, multiSector: 0 });

describe('sumWeights / sumsTo100', () => {
  it('sums the illustrative Balanced 60/40 allocation to exactly 100', () => {
    const r = sumWeights(illustrative6040);
    expect(r.total).toBe(100);
    expect(r.complete).toBe(true);
    expect(sumsTo100(r.total, 0.01)).toBe(true);
  });

  it('removes floating-point noise from decimal weights', () => {
    const w = weightsFrom({ cash: 0.1, fixedInterest: 0.2, ausEquities: 33.3, globalEquities: 33.3, property: 33.1, multiSector: 0 });
    expect(sumWeights(w).total).toBe(100);
    expect(cleanFloat(0.1 + 0.2)).toBe(0.3);
  });

  it('applies the tolerance in percentage points', () => {
    expect(sumsTo100(100.005, 0.01)).toBe(true);
    expect(sumsTo100(99.99, 0.01)).toBe(true);
    expect(sumsTo100(99.98, 0.01)).toBe(false);
    expect(sumsTo100(100.5, 0.01)).toBe(false);
    expect(sumsTo100(100.5, 0.5)).toBe(true);
  });

  it('reports missing weights instead of treating them as zero', () => {
    const w = weightsFrom({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 39, property: null, multiSector: 0 });
    const r = sumWeights(w);
    expect(r.total).toBe(96);
    expect(r.complete).toBe(false);
    expect(r.missing).toEqual(['property']);
  });
});

describe('growthDefensiveSplit', () => {
  it('classifies Cash/Fixed Interest as defensive and equities/property as growth', () => {
    const s = growthDefensiveSplit(illustrative6040, null);
    expect(s.growth).toBe(60);
    expect(s.defensive).toBe(40);
    expect(s.unclassified).toBe(0);
  });

  it('flags Multi-Sector as unclassified when no look-through fraction is set', () => {
    const w = weightsFrom({ cash: 5, fixedInterest: 35, ausEquities: 15, globalEquities: 35, property: 0, multiSector: 10 });
    const s = growthDefensiveSplit(w, null);
    expect(s.unclassified).toBe(10);
    expect(s.growth).toBe(50);
    expect(s.defensive).toBe(40);
    expect(s.growthFractionApplied).toBeNull();
  });

  it('applies a configured Multi-Sector look-through fraction', () => {
    const w = weightsFrom({ cash: 5, fixedInterest: 35, ausEquities: 15, globalEquities: 35, property: 0, multiSector: 10 });
    const s = growthDefensiveSplit(w, 0.6);
    expect(s.unclassified).toBe(0);
    expect(s.multiSectorGrowth).toBe(6);
    expect(s.multiSectorDefensive).toBe(4);
    expect(s.growth).toBe(56);
    expect(s.defensive).toBe(44);
  });

  it('treats an out-of-range fraction as unclassified rather than clamping it', () => {
    const w = weightsFrom({ cash: 5, fixedInterest: 35, ausEquities: 15, globalEquities: 35, property: 0, multiSector: 10 });
    expect(growthDefensiveSplit(w, 1.5).unclassified).toBe(10);
    expect(growthDefensiveSplit(w, -0.1).unclassified).toBe(10);
  });

  it('does not flag zero Multi-Sector weight even without a fraction', () => {
    expect(growthDefensiveSplit(illustrative6040, null).unclassified).toBe(0);
  });

  it('distinguishes Balanced 50/50 from Balanced 60/40', () => {
    const b5050 = weightsFrom({ cash: 3, fixedInterest: 47, ausEquities: 15, globalEquities: 31, property: 4, multiSector: 0 });
    expect(growthDefensiveSplit(b5050, null).growth).toBe(50);
    expect(growthDefensiveSplit(illustrative6040, null).growth).toBe(60);
  });
});

describe('diffWeights', () => {
  it('computes percentage-point differences and null against missing data', () => {
    const proposed = weightsFrom({ cash: 3, fixedInterest: 37, ausEquities: 17, globalEquities: 39, property: null, multiSector: 0 });
    const d = diffWeights(illustrative6040, proposed);
    expect(d.cash).toBe(1);
    expect(d.fixedInterest).toBe(-1);
    expect(d.ausEquities).toBe(0);
    expect(d.property).toBeNull();
  });
});
