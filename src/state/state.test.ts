import { describe, expect, it } from 'vitest';
import { createScenario } from '../lib/scenarios';
import { loadFixtureState } from './fixtures';
import { findBaseline, reducer } from './state';

describe('fixtures', () => {
  it('ships no approved data and labels the 60/40 allocation as illustrative', () => {
    const s = loadFixtureState();
    expect(s.baselines.some((b) => b.status === 'approved')).toBe(false);
    const b = findBaseline(s, 'balanced6040', 'strategic');
    expect(b?.status).toBe('illustrative');
    expect(b?.weights).toEqual({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 39, property: 4, multiSector: 0 });
    expect(s.constraints.find((c) => c.profileId === 'balanced6040')?.bounds.cash.min).toBe(2);
    expect(s.multiSector.growthFraction).toBeNull();
  });
  it('makes missing target allocations explicit', () => {
    const s = loadFixtureState();
    const t = findBaseline(s, 'defensive', 'target');
    expect(t?.status).toBe('missing');
    expect(Object.values(t!.weights).every((v) => v === null)).toBe(true);
  });
});

describe('reducer: scenario isolation', () => {
  it('scenario edits and saves never change baselines', () => {
    let s = loadFixtureState();
    const before = JSON.stringify(s.baselines);
    const base = findBaseline(s, 'balanced6040', 'strategic')!;
    const sc = createScenario({ id: 'x', name: 'X', profileId: 'balanced6040', now: 't', baseline: base, bounds: s.constraints[4]!.bounds });
    s = reducer(s, { type: 'scenario/add', scenario: sc });
    s = reducer(s, { type: 'scenario/setWeight', id: 'x', assetClass: 'cash', value: 12 });
    s = reducer(s, { type: 'scenario/setBound', id: 'x', assetClass: 'cash', bound: { min: 0, max: 50 } });
    s = reducer(s, { type: 'scenario/save', id: 'x', now: 't2' });
    s = reducer(s, { type: 'scenario/duplicate', id: 'x', newId: 'y', now: 't3' });
    s = reducer(s, { type: 'scenario/setWeight', id: 'y', assetClass: 'cash', value: 1 });
    expect(JSON.stringify(s.baselines)).toBe(before);
    expect(s.scenarios.find((x) => x.id === 'x')?.saved.weights.cash).toBe(12);
    expect(s.scenarios.find((x) => x.id === 'y')?.draft.weights.cash).toBe(1);
  });
});

describe('reducer: baseline versioning', () => {
  it('revisions create a new version and keep history', () => {
    let s = loadFixtureState();
    const cur = findBaseline(s, 'growth', 'strategic')!;
    s = reducer(s, { type: 'baseline/revise', now: 't', revision: { profileId: 'growth', kind: 'strategic', weights: { ...cur.weights, cash: 3, fixedInterest: 22 }, status: 'draft', source: 'edit', notes: '', effectiveDate: null } });
    expect(findBaseline(s, 'growth', 'strategic')?.version).toBe(2);
    expect(s.baselineHistory).toHaveLength(1);
    expect(s.baselineHistory[0]?.weights.cash).toBe(2);
    expect(s.auditLog).toHaveLength(1);
  });

  it('approved baselines are locked against revision', () => {
    let s = loadFixtureState();
    s = reducer(s, { type: 'baseline/approve', profileId: 'growth', kind: 'strategic', note: 'test', now: 't' });
    const approved = findBaseline(s, 'growth', 'strategic')!;
    expect(approved.status).toBe('approved');
    const after = reducer(s, { type: 'baseline/revise', now: 't2', revision: { profileId: 'growth', kind: 'strategic', weights: { ...approved.weights, cash: 50 }, status: 'draft', source: 'x', notes: '', effectiveDate: null } });
    expect(after).toBe(s);
  });

  it('missing baselines cannot be approved', () => {
    const s = loadFixtureState();
    expect(reducer(s, { type: 'baseline/approve', profileId: 'defensive', kind: 'target', note: '', now: 't' })).toBe(s);
  });
});
