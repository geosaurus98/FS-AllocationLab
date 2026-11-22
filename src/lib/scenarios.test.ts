import { describe, expect, it } from 'vitest';
import { emptyBounds, weightsFrom } from './reference';
import {
  createScenario,
  duplicateScenario,
  hasUnsavedChanges,
  resetDraftToBaseline,
  revertScenario,
  saveScenario,
  setDraftBound,
  setDraftWeight,
} from './scenarios';
import type { BaselineAllocation } from './types';

function deepFreeze<T>(o: T): T {
  if (o && typeof o === 'object') {
    Object.values(o).forEach(deepFreeze);
    Object.freeze(o);
  }
  return o;
}

const baseline: BaselineAllocation = deepFreeze({
  id: 'balanced6040:strategic',
  profileId: 'balanced6040',
  kind: 'strategic',
  weights: weightsFrom({ cash: 2, fixedInterest: 38, ausEquities: 17, globalEquities: 39, property: 4, multiSector: 0 }),
  status: 'illustrative',
  source: 'test',
  version: 3,
  effectiveDate: null,
  updatedAt: '2026-01-01',
  notes: '',
});
const bounds = deepFreeze((() => { const b = emptyBounds(); b.cash = { min: 2, max: null }; return b; })());

const make = () => createScenario({ id: 's1', name: 'Test', profileId: 'balanced6040', now: 't0', baseline, bounds });

describe('scenario isolation', () => {
  it('copies baseline weights and records provenance', () => {
    const s = make();
    expect(s.draft.weights).toEqual(baseline.weights);
    expect(s.draft.weights).not.toBe(baseline.weights);
    expect(s.origin).toMatchObject({ baselineId: 'balanced6040:strategic', baselineVersion: 3 });
  });

  it('editing a scenario never changes the (frozen) baseline', () => {
    let s = make();
    s = setDraftWeight(s, 'cash', 10);
    s = setDraftBound(s, 'cash', { min: 5, max: 20 });
    s = saveScenario(s, 't1');
    expect(baseline.weights.cash).toBe(2);
    expect(bounds.cash.min).toBe(2);
    expect(s.saved.weights.cash).toBe(10);
  });

  it('duplicates are independent of their source', () => {
    const a = setDraftWeight(make(), 'globalEquities', 45);
    const b = duplicateScenario(a, 's2', 't1');
    const b2 = setDraftWeight(b, 'globalEquities', 30);
    expect(a.draft.weights.globalEquities).toBe(45);
    expect(b2.draft.weights.globalEquities).toBe(30);
    expect(b.origin.duplicatedFromScenarioId).toBe('s1');
    expect(b.saved.weights).not.toBe(a.draft.weights);
  });

  it('save / revert / reset behave as drafts', () => {
    let s = setDraftWeight(make(), 'cash', 5);
    expect(hasUnsavedChanges(s)).toBe(true);
    s = revertScenario(s);
    expect(s.draft.weights.cash).toBe(2);
    expect(hasUnsavedChanges(s)).toBe(false);

    s = saveScenario(setDraftWeight(s, 'cash', 7), 't2');
    expect(hasUnsavedChanges(s)).toBe(false);
    s = resetDraftToBaseline(s, baseline);
    expect(s.draft.weights.cash).toBe(2);
    expect(s.saved.weights.cash).toBe(7); // reset is a draft edit until saved
    expect(hasUnsavedChanges(s)).toBe(true);
  });

  it('saved and draft bodies do not alias each other', () => {
    const s = saveScenario(setDraftWeight(make(), 'cash', 3), 't1');
    expect(s.saved.weights).not.toBe(s.draft.weights);
    expect(s.saved.bounds.cash).not.toBe(s.draft.bounds.cash);
  });
});
