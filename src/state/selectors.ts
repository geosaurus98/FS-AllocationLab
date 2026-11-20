import { growthDefensiveSplit, sumWeights, type GrowthDefensiveSplit, type TotalResult } from '../lib/allocation';
import { profile } from '../lib/reference';
import { effectiveFraction } from '../lib/scenarios';
import type { BaselineAllocation, Bounds, ProfileId, Scenario, ValidationResult, Weights } from '../lib/types';
import { validateAllocation } from '../lib/validation';
import type { AppState } from './state';
import { findConstraints } from './state';

export interface Evaluation {
  total: TotalResult;
  split: GrowthDefensiveSplit;
  validation: ValidationResult;
  fraction: number | null;
}

export function evaluate(state: AppState, profileId: ProfileId, weights: Weights, bounds: Bounds | null, fraction: number | null): Evaluation {
  return {
    total: sumWeights(weights),
    split: growthDefensiveSplit(weights, fraction),
    validation: validateAllocation({ weights, bounds, multiSectorGrowthFraction: fraction, settings: state.settings, profile: profile(profileId) }),
    fraction,
  };
}

export function evaluateBaseline(state: AppState, b: BaselineAllocation): Evaluation {
  const c = findConstraints(state, b.profileId);
  return evaluate(state, b.profileId, b.weights, c?.bounds ?? null, state.multiSector.growthFraction);
}

export function evaluateScenario(state: AppState, s: Scenario, which: 'draft' | 'saved' = 'draft'): Evaluation {
  const body = s[which];
  return evaluate(state, s.profileId, body.weights, body.bounds, effectiveFraction(body, state.multiSector.growthFraction));
}
