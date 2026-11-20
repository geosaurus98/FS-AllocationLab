import { cloneWeights, weightsEqual } from './allocation';
import { profile as profileDef } from './reference';
import type { AssetClassId, BaselineAllocation, Bound, Bounds, ProfileId, Scenario, ScenarioBody, Weights } from './types';
import { ASSET_CLASS_IDS } from './types';

/**
 * Pure scenario operations. Every function returns new objects and deep-copies
 * weights/bounds so scenarios can never alias (and therefore never mutate)
 * baseline allocations or each other.
 */

export function cloneBounds(b: Bounds): Bounds {
  return Object.fromEntries(ASSET_CLASS_IDS.map((id) => [id, { ...b[id] }])) as Bounds;
}

export function cloneBody(body: ScenarioBody): ScenarioBody {
  return {
    weights: cloneWeights(body.weights),
    bounds: cloneBounds(body.bounds),
    multiSectorGrowthFraction: body.multiSectorGrowthFraction,
  };
}

export interface CreateScenarioArgs {
  id: string;
  name: string;
  profileId: ProfileId;
  now: string;
  /** Baseline to copy weights from. Omit to start with blank (missing) weights. */
  baseline?: BaselineAllocation | null;
  /** Constraint bounds to copy. */
  bounds: Bounds;
  initialWeights?: Weights;
  description?: string;
}

export function createScenario(args: CreateScenarioArgs): Scenario {
  const weights = args.initialWeights
    ? cloneWeights(args.initialWeights)
    : args.baseline
      ? cloneWeights(args.baseline.weights)
      : (Object.fromEntries(ASSET_CLASS_IDS.map((id) => [id, null])) as Weights);
  const body: ScenarioBody = { weights, bounds: cloneBounds(args.bounds), multiSectorGrowthFraction: 'inherit' };
  return {
    id: args.id,
    name: args.name.trim() || `${profileDef(args.profileId).label} proposal`,
    description: args.description ?? '',
    profileId: args.profileId,
    origin: {
      baselineId: args.baseline?.id ?? null,
      baselineKind: args.baseline?.kind ?? null,
      baselineVersion: args.baseline?.version ?? null,
      duplicatedFromScenarioId: null,
    },
    saved: body,
    draft: cloneBody(body),
    createdAt: args.now,
    savedAt: args.now,
  };
}

/** Duplicate a scenario. The copy starts from the source's current draft and is saved immediately. */
export function duplicateScenario(src: Scenario, id: string, now: string, name?: string): Scenario {
  const body = cloneBody(src.draft);
  return {
    ...src,
    id,
    name: name ?? `${src.name} (copy)`,
    origin: { ...src.origin, duplicatedFromScenarioId: src.id },
    saved: body,
    draft: cloneBody(body),
    createdAt: now,
    savedAt: now,
  };
}

export function setDraftWeight(s: Scenario, assetClass: AssetClassId, value: number | null): Scenario {
  return { ...s, draft: { ...s.draft, weights: { ...s.draft.weights, [assetClass]: value } } };
}

export function setDraftBound(s: Scenario, assetClass: AssetClassId, bound: Bound): Scenario {
  return { ...s, draft: { ...s.draft, bounds: { ...s.draft.bounds, [assetClass]: { ...bound } } } };
}

export function setDraftMultiSectorFraction(s: Scenario, f: ScenarioBody['multiSectorGrowthFraction']): Scenario {
  return { ...s, draft: { ...s.draft, multiSectorGrowthFraction: f } };
}

export function saveScenario(s: Scenario, now: string): Scenario {
  return { ...s, saved: cloneBody(s.draft), draft: cloneBody(s.draft), savedAt: now };
}

/** Discard unsaved edits, returning to the last saved state. */
export function revertScenario(s: Scenario): Scenario {
  return { ...s, draft: cloneBody(s.saved) };
}

/**
 * Reset the draft weights to a baseline's current weights (and optionally its bounds).
 * The change is a draft edit — it is not saved until the user saves.
 */
export function resetDraftToBaseline(s: Scenario, baseline: BaselineAllocation, bounds?: Bounds): Scenario {
  return {
    ...s,
    origin: { ...s.origin, baselineId: baseline.id, baselineKind: baseline.kind, baselineVersion: baseline.version },
    draft: {
      ...s.draft,
      weights: cloneWeights(baseline.weights),
      bounds: bounds ? cloneBounds(bounds) : cloneBounds(s.draft.bounds),
    },
  };
}

export function hasUnsavedChanges(s: Scenario): boolean {
  if (!weightsEqual(s.draft.weights, s.saved.weights)) return true;
  if (s.draft.multiSectorGrowthFraction !== s.saved.multiSectorGrowthFraction) return true;
  return ASSET_CLASS_IDS.some(
    (id) => s.draft.bounds[id].min !== s.saved.bounds[id].min || s.draft.bounds[id].max !== s.saved.bounds[id].max,
  );
}

/** Resolve the effective Multi-Sector growth fraction for a scenario body. */
export function effectiveFraction(body: ScenarioBody, global: number | null): number | null {
  return body.multiSectorGrowthFraction === 'inherit' ? global : body.multiSectorGrowthFraction;
}
