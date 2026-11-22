import { ASSET_CLASSES } from './reference';
import type { AssetClassId, Weights } from './types';
import { ASSET_CLASS_IDS } from './types';

/**
 * Rounds away binary floating-point noise (e.g. 0.1 + 0.2 = 0.30000000000000004)
 * so that comparisons against tolerances behave as an analyst would expect.
 * 1e-9 pp is far below any displayed precision.
 */
export function cleanFloat(n: number): number {
  return Math.round(n * 1e9) / 1e9;
}

export interface TotalResult {
  /** Sum of supplied (non-null) weights. */
  total: number;
  /** Asset classes whose weight is missing (null). */
  missing: AssetClassId[];
  /** True when every asset class has a value. */
  complete: boolean;
}

/** Sum allocation weights. Missing weights are reported, never treated as zero silently. */
export function sumWeights(weights: Weights): TotalResult {
  let total = 0;
  const missing: AssetClassId[] = [];
  for (const id of ASSET_CLASS_IDS) {
    const v = weights[id];
    if (v === null || !Number.isFinite(v)) missing.push(id);
    else total += v;
  }
  return { total: cleanFloat(total), missing, complete: missing.length === 0 };
}

/** Whether weights sum to 100 within `tolerance` percentage points. */
export function sumsTo100(total: number, tolerance: number): boolean {
  return Math.abs(cleanFloat(total - 100)) <= tolerance + 1e-12;
}

export interface GrowthDefensiveSplit {
  growth: number;
  defensive: number;
  /** Weight that could not be classified (Multi-Sector with unknown look-through). */
  unclassified: number;
  /** Multi-Sector contribution to growth/defensive via look-through. */
  multiSectorGrowth: number;
  multiSectorDefensive: number;
  /** Asset classes with missing weights, excluded from the split. */
  missing: AssetClassId[];
  /** Look-through fraction actually applied (null when Multi-Sector exposure is unclassified). */
  growthFractionApplied: number | null;
}

/**
 * Growth / defensive split.
 *  - Cash, Fixed Interest → defensive
 *  - Australasian Equities, Global Equities, Property → growth
 *  - Multi-Sector → split by `multiSectorGrowthFraction` (0–1). If the fraction is
 *    null/invalid, Multi-Sector weight is reported as `unclassified` — never assumed.
 */
export function growthDefensiveSplit(
  weights: Weights,
  multiSectorGrowthFraction: number | null,
): GrowthDefensiveSplit {
  let growth = 0;
  let defensive = 0;
  let unclassified = 0;
  let msGrowth = 0;
  let msDefensive = 0;
  const missing: AssetClassId[] = [];
  const fractionValid =
    multiSectorGrowthFraction !== null &&
    Number.isFinite(multiSectorGrowthFraction) &&
    multiSectorGrowthFraction >= 0 &&
    multiSectorGrowthFraction <= 1;

  for (const ac of ASSET_CLASSES) {
    const w = weights[ac.id];
    if (w === null || !Number.isFinite(w)) {
      missing.push(ac.id);
      continue;
    }
    switch (ac.classification) {
      case 'growth':
        growth += w;
        break;
      case 'defensive':
        defensive += w;
        break;
      case 'lookThrough':
        if (w === 0) break;
        if (fractionValid) {
          msGrowth = w * (multiSectorGrowthFraction as number);
          msDefensive = w - msGrowth;
          growth += msGrowth;
          defensive += msDefensive;
        } else {
          unclassified += w;
        }
        break;
    }
  }
  return {
    growth: cleanFloat(growth),
    defensive: cleanFloat(defensive),
    unclassified: cleanFloat(unclassified),
    multiSectorGrowth: cleanFloat(msGrowth),
    multiSectorDefensive: cleanFloat(msDefensive),
    missing,
    growthFractionApplied: fractionValid ? (multiSectorGrowthFraction as number) : null,
  };
}

export type WeightDiff = Record<AssetClassId, number | null>;

/**
 * Percentage-point difference `to - from` per asset class.
 * Null when either side is missing (a difference against missing data is not meaningful).
 */
export function diffWeights(from: Weights, to: Weights): WeightDiff {
  const out = {} as WeightDiff;
  for (const id of ASSET_CLASS_IDS) {
    const a = from[id];
    const b = to[id];
    out[id] = a === null || b === null ? null : cleanFloat(b - a);
  }
  return out;
}

/** True when no asset class has a weight (entire allocation missing). */
export function isEmptyAllocation(weights: Weights): boolean {
  return ASSET_CLASS_IDS.every((id) => weights[id] === null);
}

export function weightsEqual(a: Weights, b: Weights): boolean {
  return ASSET_CLASS_IDS.every((id) => a[id] === b[id]);
}

export function cloneWeights(w: Weights): Weights {
  return { ...w };
}
