/**
 * Core domain types for AllocationLab.
 *
 * Conventions:
 *  - Weights are percentages (0–100), not fractions.
 *  - `null` means "missing / not supplied". It is never treated as zero.
 */

export const ASSET_CLASS_IDS = [
  'cash',
  'fixedInterest',
  'ausEquities',
  'globalEquities',
  'property',
  'multiSector',
] as const;
export type AssetClassId = (typeof ASSET_CLASS_IDS)[number];

export type AssetClassification = 'defensive' | 'growth' | 'lookThrough';

export interface AssetClassDef {
  id: AssetClassId;
  label: string;
  shortLabel: string;
  csvKey: string;
  classification: AssetClassification;
}

export const PROFILE_IDS = [
  'defensive',
  'conservative',
  'moderate',
  'balanced5050',
  'balanced6040',
  'growth',
  'highGrowth',
] as const;
export type ProfileId = (typeof PROFILE_IDS)[number];

export interface ProfileDef {
  id: ProfileId;
  label: string;
  csvKey: string;
  /**
   * Growth percentage implied by the profile's *name* (e.g. "Balanced 60/40" → 60).
   * Only set where the name itself states a split. Used for a consistency check,
   * never to change weights.
   */
  nominalGrowth: number | null;
}

/** Weight per asset class, percentage points. `null` = missing. */
export type Weights = Record<AssetClassId, number | null>;

export interface Bound {
  min: number | null;
  max: number | null;
}
export type Bounds = Record<AssetClassId, Bound>;

/**
 * Data provenance status. Nothing in the shipped fixtures is `approved`:
 * the prototype must not invent approved company allocations.
 */
export type DataStatus = 'approved' | 'illustrative' | 'synthetic' | 'draft' | 'missing';

export type BaselineKind = 'strategic' | 'target';
export type AllocationKind = BaselineKind | 'proposed';

export interface Provenance {
  status: DataStatus;
  source: string;
  version: number;
  effectiveDate: string | null;
  updatedAt: string;
  notes: string;
}

/** An approved/reference ("baseline") allocation for a profile. */
export interface BaselineAllocation extends Provenance {
  id: string;
  profileId: ProfileId;
  kind: BaselineKind;
  weights: Weights;
}

export interface ConstraintSet extends Provenance {
  id: string;
  profileId: ProfileId;
  bounds: Bounds;
}

/** Global assumption: Multi-Sector look-through growth fraction (0–1), or null when unknown. */
export interface MultiSectorAssumption extends Provenance {
  growthFraction: number | null;
}

export interface Settings {
  /** Tolerance, in percentage points, for the 100% sum check. */
  sumTolerance: number;
  /** Tolerance (pp) for comparing growth split with a profile's nominal split (e.g. 60/40). */
  nominalSplitTolerance: number;
}

/** A hypothetical proposal. Never approved; never written back to baselines. */
export interface Scenario {
  id: string;
  name: string;
  description: string;
  profileId: ProfileId;
  /** Where it was copied from (provenance), if anywhere. */
  origin: {
    baselineId: string | null;
    baselineKind: BaselineKind | null;
    baselineVersion: number | null;
    duplicatedFromScenarioId: string | null;
  };
  /** Last saved state. */
  saved: ScenarioBody;
  /** Working (unsaved) state. Equal to `saved` when there are no pending edits. */
  draft: ScenarioBody;
  createdAt: string;
  savedAt: string;
}

export interface ScenarioBody {
  weights: Weights;
  bounds: Bounds;
  /** Scenario-specific Multi-Sector growth fraction (0–1), `null` = explicitly unclassified, `'inherit'` = use the global assumption. */
  multiSectorGrowthFraction: number | null | 'inherit';
}

export type Severity = 'error' | 'warning' | 'info';

export interface ValidationIssue {
  code: ValidationCode;
  severity: Severity;
  message: string;
  assetClass?: AssetClassId;
}

export type ValidationCode =
  | 'MISSING_WEIGHT'
  | 'INVALID_NUMBER'
  | 'NEGATIVE_WEIGHT'
  | 'WEIGHT_OVER_100'
  | 'SUM_NOT_100'
  | 'BELOW_MIN'
  | 'ABOVE_MAX'
  | 'BELOW_MIN_CASH'
  | 'MIN_CASH_NOT_SET'
  | 'BOUND_INVALID'
  | 'MULTI_SECTOR_UNCLASSIFIED'
  | 'NOMINAL_SPLIT_MISMATCH'
  | 'NO_DATA';

export type ValidationStatus = 'valid' | 'warnings' | 'invalid' | 'incomplete' | 'missing';

export interface ValidationResult {
  status: ValidationStatus;
  issues: ValidationIssue[];
}
