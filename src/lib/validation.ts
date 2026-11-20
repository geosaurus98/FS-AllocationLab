import { growthDefensiveSplit, isEmptyAllocation, sumsTo100, sumWeights } from './allocation';
import { formatPct, formatPp } from './format';
import { ASSET_CLASSES } from './reference';
import type {
  Bounds,
  ProfileDef,
  Settings,
  ValidationIssue,
  ValidationResult,
  ValidationStatus,
  Weights,
} from './types';

export interface ValidateInput {
  weights: Weights;
  bounds: Bounds | null;
  multiSectorGrowthFraction: number | null;
  settings: Settings;
  /** When provided, the growth split is checked against the profile's nominal split (e.g. 60/40). */
  profile?: ProfileDef;
}

/**
 * Validate an allocation. Pure: never modifies the input and never "repairs" it.
 * Returns human-readable issues ordered error → warning → info.
 */
export function validateAllocation(input: ValidateInput): ValidationResult {
  const { weights, bounds, multiSectorGrowthFraction, settings, profile } = input;
  const issues: ValidationIssue[] = [];

  if (isEmptyAllocation(weights)) {
    return {
      status: 'missing',
      issues: [{ code: 'NO_DATA', severity: 'info', message: 'No allocation data has been supplied.' }],
    };
  }

  for (const ac of ASSET_CLASSES) {
    const w = weights[ac.id];
    if (w === null) {
      issues.push({
        code: 'MISSING_WEIGHT',
        severity: 'error',
        assetClass: ac.id,
        message: `${ac.label}: weight is missing. Enter a value (use 0 for no allocation).`,
      });
      continue;
    }
    if (!Number.isFinite(w)) {
      issues.push({ code: 'INVALID_NUMBER', severity: 'error', assetClass: ac.id, message: `${ac.label}: weight is not a valid number.` });
      continue;
    }
    if (w < 0) {
      issues.push({ code: 'NEGATIVE_WEIGHT', severity: 'error', assetClass: ac.id, message: `${ac.label}: weight cannot be negative (${formatPct(w)}).` });
    }
    if (w > 100) {
      issues.push({ code: 'WEIGHT_OVER_100', severity: 'error', assetClass: ac.id, message: `${ac.label}: weight cannot exceed 100% (${formatPct(w)}).` });
    }
  }

  const { total, complete } = sumWeights(weights);
  if (!sumsTo100(total, settings.sumTolerance)) {
    const diff = total - 100;
    issues.push({
      code: 'SUM_NOT_100',
      severity: 'error',
      message: complete
        ? `Weights total ${formatPct(total)}, which is ${formatPp(Math.abs(diff), false)} ${diff > 0 ? 'over' : 'under'} 100% (tolerance ±${settings.sumTolerance} pp). Adjust weights manually — nothing is rebalanced automatically.`
        : `Supplied weights total ${formatPct(total)}; the total cannot be confirmed until missing weights are entered.`,
    });
  }

  if (bounds) {
    for (const ac of ASSET_CLASSES) {
      const b = bounds[ac.id];
      const w = weights[ac.id];
      if (b.min !== null && b.max !== null && b.min > b.max) {
        issues.push({
          code: 'BOUND_INVALID',
          severity: 'error',
          assetClass: ac.id,
          message: `${ac.label}: minimum (${formatPct(b.min)}) is greater than maximum (${formatPct(b.max)}).`,
        });
      }
      if (w === null || !Number.isFinite(w)) continue;
      if (b.min !== null && w < b.min - 1e-9) {
        const isCash = ac.id === 'cash';
        issues.push({
          code: isCash ? 'BELOW_MIN_CASH' : 'BELOW_MIN',
          severity: 'error',
          assetClass: ac.id,
          message: isCash
            ? `Cash is ${formatPct(w)}, below the minimum Cash weight of ${formatPct(b.min)}.`
            : `${ac.label} is ${formatPct(w)}, below its minimum of ${formatPct(b.min)}.`,
        });
      }
      if (b.max !== null && w > b.max + 1e-9) {
        issues.push({
          code: 'ABOVE_MAX',
          severity: 'error',
          assetClass: ac.id,
          message: `${ac.label} is ${formatPct(w)}, above its maximum of ${formatPct(b.max)}.`,
        });
      }
    }
    if (bounds.cash.min === null) {
      issues.push({ code: 'MIN_CASH_NOT_SET', severity: 'warning', assetClass: 'cash', message: 'No minimum Cash weight is configured for this allocation.' });
    }
  } else {
    issues.push({ code: 'MIN_CASH_NOT_SET', severity: 'warning', message: 'No constraint set is configured, so min/max limits (including minimum Cash) were not checked.' });
  }

  const split = growthDefensiveSplit(weights, multiSectorGrowthFraction);
  if (split.unclassified > 0) {
    issues.push({
      code: 'MULTI_SECTOR_UNCLASSIFIED',
      severity: 'warning',
      assetClass: 'multiSector',
      message: `Multi-Sector exposure of ${formatPct(split.unclassified)} is unclassified: no look-through growth fraction is set, so it is excluded from the growth/defensive split rather than assumed.`,
    });
  }

  if (profile && profile.nominalGrowth !== null && complete && split.unclassified === 0) {
    const delta = split.growth - profile.nominalGrowth;
    if (Math.abs(delta) > settings.nominalSplitTolerance + 1e-9) {
      issues.push({
        code: 'NOMINAL_SPLIT_MISMATCH',
        severity: 'warning',
        message: `${profile.label} implies ${profile.nominalGrowth}% growth / ${100 - profile.nominalGrowth}% defensive, but this allocation is ${formatPct(split.growth)} growth / ${formatPct(split.defensive)} defensive.`,
      });
    }
  }

  return { status: statusFromIssues(issues), issues: sortIssues(issues) };
}

const severityRank = { error: 0, warning: 1, info: 2 } as const;

export function sortIssues(issues: ValidationIssue[]): ValidationIssue[] {
  return [...issues].sort((a, b) => severityRank[a.severity] - severityRank[b.severity]);
}

export function statusFromIssues(issues: ValidationIssue[]): ValidationStatus {
  if (issues.some((i) => i.code === 'MISSING_WEIGHT')) return 'incomplete';
  if (issues.some((i) => i.severity === 'error')) return 'invalid';
  if (issues.some((i) => i.severity === 'warning')) return 'warnings';
  return 'valid';
}

/** Validate a constraint set on its own (independent of weights). */
export function validateBounds(bounds: Bounds): ValidationIssue[] {
  const issues: ValidationIssue[] = [];
  let minSum = 0;
  for (const ac of ASSET_CLASSES) {
    const b = bounds[ac.id];
    for (const [k, v] of [['minimum', b.min], ['maximum', b.max]] as const) {
      if (v !== null && (!Number.isFinite(v) || v < 0 || v > 100)) {
        issues.push({ code: 'BOUND_INVALID', severity: 'error', assetClass: ac.id, message: `${ac.label}: ${k} must be between 0 and 100.` });
      }
    }
    if (b.min !== null && b.max !== null && b.min > b.max) {
      issues.push({ code: 'BOUND_INVALID', severity: 'error', assetClass: ac.id, message: `${ac.label}: minimum is greater than maximum.` });
    }
    minSum += b.min ?? 0;
  }
  if (minSum > 100 + 1e-9) {
    issues.push({ code: 'BOUND_INVALID', severity: 'error', message: `Minimums sum to ${formatPct(minSum)}, so no allocation can satisfy them.` });
  }
  return issues;
}
