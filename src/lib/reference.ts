import type { AssetClassDef, AssetClassId, Bounds, ProfileDef, ProfileId, Weights } from './types';
import { ASSET_CLASS_IDS } from './types';

export const ASSET_CLASSES: readonly AssetClassDef[] = [
  { id: 'cash', label: 'Cash', shortLabel: 'Cash', csvKey: 'cash', classification: 'defensive' },
  { id: 'fixedInterest', label: 'Fixed Interest', shortLabel: 'Fixed Int.', csvKey: 'fixed_interest', classification: 'defensive' },
  { id: 'ausEquities', label: 'Australasian Equities', shortLabel: 'Aus. Eq.', csvKey: 'australasian_equities', classification: 'growth' },
  { id: 'globalEquities', label: 'Global Equities', shortLabel: 'Global Eq.', csvKey: 'global_equities', classification: 'growth' },
  { id: 'property', label: 'Property', shortLabel: 'Property', csvKey: 'property', classification: 'growth' },
  { id: 'multiSector', label: 'Multi-Sector', shortLabel: 'Multi-Sector', csvKey: 'multi_sector', classification: 'lookThrough' },
];

export const PROFILES: readonly ProfileDef[] = [
  { id: 'defensive', label: 'Defensive', csvKey: 'defensive', nominalGrowth: null },
  { id: 'conservative', label: 'Conservative', csvKey: 'conservative', nominalGrowth: null },
  { id: 'moderate', label: 'Moderate', csvKey: 'moderate', nominalGrowth: null },
  { id: 'balanced5050', label: 'Balanced 50/50', csvKey: 'balanced_50_50', nominalGrowth: 50 },
  { id: 'balanced6040', label: 'Balanced 60/40', csvKey: 'balanced_60_40', nominalGrowth: 60 },
  { id: 'growth', label: 'Growth', csvKey: 'growth', nominalGrowth: null },
  { id: 'highGrowth', label: 'High Growth', csvKey: 'high_growth', nominalGrowth: null },
];

const assetById = new Map(ASSET_CLASSES.map((a) => [a.id, a]));
const profileById = new Map(PROFILES.map((p) => [p.id, p]));

export function assetClass(id: AssetClassId): AssetClassDef {
  const a = assetById.get(id);
  if (!a) throw new Error(`Unknown asset class: ${id}`);
  return a;
}

export function profile(id: ProfileId): ProfileDef {
  const p = profileById.get(id);
  if (!p) throw new Error(`Unknown profile: ${id}`);
  return p;
}

export function emptyWeights(): Weights {
  return Object.fromEntries(ASSET_CLASS_IDS.map((id) => [id, null])) as Weights;
}

export function weightsFrom(values: Partial<Record<AssetClassId, number | null>>): Weights {
  const w = emptyWeights();
  for (const id of ASSET_CLASS_IDS) {
    const v = values[id];
    w[id] = v === undefined ? null : v;
  }
  return w;
}

export function emptyBounds(): Bounds {
  return Object.fromEntries(ASSET_CLASS_IDS.map((id) => [id, { min: null, max: null }])) as Bounds;
}

/** Chart colours per asset class (fixed order, validated categorical palette, light surface). */
export const ASSET_COLORS: Record<AssetClassId, string> = {
  cash: '#2a78d6',
  fixedInterest: '#eb6834',
  ausEquities: '#1baf7a',
  globalEquities: '#eda100',
  property: '#e87ba4',
  multiSector: '#008300',
};
