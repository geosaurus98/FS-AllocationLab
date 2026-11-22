import baselinesJson from '../data/baselines.json';
import constraintsJson from '../data/constraints.json';
import assumptionsJson from '../data/assumptions.json';
import { emptyBounds, emptyWeights } from '../lib/reference';
import type {
  AssetClassId,
  BaselineAllocation,
  BaselineKind,
  Bounds,
  ConstraintSet,
  DataStatus,
  MultiSectorAssumption,
  ProfileId,
  Settings,
} from '../lib/types';
import { ASSET_CLASS_IDS, PROFILE_IDS } from '../lib/types';
import type { AppState } from './state';

const FIXTURE_DATE = '2026-11-20T00:00:00.000Z';
const STATUSES: DataStatus[] = ['approved', 'illustrative', 'synthetic', 'draft', 'missing'];

function asProfile(v: unknown): ProfileId {
  if (typeof v === 'string' && (PROFILE_IDS as readonly string[]).includes(v)) return v as ProfileId;
  throw new Error(`Fixture: unknown profile "${String(v)}"`);
}
function asStatus(v: unknown): DataStatus {
  if (typeof v === 'string' && (STATUSES as string[]).includes(v)) return v as DataStatus;
  throw new Error(`Fixture: unknown status "${String(v)}"`);
}
function asNum(v: unknown): number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  throw new Error(`Fixture: invalid number "${String(v)}"`);
}

export function baselineId(profileId: ProfileId, kind: BaselineKind): string {
  return `${profileId}:${kind}`;
}

export function constraintId(profileId: ProfileId): string {
  return `${profileId}:constraints`;
}

export function loadFixtureState(): AppState {
  const baselines: BaselineAllocation[] = baselinesJson.baselines.map((b) => {
    const profileId = asProfile(b.profileId);
    const kind = b.kind === 'strategic' || b.kind === 'target' ? b.kind : (() => { throw new Error('Fixture: bad kind'); })();
    const weights = emptyWeights();
    const src = (b.weights ?? {}) as Partial<Record<AssetClassId, unknown>>;
    for (const id of ASSET_CLASS_IDS) weights[id] = asNum(src[id]);
    return {
      id: baselineId(profileId, kind),
      profileId,
      kind,
      weights,
      status: asStatus(b.status),
      source: b.source,
      version: 1,
      effectiveDate: null,
      updatedAt: FIXTURE_DATE,
      notes: '',
    };
  });

  const constraints: ConstraintSet[] = constraintsJson.constraints.map((c) => {
    const profileId = asProfile(c.profileId);
    const bounds: Bounds = emptyBounds();
    const src = c.bounds as Partial<Record<AssetClassId, { min?: unknown; max?: unknown }>>;
    for (const id of ASSET_CLASS_IDS) {
      const b = src[id];
      if (b) bounds[id] = { min: asNum(b.min), max: asNum(b.max) };
    }
    return {
      id: constraintId(profileId),
      profileId,
      bounds,
      status: asStatus(c.status),
      source: c.source,
      version: 1,
      effectiveDate: null,
      updatedAt: FIXTURE_DATE,
      notes: '',
    };
  });

  const ms = assumptionsJson.multiSector;
  const multiSector: MultiSectorAssumption = {
    growthFraction: asNum(ms.growthFraction),
    status: asStatus(ms.status),
    source: ms.source,
    version: 1,
    effectiveDate: null,
    updatedAt: FIXTURE_DATE,
    notes: '',
  };
  const settings: Settings = { ...assumptionsJson.settings };

  return {
    schemaVersion: 1,
    baselines,
    baselineHistory: [],
    constraints,
    constraintHistory: [],
    multiSector,
    multiSectorHistory: [],
    settings,
    scenarios: [],
    auditLog: [],
  };
}
