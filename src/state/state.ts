import type {
  BaselineAllocation,
  BaselineKind,
  Bound,
  Bounds,
  ConstraintSet,
  DataStatus,
  MultiSectorAssumption,
  ProfileId,
  Scenario,
  ScenarioBody,
  Settings,
  Weights,
  AssetClassId,
} from '../lib/types';
import { cloneWeights } from '../lib/allocation';
import {
  cloneBounds,
  duplicateScenario,
  resetDraftToBaseline,
  revertScenario,
  saveScenario,
  setDraftBound,
  setDraftMultiSectorFraction,
  setDraftWeight,
} from '../lib/scenarios';

export interface AuditEntry {
  at: string;
  action: string;
  detail: string;
}

export interface AppState {
  schemaVersion: 1;
  /** Current version of each baseline (one per profile × kind). */
  baselines: BaselineAllocation[];
  /** Superseded baseline versions, oldest first. */
  baselineHistory: BaselineAllocation[];
  constraints: ConstraintSet[];
  constraintHistory: ConstraintSet[];
  multiSector: MultiSectorAssumption;
  multiSectorHistory: MultiSectorAssumption[];
  settings: Settings;
  scenarios: Scenario[];
  auditLog: AuditEntry[];
}

export interface BaselineRevision {
  profileId: ProfileId;
  kind: BaselineKind;
  weights: Weights;
  /** Revisions may not set `approved`; approval is a separate, explicit action. */
  status: Exclude<DataStatus, 'approved'>;
  source: string;
  notes: string;
  effectiveDate: string | null;
}

export type Action =
  | { type: 'scenario/add'; scenario: Scenario }
  | { type: 'scenario/duplicate'; id: string; newId: string; now: string }
  | { type: 'scenario/setWeight'; id: string; assetClass: AssetClassId; value: number | null }
  | { type: 'scenario/setBound'; id: string; assetClass: AssetClassId; bound: Bound }
  | { type: 'scenario/setMsFraction'; id: string; value: ScenarioBody['multiSectorGrowthFraction'] }
  | { type: 'scenario/save'; id: string; now: string }
  | { type: 'scenario/revert'; id: string }
  | { type: 'scenario/resetToBaseline'; id: string; kind: BaselineKind; includeConstraints: boolean }
  | { type: 'scenario/meta'; id: string; name: string; description: string }
  | { type: 'scenario/delete'; id: string }
  | { type: 'baseline/revise'; revision: BaselineRevision; now: string }
  | { type: 'baseline/approve'; profileId: ProfileId; kind: BaselineKind; note: string; now: string }
  | { type: 'constraints/revise'; profileId: ProfileId; bounds: Bounds; status: Exclude<DataStatus, 'approved'>; source: string; now: string }
  | { type: 'multiSector/revise'; growthFraction: number | null; status: Exclude<DataStatus, 'approved'>; source: string; now: string }
  | { type: 'settings/update'; settings: Settings; now: string }
  | { type: 'audit'; action: string; detail: string; now: string }
  | { type: 'state/replace'; state: AppState };

export function findBaseline(state: AppState, profileId: ProfileId, kind: BaselineKind): BaselineAllocation | undefined {
  return state.baselines.find((b) => b.profileId === profileId && b.kind === kind);
}

export function findConstraints(state: AppState, profileId: ProfileId): ConstraintSet | undefined {
  return state.constraints.find((c) => c.profileId === profileId);
}

export function findScenario(state: AppState, id: string): Scenario | undefined {
  return state.scenarios.find((s) => s.id === id);
}

/** Why a baseline cannot be revised, or null if it can. */
export function baselineLockReason(b: BaselineAllocation | undefined): string | null {
  if (!b) return null;
  if (b.status === 'approved') {
    return 'This baseline is approved and locked. Approved values are never overwritten; create a scenario to propose changes.';
  }
  return null;
}

function audit(state: AppState, now: string, action: string, detail: string): AuditEntry[] {
  return [...state.auditLog, { at: now, action, detail }].slice(-500);
}

function mapScenario(state: AppState, id: string, f: (s: Scenario) => Scenario): AppState {
  return { ...state, scenarios: state.scenarios.map((s) => (s.id === id ? f(s) : s)) };
}

/** Pure reducer. Baselines are only changed by explicit `baseline/*` actions. */
export function reducer(state: AppState, action: Action): AppState {
  switch (action.type) {
    case 'scenario/add':
      return { ...state, scenarios: [...state.scenarios, action.scenario] };

    case 'scenario/duplicate': {
      const src = findScenario(state, action.id);
      if (!src) return state;
      const copy = duplicateScenario(src, action.newId, action.now);
      const idx = state.scenarios.indexOf(src);
      const scenarios = [...state.scenarios];
      scenarios.splice(idx + 1, 0, copy);
      return { ...state, scenarios };
    }

    case 'scenario/setWeight':
      return mapScenario(state, action.id, (s) => setDraftWeight(s, action.assetClass, action.value));

    case 'scenario/setBound':
      return mapScenario(state, action.id, (s) => setDraftBound(s, action.assetClass, action.bound));

    case 'scenario/setMsFraction':
      return mapScenario(state, action.id, (s) => setDraftMultiSectorFraction(s, action.value));

    case 'scenario/save':
      return mapScenario(state, action.id, (s) => saveScenario(s, action.now));

    case 'scenario/revert':
      return mapScenario(state, action.id, revertScenario);

    case 'scenario/resetToBaseline': {
      const s = findScenario(state, action.id);
      if (!s) return state;
      const base = findBaseline(state, s.profileId, action.kind);
      if (!base) return state;
      const cons = action.includeConstraints ? findConstraints(state, s.profileId) : undefined;
      return mapScenario(state, action.id, (sc) => resetDraftToBaseline(sc, base, cons?.bounds));
    }

    case 'scenario/meta':
      return mapScenario(state, action.id, (s) => ({ ...s, name: action.name.trim() || s.name, description: action.description }));

    case 'scenario/delete':
      return { ...state, scenarios: state.scenarios.filter((s) => s.id !== action.id) };

    case 'baseline/revise': {
      const r = action.revision;
      const current = findBaseline(state, r.profileId, r.kind);
      if (!current || baselineLockReason(current)) return state;
      const next: BaselineAllocation = {
        ...current,
        weights: cloneWeights(r.weights),
        status: r.status,
        source: r.source,
        notes: r.notes,
        effectiveDate: r.effectiveDate,
        version: current.version + 1,
        updatedAt: action.now,
      };
      return {
        ...state,
        baselines: state.baselines.map((b) => (b === current ? next : b)),
        baselineHistory: [...state.baselineHistory, current],
        auditLog: audit(state, action.now, 'Baseline revised', `${current.id} v${current.version} → v${next.version} (${next.status})`),
      };
    }

    case 'baseline/approve': {
      const current = findBaseline(state, action.profileId, action.kind);
      if (!current || current.status === 'approved' || current.status === 'missing') return state;
      const next: BaselineAllocation = {
        ...current,
        status: 'approved',
        notes: action.note,
        version: current.version + 1,
        updatedAt: action.now,
      };
      return {
        ...state,
        baselines: state.baselines.map((b) => (b === current ? next : b)),
        baselineHistory: [...state.baselineHistory, current],
        auditLog: audit(state, action.now, 'Baseline marked approved', `${current.id} v${next.version}: ${action.note}`),
      };
    }

    case 'constraints/revise': {
      const current = findConstraints(state, action.profileId);
      if (!current || current.status === 'approved') return state;
      const next: ConstraintSet = {
        ...current,
        bounds: cloneBounds(action.bounds),
        status: action.status,
        source: action.source,
        version: current.version + 1,
        updatedAt: action.now,
      };
      return {
        ...state,
        constraints: state.constraints.map((c) => (c === current ? next : c)),
        constraintHistory: [...state.constraintHistory, current],
        auditLog: audit(state, action.now, 'Constraints revised', `${current.id} v${current.version} → v${next.version}`),
      };
    }

    case 'multiSector/revise': {
      const current = state.multiSector;
      const next: MultiSectorAssumption = {
        ...current,
        growthFraction: action.growthFraction,
        status: action.status,
        source: action.source,
        version: current.version + 1,
        updatedAt: action.now,
      };
      return {
        ...state,
        multiSector: next,
        multiSectorHistory: [...state.multiSectorHistory, current],
        auditLog: audit(
          state,
          action.now,
          'Multi-Sector look-through revised',
          `v${current.version} → v${next.version}: growth fraction ${action.growthFraction === null ? 'unset' : action.growthFraction}`,
        ),
      };
    }

    case 'settings/update':
      return { ...state, settings: { ...action.settings }, auditLog: audit(state, action.now, 'Settings updated', JSON.stringify(action.settings)) };

    case 'audit':
      return { ...state, auditLog: audit(state, action.now, action.action, action.detail) };

    case 'state/replace':
      return action.state;
  }
}
