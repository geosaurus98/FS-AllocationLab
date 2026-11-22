import type { AppState } from './state';
import { loadFixtureState } from './fixtures';

export const STORAGE_KEY = 'allocationlab:state:v1';

export interface LoadResult {
  state: AppState;
  notice: string | null;
}

/** Minimal shape check — enough to reject corrupt or foreign data without crashing. */
export function isAppState(v: unknown): v is AppState {
  if (!v || typeof v !== 'object') return false;
  const o = v as Record<string, unknown>;
  return (
    o.schemaVersion === 1 &&
    Array.isArray(o.baselines) &&
    Array.isArray(o.baselineHistory) &&
    Array.isArray(o.constraints) &&
    Array.isArray(o.constraintHistory) &&
    Array.isArray(o.scenarios) &&
    Array.isArray(o.auditLog) &&
    Array.isArray(o.multiSectorHistory) &&
    typeof o.multiSector === 'object' &&
    typeof o.settings === 'object'
  );
}

export function loadState(storage: Pick<Storage, 'getItem'> | null): LoadResult {
  try {
    const raw = storage?.getItem(STORAGE_KEY);
    if (!raw) return { state: loadFixtureState(), notice: null };
    const parsed: unknown = JSON.parse(raw);
    if (!isAppState(parsed)) {
      return { state: loadFixtureState(), notice: 'Saved data was in an unrecognised format and was not loaded. Demo data is shown instead; the saved copy has not been deleted.' };
    }
    return { state: parsed, notice: null };
  } catch {
    return { state: loadFixtureState(), notice: 'Saved data could not be read. Demo data is shown instead.' };
  }
}

export function saveState(storage: Pick<Storage, 'setItem'> | null, state: AppState): boolean {
  try {
    storage?.setItem(STORAGE_KEY, JSON.stringify(state));
    return true;
  } catch {
    return false;
  }
}

export function safeLocalStorage(): Storage | null {
  try {
    return typeof window !== 'undefined' ? window.localStorage : null;
  } catch {
    return null;
  }
}
