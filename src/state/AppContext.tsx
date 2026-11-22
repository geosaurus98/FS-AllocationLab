import { createContext, useContext, useEffect, useMemo, useReducer, useRef, useState, type ReactNode } from 'react';
import type { Action, AppState } from './state';
import { reducer } from './state';
import { loadState, safeLocalStorage, saveState } from './persistence';

interface Ctx {
  state: AppState;
  dispatch: (a: Action) => void;
  notice: string | null;
  dismissNotice: () => void;
  persistError: boolean;
}

const AppContext = createContext<Ctx | null>(null);

export function AppProvider({ children }: { children: ReactNode }) {
  const initial = useRef(loadState(safeLocalStorage()));
  const [state, dispatch] = useReducer(reducer, initial.current.state);
  const [notice, setNotice] = useState<string | null>(initial.current.notice);
  const [persistError, setPersistError] = useState(false);

  useEffect(() => {
    setPersistError(!saveState(safeLocalStorage(), state));
  }, [state]);

  const value = useMemo(
    () => ({ state, dispatch, notice, dismissNotice: () => setNotice(null), persistError }),
    [state, notice, persistError],
  );
  return <AppContext.Provider value={value}>{children}</AppContext.Provider>;
}

export function useApp(): Ctx {
  const c = useContext(AppContext);
  if (!c) throw new Error('useApp must be used inside AppProvider');
  return c;
}

export function nowIso(): string {
  return new Date().toISOString();
}

export function newId(prefix: string): string {
  const rnd =
    typeof crypto !== 'undefined' && 'randomUUID' in crypto
      ? crypto.randomUUID().slice(0, 8)
      : Math.random().toString(36).slice(2, 10);
  return `${prefix}-${Date.now().toString(36)}-${rnd}`;
}
