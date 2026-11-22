import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from 'react';
import { WORKBOOK_SETTINGS, type FirmSettings } from '../engine/settings';
import { safeLocalStorage } from './persistence';

/**
 * Firm-level recommendation settings (profile percentages, property model, limits). Not client data,
 * so they are persisted in browser storage with a version history.
 */

export interface SettingsVersion {
  version: number;
  savedAt: string;
  note: string;
  status: FirmSettings['status'];
  settings: FirmSettings;
}

interface Stored {
  current: SettingsVersion;
  history: SettingsVersion[];
}

const KEY = 'allocationlab:engine-settings:v1';

const initialVersion = (): SettingsVersion => ({
  version: 1,
  savedAt: '2026-11-20T00:00:00.000Z',
  note: 'Values from Recommendations Master Equity Weighting.xlsm (confirmed approved by the firm).',
  status: 'approved',
  settings: WORKBOOK_SETTINGS,
});

function load(): Stored {
  try {
    const raw = safeLocalStorage()?.getItem(KEY);
    if (raw) {
      const o = JSON.parse(raw) as Stored;
      if (o?.current?.settings?.tactical && Array.isArray(o.history)) return o;
    }
  } catch {
    /* fall through */
  }
  return { current: initialVersion(), history: [] };
}

interface Ctx {
  current: SettingsVersion;
  history: SettingsVersion[];
  save: (settings: FirmSettings, note: string, status: FirmSettings['status']) => void;
  resetToWorkbook: (note: string) => void;
}

const C = createContext<Ctx | null>(null);

export function EngineSettingsProvider({ children }: { children: ReactNode }) {
  const [stored, setStored] = useState<Stored>(load);
  useEffect(() => {
    try {
      safeLocalStorage()?.setItem(KEY, JSON.stringify(stored));
    } catch {
      /* ignore */
    }
  }, [stored]);
  const value = useMemo<Ctx>(() => {
    const push = (settings: FirmSettings, note: string, status: FirmSettings['status']) =>
      setStored((s) => ({
        current: { version: s.current.version + 1, savedAt: new Date().toISOString(), note, status, settings: { ...settings, status } },
        history: [...s.history, s.current].slice(-50),
      }));
    return {
      current: stored.current,
      history: stored.history,
      save: push,
      resetToWorkbook: (note) => push(WORKBOOK_SETTINGS, note, 'approved'),
    };
  }, [stored]);
  return <C.Provider value={value}>{children}</C.Provider>;
}

export function useEngineSettings(): Ctx {
  const c = useContext(C);
  if (!c) throw new Error('useEngineSettings must be used inside EngineSettingsProvider');
  return c;
}
