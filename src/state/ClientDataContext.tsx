import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from 'react';
import type { ClassMap, ClientRecord, HoldingRow, ParseOutcome } from '../lib/holdings';
import type { BaselineKind } from '../lib/types';
import { safeLocalStorage } from './persistence';

/**
 * Client data (holdings and client details) lives ONLY in this React state — it is never written
 * to localStorage, and is cleared when the tab is closed or reloaded.
 * Only non-client configuration (asset-class mapping, defaults) is persisted.
 */

export interface ClientSettings {
  defaultCoverMonths: number;
  defaultTolerance: number;
  referenceKind: BaselineKind;
}

interface LoadedFile<T> {
  name: string;
  outcome: ParseOutcome<T>;
  loadedAt: string;
  synthetic: boolean;
}

interface Ctx {
  clients: LoadedFile<ClientRecord> | null;
  holdings: LoadedFile<HoldingRow> | null;
  setClients: (f: LoadedFile<ClientRecord> | null) => void;
  setHoldings: (f: LoadedFile<HoldingRow> | null) => void;
  classMap: ClassMap;
  setClassMap: (m: ClassMap) => void;
  settings: ClientSettings;
  setSettings: (s: ClientSettings) => void;
  clearAll: () => void;
}

const KEY = 'allocationlab:client-config:v1';
const DEFAULT_SETTINGS: ClientSettings = { defaultCoverMonths: 12, defaultTolerance: 2, referenceKind: 'strategic' };

function loadConfig(): { classMap: ClassMap; settings: ClientSettings } {
  try {
    const raw = safeLocalStorage()?.getItem(KEY);
    if (raw) {
      const o = JSON.parse(raw) as { classMap?: ClassMap; settings?: Partial<ClientSettings> };
      return { classMap: o.classMap ?? {}, settings: { ...DEFAULT_SETTINGS, ...(o.settings ?? {}) } };
    }
  } catch {
    /* ignore */
  }
  return { classMap: {}, settings: DEFAULT_SETTINGS };
}

const ClientDataContext = createContext<Ctx | null>(null);

export function ClientDataProvider({ children }: { children: ReactNode }) {
  const [initial] = useState(loadConfig);
  const [clients, setClients] = useState<LoadedFile<ClientRecord> | null>(null);
  const [holdings, setHoldings] = useState<LoadedFile<HoldingRow> | null>(null);
  const [classMap, setClassMap] = useState<ClassMap>(initial.classMap);
  const [settings, setSettings] = useState<ClientSettings>(initial.settings);

  useEffect(() => {
    try {
      safeLocalStorage()?.setItem(KEY, JSON.stringify({ classMap, settings }));
    } catch {
      /* ignore */
    }
  }, [classMap, settings]);

  const value = useMemo<Ctx>(
    () => ({
      clients,
      holdings,
      setClients,
      setHoldings,
      classMap,
      setClassMap,
      settings,
      setSettings,
      clearAll: () => {
        setClients(null);
        setHoldings(null);
      },
    }),
    [clients, holdings, classMap, settings],
  );
  return <ClientDataContext.Provider value={value}>{children}</ClientDataContext.Provider>;
}

export function useClientData(): Ctx {
  const c = useContext(ClientDataContext);
  if (!c) throw new Error('useClientData must be used inside ClientDataProvider');
  return c;
}

export type { LoadedFile };
