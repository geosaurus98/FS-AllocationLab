import { createContext, useContext, useEffect, useMemo, useState, type ReactNode } from 'react';
import type { SecurityMaster } from '../lib/datasheet';
import type { ClassMap, ClientRecord, HoldingOverrides, HoldingRow, ParseOutcome } from '../lib/holdings';
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

export interface LoadedDatasheet {
  name: string;
  sheetName: string;
  master: SecurityMaster;
  loadedAt: string;
}

interface Ctx {
  datasheet: LoadedDatasheet | null;
  setDatasheet: (d: LoadedDatasheet | null) => void;
  manualClients: ClientRecord[];
  upsertManualClient: (c: ClientRecord) => void;
  removeManualClient: (clientNumber: string) => void;
  overrides: HoldingOverrides;
  setOverrides: (o: HoldingOverrides) => void;
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

interface Config {
  classMap: ClassMap;
  settings: ClientSettings;
  overrides: HoldingOverrides;
}

function loadConfig(): Config {
  try {
    const raw = safeLocalStorage()?.getItem(KEY);
    if (raw) {
      const o = JSON.parse(raw) as Partial<Config> & { settings?: Partial<ClientSettings> };
      return { classMap: o.classMap ?? {}, settings: { ...DEFAULT_SETTINGS, ...(o.settings ?? {}) }, overrides: o.overrides ?? {} };
    }
  } catch {
    /* ignore */
  }
  return { classMap: {}, settings: DEFAULT_SETTINGS, overrides: {} };
}

const ClientDataContext = createContext<Ctx | null>(null);

export function ClientDataProvider({ children }: { children: ReactNode }) {
  const [initial] = useState(loadConfig);
  const [clients, setClients] = useState<LoadedFile<ClientRecord> | null>(null);
  const [holdings, setHoldings] = useState<LoadedFile<HoldingRow> | null>(null);
  const [classMap, setClassMap] = useState<ClassMap>(initial.classMap);
  const [settings, setSettings] = useState<ClientSettings>(initial.settings);
  const [overrides, setOverrides] = useState<HoldingOverrides>(initial.overrides);
  const [datasheet, setDatasheet] = useState<LoadedDatasheet | null>(null);
  const [manualClients, setManualClients] = useState<ClientRecord[]>([]);

  // Only non-client configuration is persisted: asset-class label mapping, per-investment-name
  // class assignments and default settings. Client files, manual clients and the datasheet are not.
  useEffect(() => {
    try {
      safeLocalStorage()?.setItem(KEY, JSON.stringify({ classMap, settings, overrides }));
    } catch {
      /* ignore */
    }
  }, [classMap, settings, overrides]);

  const value = useMemo<Ctx>(
    () => ({
      datasheet,
      setDatasheet,
      manualClients,
      upsertManualClient: (c: ClientRecord) => setManualClients((cur) => [...cur.filter((x) => x.clientNumber !== c.clientNumber), c]),
      removeManualClient: (n: string) => setManualClients((cur) => cur.filter((x) => x.clientNumber !== n)),
      overrides,
      setOverrides,
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
        setManualClients([]);
      },
    }),
    [clients, holdings, classMap, settings, overrides, datasheet, manualClients],
  );
  return <ClientDataContext.Provider value={value}>{children}</ClientDataContext.Provider>;
}

export function useClientData(): Ctx {
  const c = useContext(ClientDataContext);
  if (!c) throw new Error('useClientData must be used inside ClientDataProvider');
  return c;
}

export type { LoadedFile };
