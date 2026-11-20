import { parsePercent } from './format';
import { ASSET_CLASSES, PROFILES } from './reference';
import type { BaselineAllocation, BaselineKind, DataStatus, ProfileId, Scenario, Weights } from './types';
import { ASSET_CLASS_IDS } from './types';

/**
 * CSV format (one allocation per row):
 *
 *   record_type,profile,kind,name,status,source,version,cash,fixed_interest,australasian_equities,global_equities,property,multi_sector
 *
 *  - record_type: "baseline" or "scenario"
 *  - profile:     e.g. balanced_60_40 (see PROFILES csvKey)
 *  - kind:        strategic | target (baseline rows); proposed (scenario rows)
 *  - weights:     percentages, "%" allowed. Blank = missing (never assumed to be 0).
 */

export const CSV_META_COLUMNS = ['record_type', 'profile', 'kind', 'name', 'status', 'source', 'version'] as const;
export const CSV_COLUMNS = [...CSV_META_COLUMNS, ...ASSET_CLASSES.map((a) => a.csvKey)];

// ---------- low-level CSV ----------

export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  const s = text.replace(/^﻿/, '');
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (inQuotes) {
      if (c === '"') {
        if (s[i + 1] === '"') {
          field += '"';
          i++;
        } else inQuotes = false;
      } else field += c;
    } else if (c === '"') inQuotes = true;
    else if (c === ',') {
      row.push(field);
      field = '';
    } else if (c === '\n' || c === '\r') {
      if (c === '\r' && s[i + 1] === '\n') i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else field += c;
  }
  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows.filter((r) => r.some((f) => f.trim() !== ''));
}

function esc(v: string): string {
  return /[",\n\r]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
}

export function toCsv(rows: (string | number | null)[][]): string {
  return rows.map((r) => r.map((v) => esc(v === null ? '' : String(v))).join(',')).join('\r\n') + '\r\n';
}

// ---------- export ----------

function weightCells(w: Weights): (number | null)[] {
  return ASSET_CLASS_IDS.map((id) => w[id]);
}

const profileCsv = (id: ProfileId) => PROFILES.find((p) => p.id === id)?.csvKey ?? id;

export function exportAllocationsCsv(baselines: BaselineAllocation[], scenarios: Scenario[]): string {
  const rows: (string | number | null)[][] = [CSV_COLUMNS];
  for (const b of baselines) {
    rows.push(['baseline', profileCsv(b.profileId), b.kind, '', b.status, b.source, b.version, ...weightCells(b.weights)]);
  }
  for (const s of scenarios) {
    rows.push(['scenario', profileCsv(s.profileId), 'proposed', s.name, 'hypothetical', 'AllocationLab scenario (saved state)', '', ...weightCells(s.saved.weights)]);
  }
  return toCsv(rows);
}

// ---------- import ----------

export interface ParsedRow {
  line: number;
  recordType: 'baseline' | 'scenario';
  profileId: ProfileId;
  kind: BaselineKind | 'proposed';
  name: string;
  status: DataStatus | 'hypothetical';
  source: string;
  weights: Weights;
}

export interface RowError {
  line: number;
  message: string;
}

export interface CsvParseResult {
  rows: ParsedRow[];
  errors: RowError[];
  /** File-level problems (bad header etc.). If non-empty, nothing should be imported. */
  fatal: string[];
}

const PROFILE_LOOKUP = new Map<string, ProfileId>();
for (const p of PROFILES) {
  PROFILE_LOOKUP.set(p.csvKey, p.id);
  PROFILE_LOOKUP.set(p.id.toLowerCase(), p.id);
  PROFILE_LOOKUP.set(p.label.toLowerCase(), p.id);
}

function normaliseHeader(h: string): string {
  return h.trim().toLowerCase().replace(/[\s-]+/g, '_');
}

/** Parse and validate CSV text. Never touches application state. */
export function parseAllocationCsv(text: string): CsvParseResult {
  const table = parseCsv(text);
  const fatal: string[] = [];
  const errors: RowError[] = [];
  const rows: ParsedRow[] = [];
  if (table.length === 0) return { rows, errors, fatal: ['The file is empty.'] };

  const header = (table[0] ?? []).map(normaliseHeader);
  const col = new Map(header.map((h, i) => [h, i]));
  const required = ['record_type', 'profile', 'kind', ...ASSET_CLASSES.map((a) => a.csvKey)];
  const missingCols = required.filter((c) => !col.has(c));
  if (missingCols.length) fatal.push(`Missing required column(s): ${missingCols.join(', ')}.`);
  const dupes = header.filter((h, i) => header.indexOf(h) !== i);
  if (dupes.length) fatal.push(`Duplicate column(s): ${[...new Set(dupes)].join(', ')}.`);
  if (fatal.length) return { rows, errors, fatal };

  const get = (r: string[], c: string) => {
    const i = col.get(c);
    return i === undefined ? '' : (r[i] ?? '').trim();
  };

  table.slice(1).forEach((r, idx) => {
    const line = idx + 2;
    const rowErrors: string[] = [];
    const rt = get(r, 'record_type').toLowerCase();
    if (rt !== 'baseline' && rt !== 'scenario') rowErrors.push(`record_type must be "baseline" or "scenario" (got "${get(r, 'record_type')}")`);
    const profileRaw = get(r, 'profile');
    const profileId = PROFILE_LOOKUP.get(profileRaw.toLowerCase());
    if (!profileId) rowErrors.push(`unknown profile "${profileRaw}"`);
    const kindRaw = get(r, 'kind').toLowerCase();
    let kind: ParsedRow['kind'] | null = null;
    if (rt === 'baseline') {
      if (kindRaw === 'strategic' || kindRaw === 'target') kind = kindRaw;
      else rowErrors.push(`baseline rows need kind "strategic" or "target" (got "${get(r, 'kind')}")`);
    } else if (rt === 'scenario') {
      if (kindRaw === '' || kindRaw === 'proposed') kind = 'proposed';
      else rowErrors.push(`scenario rows need kind "proposed" (got "${get(r, 'kind')}")`);
    }
    const statusRaw = get(r, 'status').toLowerCase();
    const allowed = ['approved', 'illustrative', 'synthetic', 'draft', 'missing', 'hypothetical'];
    let status: ParsedRow['status'] = rt === 'scenario' ? 'hypothetical' : 'draft';
    if (statusRaw) {
      if (allowed.includes(statusRaw)) status = statusRaw as ParsedRow['status'];
      else rowErrors.push(`unknown status "${get(r, 'status')}"`);
    }

    const weights = {} as Weights;
    for (const ac of ASSET_CLASSES) {
      const p = parsePercent(get(r, ac.csvKey));
      if (!p.ok) {
        rowErrors.push(`${ac.label}: ${p.error}`);
        weights[ac.id] = null;
      } else weights[ac.id] = p.value;
    }

    if (rowErrors.length || !profileId || !kind) {
      errors.push({ line, message: rowErrors.join('; ') });
      return;
    }
    rows.push({
      line,
      recordType: rt as ParsedRow['recordType'],
      profileId,
      kind,
      name: get(r, 'name'),
      status,
      source: get(r, 'source'),
      weights,
    });
  });

  return { rows, errors, fatal };
}

// ---------- import planning ----------

export type ImportAction =
  | 'create-scenario'
  | 'propose-baseline-revision'
  | 'skip';

export interface ImportPlanItem {
  row: ParsedRow;
  /** Actions the user may choose for this row. */
  options: ImportAction[];
  /** Default action — always non-destructive. */
  defaultAction: ImportAction;
  /** Messages explaining what will/won't happen. */
  notes: string[];
  /** Existing baseline this row would revise (if any). */
  existing: BaselineAllocation | null;
  /** True when the row's weights differ from the existing baseline. */
  differsFromExisting: boolean;
}

/**
 * Decide what each parsed row is allowed to do. Rules:
 *  - Default for every row is "create-scenario" — a hypothetical copy. Baselines are untouched.
 *  - A baseline row may *optionally* be applied as a new baseline version, only if the user
 *    explicitly selects it, and only when the existing baseline is not approved.
 *  - Approved baselines can never be overwritten by import.
 *  - Import can never set a baseline's status to "approved"; such rows are imported as "draft".
 */
export function planImport(rows: ParsedRow[], baselines: BaselineAllocation[]): ImportPlanItem[] {
  return rows.map((row) => {
    const notes: string[] = [];
    if (row.recordType === 'scenario') {
      return { row, options: ['create-scenario', 'skip'], defaultAction: 'create-scenario', notes: ['Will be added as a new hypothetical scenario.'], existing: null, differsFromExisting: false };
    }
    const existing = baselines.find((b) => b.profileId === row.profileId && b.kind === row.kind) ?? null;
    const differs = existing ? ASSET_CLASS_IDS.some((id) => existing.weights[id] !== row.weights[id]) : true;
    const options: ImportAction[] = ['create-scenario'];
    if (existing?.status === 'approved') {
      notes.push(`Existing ${row.kind} baseline is APPROVED and locked — it cannot be overwritten by import. The row can only be imported as a scenario.`);
    } else {
      options.push('propose-baseline-revision');
      notes.push(
        differs
          ? `Differs from current ${row.kind} baseline (v${existing?.version ?? '?'}, ${existing?.status ?? 'none'}). Applying as a baseline revision requires explicit selection and creates a new version; the old version is kept in history.`
          : `Identical to current ${row.kind} baseline; no revision needed.`,
      );
    }
    if (row.status === 'approved') notes.push('Import cannot mark data as approved. If applied as a baseline revision it will be recorded as "draft".');
    options.push('skip');
    return { row, options, defaultAction: 'create-scenario', notes, existing, differsFromExisting: differs };
  });
}
