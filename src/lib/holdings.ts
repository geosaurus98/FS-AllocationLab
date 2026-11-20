import { cleanFloat, diffWeights, growthDefensiveSplit, type GrowthDefensiveSplit } from './allocation';
import { detectDelimiter, parseCsv } from './csv';
import { lookupSecurity, normName, type MatchResult, type SecurityMaster } from './datasheet';
import { formatPct, formatPp, parsePercent } from './format';
import { ASSET_CLASSES, PROFILES } from './reference';
import type { AssetClassId, ProfileId, Severity, Weights } from './types';
import { ASSET_CLASS_IDS } from './types';

/**
 * Client holdings analysis. Pure functions only.
 *
 * Holdings file (CSV or tab-separated):
 *   Client number (or Portfolio Id), Asset Name, Asset Value (or Asset value NZD), [Asset Class], [Asset Sub-Class]
 *   When Asset Class is absent, holdings are classified from the investment datasheet by name.
 * Clients file:
 *   Client number, Risk Profile, Growth Target %, Monthly Fees and/or Annual Fee %, Monthly Withdrawals,
 *   [Cash Cover Months], [Drift Tolerance pp], [Notes]
 *
 * Output is descriptive (actual vs reference, coverage, data problems). It never proposes trades.
 */

// ---------------------------------------------------------------- types

export interface HoldingRow {
  line: number;
  clientNumber: string;
  assetName: string;
  value: number;
  /** Asset Class from the file; '' when not supplied. */
  assetClass: string;
  subClass: string;
}

export type ClassSource = 'file' | 'datasheet' | 'override' | 'datasheet-no-class' | 'ambiguous' | 'not-found';

/** A holding after classification. `assetClass` is the label used for aggregation ('' = unclassified). */
export interface ClassifiedHolding extends HoldingRow {
  source: ClassSource;
  match: MatchResult | null;
  type: string;
}

export interface ClientRecord {
  line: number;
  clientNumber: string;
  profileId: ProfileId;
  /** Client's growth asset target, percent (0–100). */
  growthTarget: number | null;
  /** Fixed monthly fee in dollars. */
  monthlyFees: number | null;
  /** Annual fee as % of portfolio value (e.g. 1 = 1% p.a.); converted to dollars per month from total value. */
  annualFeePct: number | null;
  monthlyWithdrawals: number | null;
  /** Per-client override; null = use default. */
  cashCoverMonths: number | null;
  /** Per-client override; null = use default. */
  driftTolerance: number | null;
  notes: string;
}

export interface RowError {
  line: number;
  message: string;
}

export interface ParseOutcome<T> {
  rows: T[];
  errors: RowError[];
  warnings: RowError[];
  fatal: string[];
}

/** Label → asset class. Keys are normalised (lower-case, single spaces). `null` = deliberately unmapped. */
export type ClassMap = Record<string, AssetClassId | null>;

// ---------------------------------------------------------------- parsing helpers

const norm = (s: string) => s.trim().toLowerCase().replace(/[\s_-]+/g, ' ');

/** Parse a dollar amount: accepts "$12,345.67", "12345", "(1,000)" → -1000. Blank → null. */
export function parseMoney(raw: string): { ok: true; value: number | null } | { ok: false; error: string } {
  let t = raw.trim();
  if (t === '') return { ok: true, value: null };
  let neg = false;
  if (/^\(.*\)$/.test(t)) {
    neg = true;
    t = t.slice(1, -1);
  }
  t = t.replace(/^(NZ)?\$/i, '').replace(/,/g, '').trim();
  if (t.startsWith('-')) {
    neg = !neg;
    t = t.slice(1);
  }
  t = t.replace(/^(NZ)?\$/i, '');
  if (!/^(\d+(\.\d*)?|\.\d+)$/.test(t)) return { ok: false, error: `"${raw.trim()}" is not a dollar amount` };
  const n = Number(t);
  return { ok: true, value: neg ? -n : n };
}

function headerIndex(header: string[], aliases: string[]): number {
  const h = header.map(norm);
  for (const a of aliases) {
    const i = h.indexOf(norm(a));
    if (i >= 0) return i;
  }
  return -1;
}

const PROFILE_LOOKUP = new Map<string, ProfileId>();
for (const p of PROFILES) {
  for (const k of [p.id, p.csvKey, p.label]) PROFILE_LOOKUP.set(norm(k), p.id);
}
PROFILE_LOOKUP.set(norm('Balanced 60 40'), 'balanced6040');
PROFILE_LOOKUP.set(norm('Balanced 50 50'), 'balanced5050');

export function lookupProfile(raw: string): ProfileId | null {
  return PROFILE_LOOKUP.get(norm(raw.replace(/\//g, ' '))) ?? null;
}

// ---------------------------------------------------------------- holdings CSV

const H_COLS = {
  client: ['client number', 'client no', 'client', 'client id', 'portfolio id', 'portfolio'],
  name: ['asset name', 'asset', 'investment', 'holding', 'fs name', 'apex name'],
  value: ['asset value', 'asset value nzd', 'value nzd', 'value', 'market value'],
  cls: ['asset class'],
  sub: ['asset sub class', 'asset subclass', 'sub asset class', 'sub class'],
};

export function parseHoldingsCsv(text: string): ParseOutcome<HoldingRow> {
  const table = parseCsv(text, detectDelimiter(text));
  const out: ParseOutcome<HoldingRow> = { rows: [], errors: [], warnings: [], fatal: [] };
  if (table.length === 0) return { ...out, fatal: ['The holdings file is empty.'] };
  const header = table[0] ?? [];
  const idx = {
    client: headerIndex(header, H_COLS.client),
    name: headerIndex(header, H_COLS.name),
    value: headerIndex(header, H_COLS.value),
    cls: headerIndex(header, H_COLS.cls),
    sub: headerIndex(header, H_COLS.sub),
  };
  const missing = (['client', 'name', 'value'] as const).filter((k) => idx[k] < 0);
  if (missing.length) {
    const names = { client: 'Client number / Portfolio Id', name: 'Asset Name', value: 'Asset Value' };
    return { ...out, fatal: [`Holdings file is missing column(s): ${missing.map((m) => names[m]).join(', ')}.`] };
  }
  const cell = (r: string[], i: number) => (i < 0 ? '' : (r[i] ?? '').trim());

  table.slice(1).forEach((r, i) => {
    const line = i + 2;
    const problems: string[] = [];
    const clientNumber = cell(r, idx.client);
    if (!clientNumber) problems.push('Client number is blank');
    const money = parseMoney(cell(r, idx.value));
    if (!money.ok) problems.push(money.error);
    else if (money.value === null) problems.push('Asset Value is blank');
    const assetClass = cell(r, idx.cls);
    if (!cell(r, idx.name)) problems.push('Asset Name is blank');
    if (problems.length || !money.ok || money.value === null) {
      out.errors.push({ line, message: problems.join('; ') });
      return;
    }
    if (money.value < 0) out.warnings.push({ line, message: `Negative value ${money.value} for "${cell(r, idx.name)}" (client ${clientNumber}) — included as supplied.` });
    out.rows.push({ line, clientNumber, assetName: cell(r, idx.name), value: money.value, assetClass, subClass: cell(r, idx.sub) });
  });
  return out;
}

// ---------------------------------------------------------------- clients CSV

const C_COLS = {
  client: H_COLS.client,
  profile: ['risk profile', 'profile'],
  growth: ['growth target %', 'growth target', 'growth asset target %', 'growth asset target', 'growth %'],
  fees: ['monthly fees', 'fees per month', 'monthly fees $'],
  feePct: ['annual fee %', 'annual fees %', 'fee % p.a.', 'fees % pa', 'annual fee'],
  withdrawals: ['monthly withdrawals', 'withdrawals per month', 'withdrawals'],
  cover: ['cash cover months', 'cash cover', 'months cover'],
  tol: ['drift tolerance pp', 'drift tolerance', 'tolerance'],
  notes: ['notes', 'exception notes', 'client preferences'],
};

export function parseClientsCsv(text: string): ParseOutcome<ClientRecord> {
  const table = parseCsv(text, detectDelimiter(text));
  const out: ParseOutcome<ClientRecord> = { rows: [], errors: [], warnings: [], fatal: [] };
  if (table.length === 0) return { ...out, fatal: ['The client details file is empty.'] };
  const header = table[0] ?? [];
  const idx = Object.fromEntries(Object.entries(C_COLS).map(([k, a]) => [k, headerIndex(header, a)])) as Record<keyof typeof C_COLS, number>;
  const required: [keyof typeof C_COLS, string][] = [
    ['client', 'Client number'],
    ['profile', 'Risk Profile'],
    ['growth', 'Growth Target %'],
    ['withdrawals', 'Monthly Withdrawals'],
  ];
  const missing = required.filter(([k]) => idx[k] < 0).map(([, n]) => n);
  if (idx.fees < 0 && idx.feePct < 0) missing.push('Monthly Fees or Annual Fee %');
  if (missing.length) return { ...out, fatal: [`Client details file is missing column(s): ${missing.join(', ')}.`] };
  const cell = (r: string[], i: number) => (i < 0 ? '' : (r[i] ?? '').trim());
  const seen = new Set<string>();

  table.slice(1).forEach((r, i) => {
    const line = i + 2;
    const problems: string[] = [];
    const clientNumber = cell(r, idx.client);
    if (!clientNumber) problems.push('Client number is blank');
    else if (seen.has(clientNumber)) problems.push(`duplicate Client number ${clientNumber}`);
    const profRaw = cell(r, idx.profile);
    const profileId = lookupProfile(profRaw);
    if (!profileId) problems.push(`unknown Risk Profile "${profRaw}"`);

    const pct = (i2: number, label: string, max: number): number | null => {
      const p = parsePercent(cell(r, i2));
      if (!p.ok) {
        problems.push(`${label}: ${p.error}`);
        return null;
      }
      if (p.value !== null && (p.value < 0 || p.value > max)) problems.push(`${label} must be between 0 and ${max}`);
      return p.value;
    };
    const money = (i2: number, label: string): number | null => {
      const m = parseMoney(cell(r, i2));
      if (!m.ok) {
        problems.push(`${label}: ${m.error}`);
        return null;
      }
      if (m.value !== null && m.value < 0) problems.push(`${label} cannot be negative`);
      return m.value;
    };
    const growthTarget = pct(idx.growth, 'Growth Target %', 100);
    const monthlyFees = idx.fees >= 0 ? money(idx.fees, 'Monthly Fees') : null;
    const annualFeePct = idx.feePct >= 0 ? pct(idx.feePct, 'Annual Fee %', 100) : null;
    const monthlyWithdrawals = money(idx.withdrawals, 'Monthly Withdrawals');
    const cashCoverMonths = idx.cover >= 0 ? pct(idx.cover, 'Cash Cover Months', 120) : null;
    const driftTolerance = idx.tol >= 0 ? pct(idx.tol, 'Drift Tolerance pp', 100) : null;

    if (problems.length || !profileId) {
      out.errors.push({ line, message: problems.join('; ') });
      return;
    }
    seen.add(clientNumber);
    if (growthTarget === null) out.warnings.push({ line, message: `Client ${clientNumber}: Growth Target % is blank.` });
    if ((monthlyFees === null && annualFeePct === null) || monthlyWithdrawals === null) out.warnings.push({ line, message: `Client ${clientNumber}: fees and/or withdrawals are blank, so cash coverage cannot be fully assessed.` });
    out.rows.push({ line, clientNumber, profileId, growthTarget, monthlyFees, annualFeePct, monthlyWithdrawals, cashCoverMonths, driftTolerance, notes: cell(r, idx.notes) });
  });
  return out;
}

// ---------------------------------------------------------------- asset class mapping

/** Exact synonyms only. Anything else (e.g. "Diversified", "Gold") is left for the user to map. */
export const DEFAULT_CLASS_MAP: ClassMap = {
  cash: 'cash',
  'fixed interest': 'fixedInterest',
  'fixed income': 'fixedInterest',
  'australasian equities': 'ausEquities',
  'aus equities': 'ausEquities',
  'global equities': 'globalEquities',
  'international equities': 'globalEquities',
  property: 'property',
  'multi sector': 'multiSector',
};

export function classKey(label: string): string {
  return norm(label);
}

export function resolveClass(label: string, map: ClassMap): AssetClassId | null {
  const k = classKey(label);
  if (k in map) return map[k] ?? null;
  return DEFAULT_CLASS_MAP[k] ?? null;
}

/** Distinct Asset Class labels in the holdings with the value they carry and their current mapping. */
export const UNCLASSIFIED_LABEL = 'Unclassified holdings';

export function classLabels(rows: HoldingRow[], map: ClassMap): { key: string; label: string; value: number; count: number; mapped: AssetClassId | null; isDefault: boolean }[] {
  const acc = new Map<string, { key: string; label: string; value: number; count: number }>();
  for (const r of rows) {
    if (!r.assetClass) continue;
    const k = classKey(r.assetClass);
    const e = acc.get(k) ?? { key: k, label: r.assetClass.trim(), value: 0, count: 0 };
    e.value += r.value;
    e.count++;
    acc.set(k, e);
  }
  return [...acc.values()]
    .map((e) => ({ ...e, value: cleanFloat(e.value), mapped: resolveClass(e.label, map), isDefault: !(e.key in map) }))
    .sort((a, b) => b.value - a.value);
}

// ---------------------------------------------------------------- classification

/** Per-holding-name class overrides chosen by the user, keyed by normName(asset name). */
export type HoldingOverrides = Record<string, string>;

/**
 * Give every holding an asset-class label. Precedence:
 *   1. a user override for that asset name;
 *   2. Asset Class supplied in the holdings file;
 *   3. the investment datasheet (exact name, or name without the trailing TD/bond detail suffix).
 * Anything else is left unclassified ('') with the reason recorded.
 */
export function classifyHoldings(rows: HoldingRow[], master: SecurityMaster | null, overrides: HoldingOverrides = {}): ClassifiedHolding[] {
  return rows.map((r) => {
    const ov = overrides[normName(r.assetName)];
    const match = master ? lookupSecurity(master, r.assetName) : null;
    const type = match?.status === 'matched' ? match.record.type : '';
    if (ov) return { ...r, assetClass: ov, subClass: r.subClass, source: 'override', match, type };
    if (r.assetClass) return { ...r, source: 'file', match, type };
    if (!match || match.status === 'not-found') return { ...r, source: 'not-found', match, type };
    if (match.status === 'ambiguous') return { ...r, source: 'ambiguous', match, type };
    if (!match.record.assetClass) return { ...r, source: 'datasheet-no-class', match, type };
    return { ...r, assetClass: match.record.assetClass, subClass: r.subClass || match.record.subClass, source: 'datasheet', match, type };
  });
}

export const SOURCE_LABEL: Record<ClassSource, string> = {
  file: 'From holdings file',
  datasheet: 'Datasheet',
  override: 'Manual assignment',
  'datasheet-no-class': 'In datasheet, no asset class',
  ambiguous: 'Ambiguous in datasheet',
  'not-found': 'Not in datasheet',
};

// ---------------------------------------------------------------- aggregation & analysis

export interface ClientAggregate {
  clientNumber: string;
  totalValue: number;
  valueByClass: Record<AssetClassId, number>;
  unmapped: { label: string; value: number }[];
  unmappedValue: number;
  holdingCount: number;
  /** Weight (%) per asset class, denominator = total value including unmapped. */
  weights: Weights;
}

export function aggregateClient(clientNumber: string, holdings: HoldingRow[], map: ClassMap): ClientAggregate {
  const valueByClass = Object.fromEntries(ASSET_CLASS_IDS.map((id) => [id, 0])) as Record<AssetClassId, number>;
  const unmappedByLabel = new Map<string, number>();
  let total = 0;
  let count = 0;
  for (const h of holdings) {
    if (h.clientNumber !== clientNumber) continue;
    count++;
    total += h.value;
    const ac = h.assetClass ? resolveClass(h.assetClass, map) : null;
    const label = h.assetClass.trim() || UNCLASSIFIED_LABEL;
    if (ac) valueByClass[ac] += h.value;
    else unmappedByLabel.set(label, (unmappedByLabel.get(label) ?? 0) + h.value);
  }
  total = cleanFloat(total);
  const weights = {} as Weights;
  for (const id of ASSET_CLASS_IDS) {
    valueByClass[id] = cleanFloat(valueByClass[id]);
    weights[id] = total > 0 ? cleanFloat((valueByClass[id] / total) * 100) : null;
  }
  const unmapped = [...unmappedByLabel.entries()].map(([label, value]) => ({ label, value: cleanFloat(value) }));
  return {
    clientNumber,
    totalValue: total,
    valueByClass,
    unmapped,
    unmappedValue: cleanFloat(unmapped.reduce((a, u) => a + u.value, 0)),
    holdingCount: count,
    weights,
  };
}

export interface ClientIssue {
  severity: Severity;
  message: string;
  assetClass?: AssetClassId;
}

export interface CashCoverage {
  /** Fees per month in dollars (fixed fee, or annual % × total value ÷ 12). */
  monthlyFees: number | null;
  monthlyOutflow: number | null;
  coverMonths: number;
  requiredCash: number | null;
  actualCash: number;
  monthsCovered: number | null;
  shortfall: number | null;
  /** Cash needed for coverage as % of portfolio. */
  requiredCashPct: number | null;
  /** Reference cash weight from the profile baseline (%). */
  referenceCashPct: number | null;
  /** max(reference cash %, required cash %) — the cash level implied by the coverage rule (descriptive). */
  impliedCashPct: number | null;
}

export interface ClientAnalysis {
  client: ClientRecord | null;
  aggregate: ClientAggregate;
  reference: Weights | null;
  drift: Record<AssetClassId, number | null> | null;
  tolerance: number;
  breaches: AssetClassId[];
  split: GrowthDefensiveSplit;
  /** Actual growth − client growth target (pp). */
  growthGap: number | null;
  /** Reference baseline growth − client growth target (pp): the client's target differs from the profile model. */
  referenceGrowthGap: number | null;
  cash: CashCoverage;
  issues: ClientIssue[];
}

export interface AnalyseArgs {
  client: ClientRecord | null;
  aggregate: ClientAggregate;
  /** Profile baseline weights to measure drift against, or null if missing. */
  reference: Weights | null;
  referenceLabel: string;
  multiSectorGrowthFraction: number | null;
  defaultCoverMonths: number;
  defaultTolerance: number;
}

export function analyseClient(a: AnalyseArgs): ClientAnalysis {
  const { client, aggregate: agg, reference } = a;
  const issues: ClientIssue[] = [];
  const tolerance = client?.driftTolerance ?? a.defaultTolerance;
  const coverMonths = client?.cashCoverMonths ?? a.defaultCoverMonths;

  if (!client) issues.push({ severity: 'error', message: 'No client details (risk profile, growth target, fees, withdrawals) were supplied for this client number.' });
  if (agg.holdingCount === 0) issues.push({ severity: 'error', message: 'No holdings were supplied for this client.' });
  if (agg.holdingCount > 0 && agg.totalValue <= 0) issues.push({ severity: 'error', message: `Total holdings value is ${agg.totalValue}; weights cannot be calculated.` });
  if (agg.unmappedValue !== 0) {
    issues.push({
      severity: 'warning',
      message: `$${fmtMoney(agg.unmappedValue)} (${formatPct(agg.totalValue > 0 ? (agg.unmappedValue / agg.totalValue) * 100 : null)}) is in asset classes not mapped to the six classes (${agg.unmapped.map((u) => u.label).join(', ')}). It is excluded from class weights, which therefore do not sum to 100%.`,
    });
  }

  const split = growthDefensiveSplit(agg.weights, a.multiSectorGrowthFraction);
  if (split.unclassified > 0) issues.push({ severity: 'warning', message: `Multi-Sector holdings of ${formatPct(split.unclassified)} are unclassified because no look-through growth fraction is set.` });

  let drift: ClientAnalysis['drift'] = null;
  const breaches: AssetClassId[] = [];
  if (!reference) {
    if (client) issues.push({ severity: 'warning', message: `${a.referenceLabel} has no data, so drift cannot be measured.` });
  } else if (agg.totalValue > 0) {
    drift = diffWeights(reference, agg.weights);
    for (const ac of ASSET_CLASSES) {
      const d = drift[ac.id];
      if (reference[ac.id] === null) issues.push({ severity: 'warning', assetClass: ac.id, message: `${a.referenceLabel}: ${ac.label} weight is missing, so drift for this class cannot be measured.` });
      else if (d !== null && Math.abs(d) > tolerance + 1e-9) {
        breaches.push(ac.id);
        issues.push({ severity: 'warning', assetClass: ac.id, message: `${ac.label} is ${formatPct(agg.weights[ac.id])} vs ${formatPct(reference[ac.id])} (${formatPp(d)}), outside ±${tolerance} pp.` });
      }
    }
  }

  let growthGap: number | null = null;
  let referenceGrowthGap: number | null = null;
  if (client?.growthTarget !== null && client?.growthTarget !== undefined && agg.totalValue > 0) {
    growthGap = cleanFloat(split.growth - client.growthTarget);
    if (split.unclassified === 0 && agg.unmappedValue === 0 && Math.abs(growthGap) > tolerance + 1e-9) {
      issues.push({ severity: 'warning', message: `Growth assets are ${formatPct(split.growth)} vs the client's growth target of ${formatPct(client.growthTarget)} (${formatPp(growthGap)}).` });
    } else if (split.unclassified > 0 || agg.unmappedValue !== 0) {
      issues.push({ severity: 'info', message: 'Growth vs target is approximate because some value is unclassified or unmapped.' });
    }
    if (reference) {
      const refGrowth = growthDefensiveSplit(reference, a.multiSectorGrowthFraction);
      referenceGrowthGap = cleanFloat(refGrowth.growth - client.growthTarget);
      if (Math.abs(referenceGrowthGap) > 1e-9) {
        issues.push({
          severity: 'info',
          message: `The client's growth target (${formatPct(client.growthTarget)}) differs from the ${a.referenceLabel} growth share (${formatPct(refGrowth.growth)}). Asset-class drift is measured against the profile allocation, not a client-specific allocation.`,
        });
      }
    }
  }

  // Cash coverage
  let monthlyFees: number | null = null;
  if (client) {
    const pctFee = client.annualFeePct !== null && agg.totalValue > 0 ? (agg.totalValue * client.annualFeePct) / 100 / 12 : null;
    if (client.monthlyFees !== null || pctFee !== null) monthlyFees = cleanFloat((client.monthlyFees ?? 0) + (pctFee ?? 0));
  }
  const outflow = client && monthlyFees !== null && client.monthlyWithdrawals !== null ? cleanFloat(monthlyFees + client.monthlyWithdrawals) : null;
  // Without holdings there is no cash position to assess.
  const assessable = agg.holdingCount > 0 && outflow !== null;
  const actualCash = agg.valueByClass.cash;
  const requiredCash = outflow === null ? null : cleanFloat(outflow * coverMonths);
  const monthsCovered = !assessable ? null : outflow === 0 ? Infinity : cleanFloat(actualCash / (outflow as number));
  const shortfall = !assessable || requiredCash === null ? null : cleanFloat(Math.max(0, requiredCash - actualCash));
  const requiredCashPct = requiredCash === null || agg.totalValue <= 0 ? null : cleanFloat((requiredCash / agg.totalValue) * 100);
  const referenceCashPct = reference?.cash ?? null;
  const impliedCashPct = requiredCashPct === null ? referenceCashPct : referenceCashPct === null ? requiredCashPct : Math.max(referenceCashPct, requiredCashPct);
  if (client && outflow === null) issues.push({ severity: 'warning', message: 'Monthly fees and/or withdrawals are missing, so cash coverage cannot be assessed.' });
  if (shortfall !== null && shortfall > 0) {
    issues.push({
      severity: 'warning',
      assetClass: 'cash',
      message: `Cash of $${fmtMoney(actualCash)} covers ${monthsCovered === null ? '—' : monthsCovered.toFixed(1)} months of fees and withdrawals ($${fmtMoney(outflow ?? 0)}/month); ${coverMonths} months requires $${fmtMoney(requiredCash ?? 0)} (shortfall $${fmtMoney(shortfall)}).`,
    });
  }
  if (requiredCashPct !== null && referenceCashPct !== null && requiredCashPct > referenceCashPct + 1e-9) {
    issues.push({ severity: 'info', assetClass: 'cash', message: `Covering ${coverMonths} months of outflows needs ${formatPct(requiredCashPct)} in cash, above the profile's ${formatPct(referenceCashPct)} cash weight.` });
  }

  const order = { error: 0, warning: 1, info: 2 } as const;
  issues.sort((x, y) => order[x.severity] - order[y.severity]);

  return {
    client,
    aggregate: agg,
    reference,
    drift,
    tolerance,
    breaches,
    split,
    growthGap,
    referenceGrowthGap,
    cash: { monthlyFees, monthlyOutflow: outflow, coverMonths, requiredCash, actualCash, monthsCovered, shortfall, requiredCashPct, referenceCashPct, impliedCashPct },
    issues,
  };
}

/** All client numbers appearing in either file, sorted naturally. */
export function allClientNumbers(clients: ClientRecord[], holdings: HoldingRow[]): string[] {
  const s = new Set<string>([...clients.map((c) => c.clientNumber), ...holdings.map((h) => h.clientNumber)]);
  return [...s].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

export function fmtMoney(v: number): string {
  return v.toLocaleString('en-NZ', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
