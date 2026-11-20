import type { HoldingInput, SectionKey } from './scenario';

/**
 * Port of the workbook's Apex import path (modPortfolioData.NormalizeAssetName, ParseApexTermDeposit,
 * BuildHoldingKey, FormatTermDepositDisplay, Custodian_Importing.ImportHoldingRow):
 * custodian asset names → FS names, term deposits split by rate and maturity, holdings summed per
 * FS name, and each holding placed in its Portfolio section by asset class.
 */

export interface DataSheetRow {
  oaName: string;
  apexName: string;
  fsName: string;
  assetClass: string;
}

export interface RawHolding {
  name: string;
  value: number;
  /** Custodian asset class, if supplied. Otherwise the Data sheet class of the FS name is used. */
  assetClass?: string;
}

export interface ImportedHolding {
  rawNames: string[];
  fsName: string;
  /** Name as written into the Portfolio sheet (term deposits include rate and maturity). */
  portfolioName: string;
  value: number;
  section: SectionKey | null;
  matchedBy: 'apex-name' | 'security-code' | 'term-deposit' | 'unmatched';
  rate: number | null;
  maturity: string | null;
}

export interface ImportResult {
  holdings: ImportedHolding[];
  errors: string[];
  bySection: Record<SectionKey, HoldingInput[]>;
  total: number;
}

const SECTION_BY_CLASS: Record<string, SectionKey> = {
  cash: 'cash',
  'fixed interest': 'fi',
  'australasian equities': 'ae',
  'global equities': 'ge',
  property: 'prop',
  'multi sector': 'ms',
  diversified: 'ms',
};

/** ExtractSecurityCode: the most code-like bracketed token, e.g. "(WNZ-0726)" → "WNZ-0726". */
export function extractSecurityCode(s: string): string {
  let best = '';
  let bestScore = -9999;
  const t0 = s.trim();
  // VBA scans "(" … next ")" pairs left to right
  let open = t0.indexOf('(');
  while (open >= 0) {
    const close = t0.indexOf(')', open + 1);
    if (close < 0) break;
    const cand = t0.slice(open + 1, close).trim();
    const t = cand.toUpperCase();
    open = t0.indexOf('(', close + 1);
    if (t.includes(';') || t.includes(' ')) continue;
    if (t.length < 5) continue;
    if (['PIE', 'HEDGED', 'UNHEDGED', 'CASH', 'DIST', 'RE-INVEST', 'REINVEST', 'NZD', 'AUD', 'USD', 'GBP', 'EUR', 'AR', 'SR', 'XR', 'ETF', 'FUND', 'FUNDS'].includes(t)) continue;
    let score = 0;
    if (/[0-9]/.test(t)) score += 5;
    if (t.includes('.') || t.includes('-')) score += 4;
    if (t.length >= 5 && t.length <= 10) score += 3;
    if (score > bestScore) {
      bestScore = score;
      best = t;
    }
  }
  return bestScore > 0 ? best : '';
}

export function buildApexLookups(rows: DataSheetRow[]) {
  const byName = new Map<string, string>();
  const byCode = new Map<string, string>();
  const classByFs = new Map<string, string>();
  for (const r of rows) {
    const fs = r.fsName.trim();
    const apex = r.apexName.trim();
    if (fs && !classByFs.has(fs.toLowerCase())) classByFs.set(fs.toLowerCase(), r.assetClass.trim());
    if (!apex || !fs || apex.toUpperCase() === 'X') continue;
    if (!byName.has(apex.toLowerCase())) byName.set(apex.toLowerCase(), fs);
    const code = extractSecurityCode(apex);
    if (code) byCode.set(code.toLowerCase(), fs); // later rows overwrite, as in the macro
  }
  return { byName, byCode, classByFs };
}

const TD_BANKS: [string, string][] = [
  ['Bank of New Zealand', 'BNZ Term Deposit'],
  ['ANZ Bank', 'ANZ Term Deposit'],
  ['Heartland Bank Limited', 'Heartland Bank Limited Term Deposit'],
  ['Rabobank', 'Rabobank Term Deposit'],
  ['SBS Bank', 'SBS Term Deposit'],
  ['Bank of China', 'Bank of China Term Deposit'],
];
const TD_PREFIXES: Record<string, string> = {
  ANZTD: 'ANZ Term Deposit',
  BNZTD: 'BNZ Term Deposit',
  WPCTD: 'Westpac Term Deposit',
  HBLTD: 'Heartland Bank Limited Term Deposit',
  RABON: 'Rabobank Term Deposit',
  SBSRS: 'SBS Term Deposit',
};

export function normalizeApexName(raw: string, lk: ReturnType<typeof buildApexLookups>): { name: string; matchedBy: ImportedHolding['matchedBy'] } {
  let rawName = raw.trim();
  if (/SPDR S&P 500/i.test(rawName) && !/STATE STREET SPDR S&P 500/i.test(rawName)) rawName = rawName.replace(/SPDR S&P 500/gi, 'State Street SPDR S&P 500');
  for (const [needle, td] of TD_BANKS) {
    if (rawName.toLowerCase().includes(needle.toLowerCase()) && extractSecurityCode(rawName) === '') return { name: td, matchedBy: 'term-deposit' };
  }
  const prefix = TD_PREFIXES[rawName.slice(0, 5).toUpperCase()];
  if (prefix) return { name: prefix, matchedBy: 'term-deposit' };
  const byName = lk.byName.get(rawName.toLowerCase());
  if (byName) return { name: byName, matchedBy: 'apex-name' };
  const code = extractSecurityCode(rawName);
  if (code) {
    const byCode = lk.byCode.get(code.toLowerCase());
    if (byCode) return { name: byCode, matchedBy: 'security-code' };
  }
  return { name: rawName, matchedBy: 'unmatched' };
}

/** ParseApexTermDeposit: "(188,000.00; 3.53%; 18 Months; 24/12/2026)" → rate 3.53, maturity 24/12/2026. */
export function parseApexTermDeposit(s: string): { rate: number | null; maturity: string | null } {
  const t = s.trim();
  const open = t.indexOf('(');
  const close = t.lastIndexOf(')');
  if (open < 0 || close <= open) return { rate: null, maturity: null };
  const parts = t.slice(open + 1, close).split(';');
  if (parts.length < 4) return { rate: null, maturity: null };
  const r = Number(parts[1]!.trim().replace('%', ''));
  const d = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(parts[3]!.trim());
  return {
    rate: Number.isFinite(r) && parts[1]!.trim() !== '' ? r : null,
    maturity: d ? `${d[1]!.padStart(2, '0')}/${d[2]!.padStart(2, '0')}/${d[3]}` : null,
  };
}

export function importApexHoldings(rows: RawHolding[], data: DataSheetRow[]): ImportResult {
  const lk = buildApexLookups(data);
  const errors: string[] = [];
  const merged = new Map<string, ImportedHolding>();
  for (const row of rows) {
    if (!row.name.trim()) continue;
    const { name, matchedBy } = normalizeApexName(row.name, lk);
    let rate: number | null = null;
    let maturity: string | null = null;
    const isTd = name.toLowerCase().includes('term deposit');
    if (isTd) ({ rate, maturity } = parseApexTermDeposit(row.name));
    const key = isTd ? `${name.toLowerCase()}|${rate ?? 0}|${maturity ?? ''}` : name.toLowerCase();
    const cls = (row.assetClass?.trim() || lk.classByFs.get(name.toLowerCase()) || '').toLowerCase();
    const section = SECTION_BY_CLASS[cls] ?? null;
    const portfolioName = isTd ? `${name}${rate !== null ? ` ${rate.toFixed(2)}%` : ''}${maturity ? ` ${maturity}` : ''}` : name;
    const existing = merged.get(key);
    if (existing) {
      existing.value += row.value;
      existing.rawNames.push(row.name);
    } else {
      merged.set(key, { rawNames: [row.name], fsName: name, portfolioName, value: row.value, section, matchedBy, rate, maturity });
    }
  }
  const bySection: Record<SectionKey, HoldingInput[]> = { cash: [], fi: [], ae: [], ge: [], prop: [], ms: [] };
  let total = 0;
  for (const h of merged.values()) {
    total += h.value;
    if (!h.section) {
      errors.push(`No asset class found for "${h.fsName}"${h.matchedBy === 'unmatched' ? ' (not found in the Data sheet)' : ''} — not imported.`);
      continue;
    }
    bySection[h.section].push([h.portfolioName, Math.round(h.value * 100) / 100]);
  }
  return { holdings: [...merged.values()], errors, bySection, total };
}

/** Data-sheet rows (OA/Apex/FS name + asset class) from the workbook model's Data sheet. */
export function dataRowsFromWorkbook(get: (r: number, c: number) => unknown, maxRow: number): DataSheetRow[] {
  const out: DataSheetRow[] = [];
  const s = (v: unknown) => (typeof v === 'string' ? v : v === null || v === undefined || typeof v === 'symbol' ? '' : String(v));
  for (let r = 2; r <= maxRow; r++) {
    const fs = s(get(r, 4));
    if (!fs) continue;
    out.push({ oaName: s(get(r, 2)), apexName: s(get(r, 3)), fsName: fs, assetClass: s(get(r, 5)) });
  }
  return out;
}
