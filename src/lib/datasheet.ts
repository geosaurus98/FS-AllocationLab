/**
 * Investment datasheet ("security master") lookup: holding name → Asset Class / Sub-Asset Class.
 *
 * Built from the firm's investment datasheet (header row containing "Asset Class" plus any of
 * "OA Name", "Apex Name", "FS Name"). Matching is deliberately conservative:
 *   1. exact name (case/whitespace-insensitive) against any of the name columns;
 *   2. the same, after removing a trailing holding-detail suffix such as
 *      "(188,000.00; 3.53%; 18 Months; 24/12/2026)" that platforms append to term deposits and bonds.
 * No fuzzy or partial matching — anything else is reported as not found.
 */

export type NameSource = 'OA' | 'Apex' | 'FS';

export interface SecurityRecord {
  line: number;
  assetClass: string;
  subClass: string;
  type: string;
  names: { source: NameSource; name: string }[];
}

interface IndexEntry {
  record: SecurityRecord;
  source: NameSource;
  name: string;
}

export interface SecurityMaster {
  records: SecurityRecord[];
  /** key → entries (more than one distinct classification = ambiguous). */
  index: Map<string, IndexEntry[]>;
  skippedRows: number;
  warnings: string[];
}

export function normName(s: string): string {
  return s.replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Remove a trailing "(a; b; c; d)" detail suffix (must contain a semicolon). */
export function stripDetailSuffix(s: string): string {
  return s.replace(/\s*\([^()]*;[^()]*\)\s*$/, '').trim();
}

const PLACEHOLDER = /^(-|x|na\b.*|n\/a)$/i;

function cleanName(raw: string, source: NameSource): string[] {
  const t = raw.replace(/\s+/g, ' ').trim();
  if (!t || PLACEHOLDER.test(t)) return [];
  const out = [t];
  // OA names are "CODE: Name" — index the part after the code as well.
  if (source === 'OA') {
    const m = /^[A-Z0-9.\-/]+:\s*(.+)$/.exec(t);
    if (m?.[1]) out.push(m[1].trim());
  }
  return out;
}

export function buildSecurityMaster(rows: string[][]): SecurityMaster {
  const warnings: string[] = [];
  const headerAt = rows.findIndex((r) => {
    const h = r.map((c) => normName(c ?? ''));
    return h.includes('asset class') && (h.includes('apex name') || h.includes('fs name') || h.includes('oa name'));
  });
  if (headerAt < 0) throw new Error('Could not find a header row with "Asset Class" and an "Apex Name", "FS Name" or "OA Name" column.');
  const h = (rows[headerAt] ?? []).map((c) => normName(c ?? ''));
  const col = (n: string) => h.indexOf(n);
  const cols: { source: NameSource; i: number }[] = (
    [
      ['FS', col('fs name')],
      ['Apex', col('apex name')],
      ['OA', col('oa name')],
    ] as [NameSource, number][]
  )
    .filter(([, i]) => i >= 0)
    .map(([source, i]) => ({ source, i }));
  const classCol = col('asset class');
  const subCol = h.findIndex((x) => x === 'sub-asset class' || x === 'sub asset class' || x === 'asset sub-class' || x === 'asset sub class');
  const typeCol = col('type');

  const records: SecurityRecord[] = [];
  const index = new Map<string, IndexEntry[]>();
  let skipped = 0;
  rows.slice(headerAt + 1).forEach((r, k) => {
    const names = cols.flatMap(({ source, i }) => cleanName(r[i] ?? '', source).map((name) => ({ source, name })));
    if (names.length === 0) {
      if (r.some((c) => (c ?? '').trim())) skipped++;
      return;
    }
    const rawClass = (r[classCol] ?? '').trim();
    const rec: SecurityRecord = {
      line: headerAt + k + 2,
      assetClass: PLACEHOLDER.test(rawClass) ? '' : rawClass,
      subClass: subCol >= 0 ? ((r[subCol] ?? '').trim().replace(/^-$/, '')) : '',
      type: typeCol >= 0 ? ((r[typeCol] ?? '').trim().replace(/^-$/, '')) : '',
      names,
    };
    records.push(rec);
    const keys = new Set<string>();
    for (const n of names) {
      for (const key of [normName(n.name), normName(stripDetailSuffix(n.name))]) {
        if (keys.has(key)) continue;
        keys.add(key);
        const list = index.get(key) ?? [];
        list.push({ record: rec, source: n.source, name: n.name });
        index.set(key, list);
      }
    }
  });
  if (records.length === 0) warnings.push('No investments were found below the header row.');
  return { records, index, skippedRows: skipped, warnings };
}

export type MatchResult =
  | { status: 'matched'; record: SecurityRecord; via: NameSource; matchedName: string; stripped: boolean }
  | { status: 'ambiguous'; candidates: SecurityRecord[]; stripped: boolean }
  | { status: 'not-found' };

function resolve(entries: IndexEntry[] | undefined, stripped: boolean): MatchResult | null {
  if (!entries || entries.length === 0) return null;
  const distinct = new Map<string, SecurityRecord>();
  for (const e of entries) distinct.set(`${normName(e.record.assetClass)}|${normName(e.record.subClass)}`, e.record);
  if (distinct.size > 1) return { status: 'ambiguous', candidates: [...distinct.values()], stripped };
  const first = entries[0]!;
  return { status: 'matched', record: first.record, via: first.source, matchedName: first.name, stripped };
}

export function lookupSecurity(master: SecurityMaster, assetName: string): MatchResult {
  return (
    resolve(master.index.get(normName(assetName)), false) ??
    resolve(master.index.get(normName(stripDetailSuffix(assetName))), true) ?? { status: 'not-found' }
  );
}
