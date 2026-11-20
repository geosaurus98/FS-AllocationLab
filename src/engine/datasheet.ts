/**
 * Lookups into the workbook's Data sheet (investment datasheet), keyed by FS Name (column D) —
 * the name used in the Portfolio sheet. Mirrors `INDEX(Data_x, MATCH(A, Data!$D:$D, 0))`:
 * first match, case-insensitive.
 */

export interface DataRow {
  fsName: string;
  assetClass: string;
  subClass: string;
  style: string;
  type: string;
  bondType: string;
  hedged: string;
  equivalents: string;
}

export interface DataLookup {
  get(fsName: string): DataRow | undefined;
}

export function buildDataLookup(rows: DataRow[]): DataLookup {
  const m = new Map<string, DataRow>();
  for (const r of rows) {
    const k = r.fsName.trim().toLowerCase();
    if (k && !m.has(k)) m.set(k, r);
  }
  return { get: (n) => m.get(n.trim().toLowerCase()) };
}

/** Build DataRows from a raw Data-sheet grid (header row with "FS Name"). */
export function dataRowsFromGrid(grid: string[][]): DataRow[] {
  const hi = grid.findIndex((r) => r.some((c) => (c ?? '').trim().toLowerCase() === 'fs name'));
  if (hi < 0) throw new Error('Data sheet header row with "FS Name" not found.');
  const h = (grid[hi] ?? []).map((c) => (c ?? '').trim().toLowerCase());
  const col = (name: string) => h.indexOf(name);
  const c = {
    fs: col('fs name'),
    cls: col('asset class'),
    sub: col('sub-asset class'),
    style: col('style'),
    type: col('type'),
    bond: col('bond type'),
    hedged: col('hedged'),
    eq: col('equivalents'),
  };
  const cell = (r: string[], i: number) => (i < 0 ? '' : (r[i] ?? '').trim());
  return grid.slice(hi + 1).map((r) => ({
    fsName: cell(r, c.fs),
    assetClass: cell(r, c.cls),
    subClass: cell(r, c.sub),
    style: cell(r, c.style),
    type: cell(r, c.type),
    bondType: cell(r, c.bond),
    hedged: cell(r, c.hedged),
    equivalents: cell(r, c.eq),
  })).filter((r) => r.fsName);
}
