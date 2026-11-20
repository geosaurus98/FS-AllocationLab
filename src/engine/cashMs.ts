import { ExcelError, n, type Cell, type XValue } from './excel';
import type { HoldingRow } from './property';

/**
 * Cash rows (Portfolio!F4:F10) and Multi Sector rows (F194:F200).
 */

/** Sector change totals E72, E120, E168, E192, E201 (sum of row changes F−B). */
export interface SectorChanges {
  fi: number;
  ae: number;
  ge: number;
  prop: number;
  ms: number;
}

/**
 * F4 =IF(B4<>"",B4-E72-E120-E168-E192-E201+BD11-SUM(G5:G10),"")
 * F5:F10 =IF(G<>"",B+G,IF(B="","",B))
 * The FIRST cash row absorbs every change elsewhere; its own override (G4) is ignored.
 */
export function cashRecommendations(rows: HoldingRow[], changes: SectorChanges, addRemove: Cell): (number | null)[] {
  const otherOverrides = rows.slice(1).reduce((a, r) => a + n(r.override), 0);
  return rows.map((r, i) => {
    if (i === 0) {
      if (r.value === null) return null;
      return r.value - changes.fi - changes.ae - changes.ge - changes.prop - changes.ms + n(addRemove) - otherOverrides;
    }
    if (r.override !== null) return n(r.value) + r.override;
    return r.value === null ? null : r.value;
  });
}

/**
 * F194 =IF(B194<>"", IF(G194<>"", B194+G194, G201/COUNTIF(B193:B200,"<>0")), "")
 * The relative references shift row by row, so row k (0-based) divides G(201+k) — empty for k>0 —
 * by COUNTIF over B(193+k):B(200+k). Only the first multi-sector holding receives the target.
 *
 * `bColumn` gives the B values of Portfolio rows 193..207 (null = empty cell) so COUNTIF can be
 * evaluated exactly; `gColumn` gives G201..G207 (G201 = the multi-sector target).
 */
export function msRecommendations(rows: HoldingRow[], bColumn: (number | string | null)[], gColumn: XValue[]): (XValue | null)[] {
  return rows.map((r, k) => {
    if (r.value === null) return null;
    if (r.override !== null) return r.value + r.override;
    const window = bColumn.slice(k, k + 8); // B(193+k):B(200+k)
    // COUNTIF(range,"<>0") counts every cell that is not the number 0 — including empty cells.
    const count = window.filter((v) => v !== 0).length;
    const g = gColumn[k] ?? 0;
    if (typeof g !== 'number') return g;
    if (count === 0) return new ExcelError('#DIV/0!');
    return g / count;
  });
}
