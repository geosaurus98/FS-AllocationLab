import type { FirmSettings, ModelHolding } from '../settings';
import type { Workbook } from './evaluator';

/**
 * Write firm settings into the Asset Allocation sheet's input cells. Cells that hold a formula in the
 * workbook are never overwritten (e.g. derived profile splits), so the workbook's own logic is kept.
 */
export function applyFirmSettings(wb: Workbook, s: FirmSettings): void {
  const AA = 'Asset Allocation';
  const put = (a1: string, v: number | string | null | undefined) => {
    const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
    const col = m[1]!.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
    if (wb.formulaAt(AA, Number(m[2]), col) !== null) return; // keep formulas
    wb.set(AA, a1, v === undefined ? null : v);
  };
  const cols: [keyof FirmSettings['tactical'], string][] = [
    ['Conservative', 'C'],
    ['Moderate', 'D'],
    ['Balanced', 'E'],
    ['Growth', 'F'],
    ['Aggressive', 'G'],
  ];
  const rows = { lowerRisk: 3, cash: 4, fi: 5, growth: 6, ae: 7, ge: 8, prop: 9, ms: 10 } as const;
  for (const [p, c] of cols) {
    const col = s.tactical[p] as unknown as Record<string, number | null>;
    for (const [k, r] of Object.entries(rows)) if (col[k] !== null && col[k] !== undefined) put(`${c}${r}`, col[k]);
  }
  const cu = s.tactical.Custom;
  put('H3', cu.lowerRisk);
  put('H4', cu.cash);
  put('H5', cu.fi);
  put('H6', cu.growth);
  put('H10', cu.ms);
  const m = s.tactical['Master Risk Alloc'];
  put('J4', m.cash);
  put('J5', m.fi);
  put('J7', m.ae);
  put('J8', m.ge);
  put('J9', m.prop);
  put('J10', m.ms);

  const p = s.property;
  const bandRows = [119, 124, 129, 134, 139];
  p.bands.forEach((b, i) => {
    const r = bandRows[i]!;
    if (i > 0) put(`H${r}`, b.from);
    if (b.to !== null) put(`I${r}`, b.to);
    put(`J${r}`, b.nzFunds);
    put(`J${r + 1}`, b.nzShares);
    put(`J${r + 2}`, b.aus);
    put(`J${r + 3}`, b.int);
  });
  const list = (first: number, items: ModelHolding[]) =>
    items.forEach((it, i) => {
      put(`B${first + i}`, it.name || null);
      put(`C${first + i}`, it.distribution);
    });
  list(119, p.nzShares);
  list(130, p.nzFunds);
  list(137, [p.ausBanded]);
  list(138, p.ausFlat);
  list(144, [p.intBanded]);
  list(145, p.intFlat);
  put('K146', p.fundLimit);
  put('K147', p.shareLimit);

  for (const [a1, v] of Object.entries(s.inputs ?? {})) put(a1, v);

  const g = s.general;
  put('Q10', g.portfolioTolerance);
  put('S10', g.portfolioRounding);
  put('U10', g.portfolioMinRemainder);
  put('W10', g.assetClassTolerance);
}
