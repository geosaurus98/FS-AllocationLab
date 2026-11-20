import { ExcelError, isErr, n, textEq, xmod, xround, type Cell, type XValue } from './excel';
import type { DataLookup } from './datasheet';
import type { FirmSettings, ModelHolding } from './settings';

/**
 * Property section — Asset Allocation!B116:D149 ("Properties") and Portfolio!F170:F191.
 */

export interface HoldingRow {
  /** Column A: FS name. */
  name: string;
  /** Column B: current value (null = blank). */
  value: Cell;
  /** Column G: adviser override (null = blank). */
  override: Cell;
}

export interface PropertyModelLine {
  group: 'nzShares' | 'nzFunds' | 'aus' | 'int';
  name: string;
  distribution: number | null;
  amount: XValue;
}

export interface PropertyModel {
  /** C117 — property target less "kept" holdings that are not model assets. */
  pot: XValue;
  /** Index of the band used (0–4), or null when the pot matches no band. */
  band: number | null;
  lines: PropertyModelLine[];
}

function inColumnB(settings: FirmSettings, name: string): boolean {
  if (settings.assetAllocationColumnB.some((b) => textEq(b, name))) return true;
  const p = settings.property;
  const models: ModelHolding[] = [...p.nzShares, ...p.nzFunds, p.ausBanded, ...p.ausFlat, p.intBanded, ...p.intFlat];
  return models.some((m) => m.name && textEq(m.name, name));
}

/** C117 =Portfolio!G192-SUMPRODUCT((G<>"")*(B)*(G=0)*(COUNTIF($B:$B, A)=0)) */
export function propertyPot(settings: FirmSettings, target: XValue, rows: HoldingRow[]): XValue {
  if (isErr(target)) return target;
  let kept = 0;
  for (const r of rows) {
    if (r.override !== null && r.override === 0 && !inColumnB(settings, r.name)) kept += n(r.value);
  }
  return target - kept;
}

export function propertyModel(settings: FirmSettings, target: XValue, rows: HoldingRow[]): PropertyModel {
  const p = settings.property;
  const pot = propertyPot(settings, target, rows);
  let band: number | null = null;
  if (!isErr(pot)) {
    const b = p.bands;
    // AND(C117>=0, C117<=I119), AND(C117>H124, C117<=I124), …, C117>H139
    if (pot >= 0 && pot <= n(b[0]!.to)) band = 0;
    else if (pot > b[1]!.from && pot <= n(b[1]!.to)) band = 1;
    else if (pot > b[2]!.from && pot <= n(b[2]!.to)) band = 2;
    else if (pot > b[3]!.from && pot <= n(b[3]!.to)) band = 3;
    else if (pot > b[4]!.from) band = 4;
  }
  const banded = (weight: 'nzShares' | 'nzFunds' | 'aus' | 'int', m: ModelHolding): XValue => {
    if (isErr(pot)) return pot;
    if (band === null) return 0;
    return pot * n(p.bands[band]![weight]) * n(m.distribution);
  };
  const flat = (m: ModelHolding): XValue => {
    if (isErr(pot)) return pot;
    return pot >= p.flatMinimumPot ? pot * n(m.distribution) : 0; // =IF(C$117>=30000,C$117*C138,0)
  };
  const lines: PropertyModelLine[] = [
    ...p.nzShares.map((m) => ({ group: 'nzShares' as const, name: m.name, distribution: m.distribution, amount: banded('nzShares', m) })),
    ...p.nzFunds.map((m) => ({ group: 'nzFunds' as const, name: m.name, distribution: m.distribution, amount: banded('nzFunds', m) })),
    { group: 'aus', name: p.ausBanded.name, distribution: p.ausBanded.distribution, amount: banded('aus', p.ausBanded) },
    ...p.ausFlat.map((m) => ({ group: 'aus' as const, name: m.name, distribution: m.distribution, amount: flat(m) })),
    { group: 'int', name: p.intBanded.name, distribution: p.intBanded.distribution, amount: banded('int', p.intBanded) },
    ...p.intFlat.map((m) => ({ group: 'int' as const, name: m.name, distribution: m.distribution, amount: flat(m) })),
  ];
  return { pot, band, lines };
}

/** INDEX(Properties, MATCH(A, INDEX(Properties,0,1), 0), 3) — first line with that name. */
function modelAmount(model: PropertyModel, name: string): XValue {
  const line = model.lines.find((l) => l.name && textEq(l.name, name));
  return line ? line.amount : new ExcelError('#N/A');
}

/**
 * Recommended value for one property holding (Portfolio!F170:F191).
 * Returns null for the "" result (blank name).
 */
export function propertyRecommendation(settings: FirmSettings, model: PropertyModel, row: HoldingRow, data: DataLookup): number | null {
  const B = n(row.value);
  if (row.override !== null) return B + row.override; // IF(G<>"", B+G, …)
  if (!row.name) return null;
  const amt = modelAmount(model, row.name);
  if (isErr(amt)) return 0; // IFERROR(…, 0): holdings not in the property model are sold
  const rounding = n(settings.general.portfolioRounding);
  const m100 = xmod(amt, 100);
  const mB = xmod(B, 100);
  if (isErr(m100) || isErr(mB)) return 0;
  const target = amt - m100 + mB;
  const type = data.get(row.name)?.type ?? ''; // K = INDEX(Data_Type, MATCH(A, Data!D:D, 0))
  const tol = B * n(settings.general.portfolioTolerance);
  const limit = type === 'Managed Fund' ? Math.max(n(settings.property.fundLimit), tol) : Math.max(n(settings.property.shareLimit), tol);
  if (amt === 0 || Math.abs(B - target) >= limit) {
    if (amt === 0) return 0;
    if (B > 0) return target;
    return xround(amt, -Math.log10(rounding)); // ROUND(amt, -LOG10(Portfolio_Rounding))
  }
  return B;
}

/**
 * Model assets the "Update Sheet" macro (Trigger_Recalculate.Recalculate) adds to the property
 * section with value 0: name has more than one word, does not contain "Direct", amount <> 0.
 */
export function propertyAssetsToInsert(model: PropertyModel): string[] {
  const out: string[] = [];
  for (const l of model.lines) {
    const name = l.name.trim();
    if (!name || name.includes('Direct')) continue;
    if (name.split(' ').length <= 1) continue;
    if (isErr(l.amount) || l.amount === 0) continue;
    if (!out.includes(name)) out.push(name);
  }
  return out;
}
