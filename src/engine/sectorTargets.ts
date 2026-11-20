import { add, div, eqBlank, eqZero, iferror, mul, n, notBlank, sub, type Cell, type XValue } from './excel';
import type { ProfileColumn } from './profiles';

/**
 * Sector target amounts — Portfolio!G11 (cash), G72 (fixed interest), G120 (Australasian equities),
 * G168 (global equities), G192 (property), G201 (multi-sector). Ported formula-for-formula, including
 * the workbook's mix of `=""` and `=0` blank tests, so results match the workbook exactly.
 */

export interface SectorTargetInputs {
  /** Profile column of the Targets table (fractions). */
  t: ProfileColumn;
  /** Column D % targets: Cash_Target D11, Fixed_Interest_Target D72, Multi_Sector_Target D201. */
  cashTargetPct: Cell;
  fiTargetPct: Cell;
  msTargetPct: Cell;
  /** Portfolio_Total (B204) — set by the import macros to the total holdings value. */
  portfolioTotal: Cell;
  /** Fees_and_Withdrawals (BH2) — dollars that cash must cover. */
  feesAndWithdrawals: Cell;
  /** Add_remove_to_Portfolio (BD11). */
  addRemove: Cell;
  /** Cash_Total_Override (BD12). */
  cashTotalOverride: Cell;
  /** Sector total overrides BD72 / BD120 / BD168 / BD192 (added to the current sector total). */
  fiOverride: Cell;
  aeOverride: Cell;
  geOverride: Cell;
  propOverride: Cell;
  /** Current sector totals B72 / B120 / B168 / B192. */
  current: { fi: number; ae: number; ge: number; prop: number };
  /**
   * Recommended sector totals F120 / F168 / F192 — only needed when a cash total override is set
   * (G72 then becomes the balancing item). Supplied by the per-holding engine.
   */
  recommended?: { ae: number; ge: number; prop: number };
}

export interface SectorTargets {
  cash: XValue;
  fi: XValue;
  ae: XValue;
  ge: XValue;
  prop: XValue;
  ms: XValue;
  /** Final_Portfolio_Total (F204). */
  finalPortfolioTotal: number;
}

export function finalPortfolioTotal(portfolioTotal: Cell, addRemove: Cell): number {
  // =IF(Portfolio_Total+Add_remove_to_Portfolio<>Portfolio_Total,Portfolio_Total+Add_remove_to_Portfolio,Portfolio_Total)
  const P = n(portfolioTotal);
  return P + n(addRemove) !== P ? P + n(addRemove) : P;
}

export function sectorTargets(i: SectorTargetInputs): SectorTargets {
  const { t } = i;
  const P = n(i.portfolioTotal);
  const CT = i.cashTargetPct;
  const FIT = i.fiTargetPct;
  const T2 = t.lowerRisk, T3 = t.cash, T4 = t.fi, T5 = t.growth;

  const cashPctBlankTest: XValue = eqBlank(CT) ? mul(T3, 100) : n(CT); // IF(Cash_Target="", T3*100, Cash_Target)
  const cashPctZeroTest: XValue = eqZero(CT) ? mul(T3, 100) : n(CT); //  IF(Cash_Target=0, …)
  const fiPctZeroTest: XValue = eqZero(FIT) ? mul(T4, 100) : n(FIT); //  IF(Fixed_Interest_Target=0, T4*100, FIT)
  const fiPctBlankTest: XValue = eqBlank(FIT) ? mul(T4, 100) : n(FIT); // IF(Fixed_Interest_Target="", …)

  // share(Tk) = (Tk*100)/(T2*100) * (fiPct + cashPct)/100 * base
  const defensiveShare = (Tk: XValue, cashPct: XValue, base: number): XValue =>
    mul(mul(div(mul(Tk, 100), mul(T2, 100)), div(add(fiPctZeroTest, cashPct), 100)), base);

  // ---- G11 cash
  let cash: XValue;
  if (n(i.cashTotalOverride) > 0) cash = n(i.cashTotalOverride);
  else {
    const x = iferror(defensiveShare(T3, cashPctBlankTest, P), mul(div(cashPctBlankTest, 100), P));
    const fees = n(i.feesAndWithdrawals);
    // IF(x > BH2, x, BH2); an error in x propagates through the comparison.
    cash = typeof x === 'number' ? (x > fees ? x : fees) : x;
  }

  // ---- G72 fixed interest
  let fi: XValue;
  if (notBlank(i.fiOverride)) fi = i.current.fi + n(i.fiOverride);
  else if (notBlank(i.cashTotalOverride)) {
    const r = i.recommended ?? { ae: 0, ge: 0, prop: 0 };
    fi = P - r.ae - r.ge - r.prop - n(i.cashTotalOverride) + n(i.addRemove);
  } else if (notBlank(i.addRemove)) {
    const PT = P + n(i.addRemove);
    const fiShare = mul(mul(div(mul(T4, 100), mul(T2, 100)), div(add(fiPctZeroTest, cashPctZeroTest), 100)), PT);
    const cashNormal = iferror(defensiveShare(T3, cashPctBlankTest, PT), 0);
    fi = mul(div(iferror(mul(div(sub(fiShare, sub(cash, cashNormal)), PT), 100), 0), 100), PT);
  } else {
    const fiShare = defensiveShare(T4, cashPctBlankTest, P);
    const cashNormal = iferror(defensiveShare(T3, cashPctBlankTest, P), 0);
    fi = sub(fiShare, sub(cash, cashNormal));
  }

  // ---- G120 / G168 / G192 growth sectors
  const growthSector = (Tk: XValue, override: Cell, currentTotal: number): XValue => {
    if (notBlank(override)) return currentTotal + n(override);
    const growthPct = sub(100, add(fiPctBlankTest, cashPctBlankTest)); // 100-(FIT+CT) with ="" tests
    if (notBlank(i.addRemove)) {
      const PT = P + n(i.addRemove);
      const v = mul(mul(div(mul(Tk, 100), mul(T5, 100)), div(growthPct, 100)), PT);
      return mul(div(iferror(mul(div(v, PT), 100), 0), 100), PT);
    }
    return mul(mul(div(mul(Tk, 100), mul(T5, 100)), div(growthPct, 100)), P);
  };

  const fpt = finalPortfolioTotal(i.portfolioTotal, i.addRemove);
  return {
    cash,
    fi,
    ae: growthSector(t.ae, i.aeOverride, i.current.ae),
    ge: growthSector(t.ge, i.geOverride, i.current.ge),
    prop: growthSector(t.prop, i.propOverride, i.current.prop),
    ms: mul(div(n(i.msTargetPct), 100), fpt), // =(Multi_Sector_Target/100)*Final_Portfolio_Total
    finalPortfolioTotal: fpt,
  };
}
