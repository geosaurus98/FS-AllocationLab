import { div, iferror, isErr, mul, sub, type XValue } from './excel';

/**
 * Tactical asset allocation table ("Targets", Asset Allocation!C2:J10) of the recommendations workbook.
 * Values are fractions (0.02 = 2%). Rows: Lower Risk, Cash, Fixed Income, Growth, Australasian Equities,
 * Global Equities, Property, Multi Sector.
 */

export const WORKBOOK_PROFILES = ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive', 'Custom', 'Current', 'Master Risk Alloc'] as const;
export type WorkbookProfile = (typeof WORKBOOK_PROFILES)[number];

export interface ProfileColumn {
  lowerRisk: XValue;
  cash: XValue;
  fi: XValue;
  growth: XValue;
  ae: XValue;
  ge: XValue;
  prop: XValue;
  ms: XValue;
}

/** Literal inputs of the table. `null` marks cells that the workbook computes with a formula. */
export interface ProfileInputs {
  lowerRisk: number;
  cash: number;
  fi: number;
  growth: number;
  ae: number | null;
  ge: number | null;
  prop: number | null;
  ms: number;
}

export interface TacticalTable {
  Conservative: ProfileInputs;
  Moderate: ProfileInputs;
  /** Balanced is the reference column: its AE/GE/Property split is scaled for the other profiles. */
  Balanced: ProfileInputs & { ae: number; ge: number; prop: number };
  Growth: ProfileInputs;
  Aggressive: ProfileInputs;
  /** Custom: adviser-entered lower risk / cash / FI / growth; equities and property derived. */
  Custom: { lowerRisk: number; cash: number; fi: number; growth: number; ms: number };
  /** Master Risk Alloc: per-client values fetched from the client master (FetchClientInfo). */
  'Master Risk Alloc': { cash: number; fi: number; ae: number; ge: number; prop: number; ms: number };
}

/** Values as shipped in the workbook (Asset Allocation!C3:J10). Not company-approved data. */
export const WORKBOOK_TACTICAL_TABLE: TacticalTable = {
  Conservative: { lowerRisk: 0.75, cash: 0.075, fi: 0.675, growth: 0.25, ae: null, ge: null, prop: null, ms: 0 },
  Moderate: { lowerRisk: 0.6, cash: 0.05, fi: 0.55, growth: 0.4, ae: null, ge: null, prop: 0.04, ms: 0 },
  Balanced: { lowerRisk: 0.4, cash: 0.02, fi: 0.38, growth: 0.6, ae: 0.16, ge: 0.4, prop: 0.04, ms: 0 },
  Growth: { lowerRisk: 0.2, cash: 0.02, fi: 0.18, growth: 0.8, ae: null, ge: null, prop: null, ms: 0 },
  Aggressive: { lowerRisk: 0.1, cash: 0.02, fi: 0.08, growth: 0.9, ae: null, ge: null, prop: null, ms: 0 },
  Custom: { lowerRisk: 0.3, cash: 0.02, fi: 0.28, growth: 0.7, ms: 0 },
  'Master Risk Alloc': { cash: 0.04, fi: 0.46, ae: 0.2, ge: 0.2, prop: 0.05, ms: 0.05 },
};

/** Current sector weights (Portfolio!C11, C72, C120, C168, C192, C201) used by the "Current" profile. */
export interface CurrentSectorPct {
  /** C11: "" when cash total is 0 (the workbook formula returns blank text). */
  cash: number | '';
  fi: number;
  ae: number;
  ge: number;
  prop: number;
  /** C201: "" when multi-sector total is 0. */
  ms: number | '';
}

function derivedFromBalanced(b: TacticalTable['Balanced'], growth: number, prop: XValue) {
  // =($E7/($E$6-$E$9))*(X$6-X$9)
  const k = sub(b.growth, b.prop);
  return {
    ae: mul(div(b.ae, k), sub(growth, prop)),
    ge: mul(div(b.ge, k), sub(growth, prop)),
  };
}

/** Property for profiles whose cell is a formula: = E9 / E6 * X6 */
function derivedProp(b: TacticalTable['Balanced'], growth: number): XValue {
  return mul(div(b.prop, b.growth), growth);
}

export function profileColumn(table: TacticalTable, profile: WorkbookProfile, current?: CurrentSectorPct): ProfileColumn {
  const b = table.Balanced;
  switch (profile) {
    case 'Current': {
      const c = current ?? { cash: '', fi: 0, ae: 0, ge: 0, prop: 0, ms: '' };
      // I3 =IFERROR((C11+Current_FI)/100,0) — "" + number is #VALUE!, so 0 when cash is blank.
      const lowerRisk = c.cash === '' ? 0 : (c.cash + c.fi) / 100;
      const cash = c.cash === '' ? 0 : c.cash / 100;
      // I6: SUM over references ignores "" text.
      const growth = ((c.ms === '' ? 0 : c.ms) + c.prop + c.ge + c.ae) / 100;
      return { lowerRisk, cash, fi: c.fi / 100, growth, ae: c.ae / 100, ge: c.ge / 100, prop: c.prop / 100, ms: c.ms === '' ? 0 : c.ms / 100 };
    }
    case 'Master Risk Alloc': {
      const m = table['Master Risk Alloc'];
      return { lowerRisk: m.cash + m.fi, cash: m.cash, fi: m.fi, growth: m.ae + m.ge + m.prop + m.ms, ae: m.ae, ge: m.ge, prop: m.prop, ms: m.ms };
    }
    case 'Custom': {
      const c = table.Custom;
      // H9 is a self-referencing formula (=($E9/($E$6-$E$9))*(H$6-H$9)) solved by Excel's iterative
      // calculation. Its fixed point is exactly E9/E6*H6, the same rule as the other derived columns.
      const prop = derivedProp(b, c.growth);
      const { ae, ge } = derivedFromBalanced(b, c.growth, prop);
      return { lowerRisk: c.lowerRisk, cash: c.cash, fi: c.fi, growth: c.growth, ae, ge, prop, ms: c.ms };
    }
    case 'Balanced':
      return { ...b };
    default: {
      const p = table[profile];
      const prop: XValue = p.prop ?? derivedProp(b, p.growth);
      const d = derivedFromBalanced(b, p.growth, prop);
      return { lowerRisk: p.lowerRisk, cash: p.cash, fi: p.fi, growth: p.growth, ae: p.ae ?? d.ae, ge: p.ge ?? d.ge, prop, ms: p.ms };
    }
  }
}

/** Default % targets the workbook writes into column D when a profile is selected (ComboBox1_Change). */
export function defaultTargetsPct(col: ProfileColumn): Record<'cash' | 'fi' | 'ae' | 'ge' | 'prop' | 'ms', number | null> {
  const pct = (v: XValue) => (isErr(v) ? null : iferror(mul(v, 100), 0) as number);
  return { cash: pct(col.cash), fi: pct(col.fi), ae: pct(col.ae), ge: pct(col.ge), prop: pct(col.prop), ms: pct(col.ms) };
}


/** Portfolio!C11, C72, C120, C168, C192, C201 from current sector totals (B column). */
export function currentSectorPct(totals: { cash: number; fi: number; ae: number; ge: number; prop: number; ms: number }): CurrentSectorPct {
  const all = totals.cash + totals.fi + totals.ae + totals.ge + totals.prop + totals.ms; // B202
  const pct = (v: number) => (v / all) * 100;
  return {
    cash: totals.cash === 0 ? '' : pct(totals.cash), // =IF(B11=0,"",B11/$B$202*100)
    fi: totals.fi === 0 ? 0 : pct(totals.fi),
    ae: totals.ae === 0 ? 0 : pct(totals.ae),
    ge: totals.ge === 0 ? 0 : pct(totals.ge),
    prop: totals.prop === 0 ? 0 : pct(totals.prop),
    ms: totals.ms === 0 ? '' : pct(totals.ms), // =IF(B201=0,"",…)
  };
}
