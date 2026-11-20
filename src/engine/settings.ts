import defaults from './defaults/workbook-settings.json';
import type { TacticalTable } from './profiles';

/** Firm-level settings of the recommendations engine (the workbook's Asset Allocation sheet inputs). */

export interface ModelHolding {
  name: string;
  distribution: number | null;
}

export interface PropertyBand {
  from: number;
  to: number | null;
  nzFunds: number | null;
  nzShares: number | null;
  aus: number | null;
  int: number | null;
}

export interface PropertySettings {
  /** Five portfolio-pot bands (Asset Allocation H119:K142). */
  bands: PropertyBand[];
  nzShares: ModelHolding[];
  nzFunds: ModelHolding[];
  ausBanded: ModelHolding;
  ausFlat: ModelHolding[];
  intBanded: ModelHolding;
  intFlat: ModelHolding[];
  /** Pot size from which the unbanded AUS/INT rows apply (hard-coded 30000 in the workbook). */
  flatMinimumPot: number;
  /** No purchase/sale of managed funds below this amount (K146). */
  fundLimit: number | null;
  /** No purchase/sale of shares below this amount (K147). */
  shareLimit: number | null;
}

export interface GeneralSettings {
  /** Q10 — don't change an investment unless the change exceeds this fraction of its value. */
  portfolioTolerance: number | null;
  /** S10 — rounding unit for new purchases. */
  portfolioRounding: number | null;
  portfolioMinRemainder: number | null;
  assetClassTolerance: number | null;
}

export interface FirmSettings {
  source: string;
  status: 'approved' | 'draft';
  tactical: TacticalTable;
  property: PropertySettings;
  general: GeneralSettings;
  /** Every text value in Asset Allocation column B (used by COUNTIF($B:$B, name) in the property pot). */
  assetAllocationColumnB: string[];
}

export const WORKBOOK_SETTINGS = defaults as unknown as FirmSettings;
