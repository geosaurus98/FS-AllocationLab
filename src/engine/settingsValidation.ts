import { isErr } from './excel';
import { profileColumn, WORKBOOK_PROFILES } from './profiles';
import type { FirmSettings } from './settings';

export interface SettingsIssue {
  severity: 'error' | 'warning';
  message: string;
}

const close = (a: number, b: number) => Math.abs(a - b) < 1e-9;
const pct = (v: number) => `${(v * 100).toFixed(2)}%`;

/** Consistency checks on the tactical table. Nothing is adjusted automatically. */
export function validateFirmSettings(s: FirmSettings): SettingsIssue[] {
  const issues: SettingsIssue[] = [];
  const t = s.tactical;
  for (const p of ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive', 'Custom'] as const) {
    const c = t[p];
    for (const [k, v] of Object.entries(c)) {
      if (v !== null && (typeof v !== 'number' || !Number.isFinite(v) || v < 0 || v > 1)) issues.push({ severity: 'error', message: `${p}: ${k} must be between 0% and 100%.` });
    }
    if (!close(c.cash + c.fi, c.lowerRisk)) issues.push({ severity: 'error', message: `${p}: cash ${pct(c.cash)} + fixed income ${pct(c.fi)} ≠ lower risk ${pct(c.lowerRisk)}.` });
    if (!close(c.lowerRisk + c.growth, 1)) issues.push({ severity: 'error', message: `${p}: lower risk ${pct(c.lowerRisk)} + growth ${pct(c.growth)} ≠ 100%.` });
  }
  const b = t.Balanced;
  if (!close(b.ae + b.ge + b.prop + b.ms, b.growth)) {
    issues.push({ severity: 'error', message: `Balanced: Australasian ${pct(b.ae)} + global ${pct(b.ge)} + property ${pct(b.prop)} + multi-sector ${pct(b.ms)} ≠ growth ${pct(b.growth)}. The other profiles' equity splits are scaled from Balanced.` });
  }
  if (close(b.growth, b.prop) || b.growth === 0) issues.push({ severity: 'error', message: 'Balanced growth must be greater than Balanced property (other profiles divide by the difference).' });
  for (const p of WORKBOOK_PROFILES) {
    if (p === 'Current') continue;
    const col = profileColumn(t, p);
    const parts = [col.cash, col.fi, col.ae, col.ge, col.prop, col.ms];
    if (parts.some(isErr)) {
      issues.push({ severity: 'error', message: `${p}: a derived value cannot be calculated.` });
      continue;
    }
    const total = (parts as number[]).reduce((a, x) => a + x, 0);
    if (!close(total, 1)) issues.push({ severity: 'warning', message: `${p}: sector weights total ${pct(total)}, not 100%.` });
  }
  const m = t['Master Risk Alloc'];
  const mt = m.cash + m.fi + m.ae + m.ge + m.prop + m.ms;
  if (!close(mt, 1)) issues.push({ severity: 'warning', message: `Master Risk Alloc defaults total ${pct(mt)}, not 100%.` });
  const bands = s.property.bands;
  for (let i = 1; i < bands.length; i++) {
    if (bands[i - 1]!.to !== null && bands[i]!.from !== bands[i - 1]!.to) issues.push({ severity: 'warning', message: `Property band ${i + 1} starts at ${bands[i]!.from} but band ${i} ends at ${bands[i - 1]!.to}; pots in the gap get no model allocation.` });
  }
  bands.forEach((bd, i) => {
    const w = (bd.nzFunds ?? 0) + (bd.nzShares ?? 0) + (bd.aus ?? 0) + (bd.int ?? 0);
    if (!close(w, 1)) issues.push({ severity: 'warning', message: `Property band ${i + 1}: weights total ${pct(w)}, not 100%.` });
  });
  return issues;
}
