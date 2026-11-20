import { describe, expect, it } from 'vitest';
import { WORKBOOK_SETTINGS, type FirmSettings } from './settings';
import { validateFirmSettings } from './settingsValidation';

const clone = (): FirmSettings => JSON.parse(JSON.stringify(WORKBOOK_SETTINGS));

describe('validateFirmSettings', () => {
  it('accepts the workbook settings without errors', () => {
    expect(validateFirmSettings(WORKBOOK_SETTINGS).filter((i) => i.severity === 'error')).toEqual([]);
  });
  it('flags a profile whose cash + FI do not equal lower risk', () => {
    const s = clone();
    s.tactical.Growth.cash = 0.05;
    expect(validateFirmSettings(s).map((i) => i.message).join(' ')).toMatch(/Growth: cash 5\.00% \+ fixed income 18\.00% ≠ lower risk 20\.00%/);
  });
  it('flags a Balanced split that does not add up to growth', () => {
    const s = clone();
    s.tactical.Balanced.ae = 0.2;
    expect(validateFirmSettings(s).some((i) => /Balanced: Australasian/.test(i.message))).toBe(true);
  });
});
