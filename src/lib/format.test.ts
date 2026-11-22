import { describe, expect, it } from 'vitest';
import { formatPct, formatPp, parsePercent } from './format';

describe('parsePercent', () => {
  it('parses plain numbers and percent signs', () => {
    expect(parsePercent('38')).toEqual({ ok: true, value: 38 });
    expect(parsePercent(' 38.5 % ')).toEqual({ ok: true, value: 38.5 });
  });
  it('treats blank as missing, not zero', () => {
    expect(parsePercent('   ')).toEqual({ ok: true, value: null });
  });
  it('rejects text', () => {
    expect(parsePercent('abc').ok).toBe(false);
    expect(parsePercent('1,000').ok).toBe(false);
  });
});

describe('formatting', () => {
  it('formats percentages and signed pp differences', () => {
    expect(formatPct(2)).toBe('2.00%');
    expect(formatPct(null)).toBe('—');
    expect(formatPp(1.5)).toBe('+1.50 pp');
    expect(formatPp(-0.25)).toBe('−0.25 pp');
    expect(formatPp(-0.0001)).toBe('0.00 pp');
  });
});
