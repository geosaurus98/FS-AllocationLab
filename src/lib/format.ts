/** Display helpers. Formatting never changes stored values. */

/** Format a percentage weight, e.g. 38 → "38.00%". Null → "—". */
export function formatPct(v: number | null | undefined, decimals = 2): string {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  return `${fixed(v, decimals)}%`;
}

/** Format a percentage-point difference, e.g. +1.5 → "+1.50 pp". */
export function formatPp(v: number | null | undefined, signed = true, decimals = 2): string {
  if (v === null || v === undefined || !Number.isFinite(v)) return '—';
  const s = fixed(Math.abs(v), decimals);
  if (!signed) return `${s} pp`;
  if (Number(s) === 0) return `0.${'0'.repeat(decimals)} pp`;
  return `${v > 0 ? '+' : '−'}${s} pp`;
}

function fixed(v: number, decimals: number): string {
  // Avoid "-0.00".
  const r = Number(v.toFixed(decimals));
  return (Object.is(r, -0) ? 0 : r).toFixed(decimals);
}

export type ParseResult = { ok: true; value: number | null } | { ok: false; error: string };

/**
 * Parse a user- or CSV-supplied percentage. Accepts "38", "38.5", "38%", " 38 % ".
 * Empty string → null (missing). Anything else non-numeric → error.
 * Does not clamp or round.
 */
export function parsePercent(raw: string): ParseResult {
  const t = raw.trim().replace(/%$/, '').trim();
  if (t === '') return { ok: true, value: null };
  if (!/^[-+]?(\d+(\.\d*)?|\.\d+)$/.test(t)) return { ok: false, error: `"${raw.trim()}" is not a number` };
  const n = Number(t);
  if (!Number.isFinite(n)) return { ok: false, error: `"${raw.trim()}" is not a finite number` };
  return { ok: true, value: n };
}

export function formatDate(iso: string | null): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleDateString('en-NZ', { year: 'numeric', month: 'short', day: 'numeric' });
}
