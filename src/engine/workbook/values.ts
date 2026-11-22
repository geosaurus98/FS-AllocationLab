/** Excel runtime values for the workbook evaluator. */

export class XErr {
  constructor(public readonly code: string) {}
  toString() {
    return this.code;
  }
}
export const BLANK = Symbol('blank');
export type Blank = typeof BLANK;
export type Scalar = number | string | boolean | XErr | Blank;

export interface Ref {
  kind: 'ref';
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface Arr {
  kind: 'arr';
  rows: Scalar[][];
}

export type Value = Scalar | Ref | Arr;

export const isRef = (v: Value): v is Ref => typeof v === 'object' && v !== null && !(v instanceof XErr) && (v as Ref).kind === 'ref';
export const isArr = (v: Value): v is Arr => typeof v === 'object' && v !== null && !(v instanceof XErr) && (v as Arr).kind === 'arr';
export const isErr = (v: unknown): v is XErr => v instanceof XErr;

export const E = {
  NA: new XErr('#N/A'),
  VALUE: new XErr('#VALUE!'),
  DIV0: new XErr('#DIV/0!'),
  REF: new XErr('#REF!'),
  NAME: new XErr('#NAME?'),
  NUM: new XErr('#NUM!'),
};

export function arr(rows: Scalar[][]): Arr {
  return { kind: 'arr', rows };
}

export const dims = (a: Arr) => ({ h: a.rows.length, w: a.rows[0]?.length ?? 0 });

/** Number formatting for text conversion ("General"): up to 15 significant digits. */
export function numToText(n: number): string {
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  let s = Number(n.toPrecision(15)).toString();
  if (s.includes('e')) {
    const [m, e] = s.split('e');
    s = `${m}E${Number(e) >= 0 ? '+' : ''}${e}`;
  }
  return s;
}

export function toText(v: Scalar): string | XErr {
  if (isErr(v)) return v;
  if (v === BLANK) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numToText(v);
  return v;
}

/** Coerce for arithmetic. */
export function toNum(v: Scalar): number | XErr {
  if (isErr(v)) return v;
  if (v === BLANK) return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const t = v.trim();
  if (t === '') return E.VALUE;
  const pct = /^([-+]?\d*\.?\d+(?:[eE][-+]?\d+)?)%$/.exec(t);
  if (pct) return Number(pct[1]) / 100;
  const cleaned = t.replace(/^\$/, '').replace(/,/g, '');
  if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(cleaned)) return Number(cleaned);
  const d = parseDateText(t);
  if (d !== null) return d;
  return E.VALUE;
}

export function toBool(v: Scalar): boolean | XErr {
  if (isErr(v)) return v;
  if (v === BLANK) return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const u = v.toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  return E.VALUE;
}

// ---------- dates (Excel serial numbers, 1900 system, NZ day/month/year text) ----------

export function serialFromYMD(y: number, m: number, d: number): number {
  const t = Date.UTC(y, m - 1, d);
  return Math.round(t / 86400000) + 25569;
}

export function ymdFromSerial(s: number): { y: number; m: number; d: number } {
  const t = (Math.floor(s) - 25569) * 86400000;
  const dt = new Date(t);
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
}

/** Text → date serial using day/month/year order (en-NZ), or ISO yyyy-mm-dd. */
export function parseDateText(t: string): number | null {
  const s = t.trim();
  let m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(s);
  if (m) return valid(+m[1]!, +m[2]!, +m[3]!);
  m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(s);
  if (m) {
    let y = +m[3]!;
    if (y < 100) y += y < 30 ? 2000 : 1900;
    return valid(y, +m[2]!, +m[1]!);
  }
  return null;
  function valid(y: number, mo: number, d: number): number | null {
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const ser = serialFromYMD(y, mo, d);
    const back = ymdFromSerial(ser);
    return back.m === mo && back.d === d ? ser : null;
  }
}

// ---------- comparison ----------

/** Excel comparison: numbers < text < booleans; text is case-insensitive; blank equals 0, "" and FALSE. */
export function compare(a: Scalar, b: Scalar): number | XErr {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (a === BLANK && b === BLANK) return 0;
  if (a === BLANK) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === BLANK) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const rank = (x: number | string | boolean) => (typeof x === 'number' ? 0 : typeof x === 'string' ? 1 : 2);
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'string' && typeof b === 'string') {
    const x = a.toLowerCase(), y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof a === 'number' && typeof b === 'number') {
    // Excel compares at 15 significant digits
    if (a === b) return 0;
    const pa = Number(a.toPrecision(15)), pb = Number(b.toPrecision(15));
    return pa < pb ? -1 : pa > pb ? 1 : 0;
  }
  return a === b ? 0 : a ? 1 : -1;
}

// ---------- wildcards & criteria ----------

export function wildcardRegex(pattern: string): RegExp {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i]!;
    if (ch === '~' && i + 1 < pattern.length) {
      re += escape(pattern[++i]!);
    } else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += escape(ch);
  }
  return new RegExp(`^${re}$`, 'is');
  function escape(c: string) {
    return c.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&');
  }
}

/** Build a predicate for SUMIF/COUNTIF-style criteria. */
export function criteria(c: Scalar): (v: Scalar) => boolean {
  if (isErr(c)) return (v) => isErr(v) && v.code === c.code;
  if (typeof c === 'number') return (v) => typeof v === 'number' ? v === c : typeof v === 'string' && toNum(v) === c;
  if (typeof c === 'boolean') return (v) => v === c;
  if (c === BLANK) return (v) => v === BLANK || v === '';
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(c)!;
  const op = m[1] ?? '';
  const rhsText = m[2]!;
  const rhsNum = rhsText.trim() === '' ? null : toNum(rhsText);
  const numeric = typeof rhsNum === 'number';
  if (op === '' || op === '=') {
    if (rhsText === '') return op === '=' ? (v) => v === BLANK : (v) => v === BLANK || v === '';
    if (numeric) return (v) => (typeof v === 'number' ? v === rhsNum : typeof v === 'string' && toNum(v) === rhsNum);
    const re = wildcardRegex(rhsText);
    return (v) => typeof v === 'string' && re.test(v);
  }
  if (op === '<>') {
    if (rhsText === '') return (v) => !(v === BLANK || v === '');
    if (numeric) return (v) => !(typeof v === 'number' && v === rhsNum);
    const re = wildcardRegex(rhsText);
    return (v) => !(typeof v === 'string' && re.test(v));
  }
  return (v) => {
    if (numeric) {
      if (typeof v !== 'number') return false;
      const d = v - (rhsNum as number);
      return op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0;
    }
    if (typeof v !== 'string') return false;
    const d = compare(v, rhsText) as number;
    return op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0;
  };
}
