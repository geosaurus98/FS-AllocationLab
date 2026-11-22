/**
 * Minimal Excel value semantics needed to port the recommendations workbook faithfully.
 *
 * Input cells are `number | null` (null = blank cell). In Excel:
 *   blank = ""  → TRUE,   blank = 0 → TRUE
 *   0     = ""  → FALSE,  0     = 0 → TRUE
 * Formula errors (#DIV/0!, #VALUE!, …) are represented by ExcelError and propagate like Excel's.
 */

export type Cell = number | null;

export class ExcelError {
  constructor(public readonly code: '#DIV/0!' | '#VALUE!' | '#N/A' | '#REF!') {}
  toString(): string {
    return this.code;
  }
}

export type XValue = number | ExcelError;

export const isErr = (v: unknown): v is ExcelError => v instanceof ExcelError;

/** `cell=""` */
export const eqBlank = (c: Cell): boolean => c === null;
/** `cell=0` */
export const eqZero = (c: Cell): boolean => c === null || c === 0;
/** `cell<>""` */
export const notBlank = (c: Cell): boolean => c !== null;
/** numeric value of a cell reference (blank → 0) */
export const n = (c: Cell): number => c ?? 0;

export function div(a: XValue, b: XValue): XValue {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  if (b === 0) return new ExcelError('#DIV/0!');
  return a / b;
}

/** Arithmetic that propagates errors: op(a, b). */
export function ap(a: XValue, b: XValue, op: (x: number, y: number) => number): XValue {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  return op(a, b);
}

export const add = (a: XValue, b: XValue) => ap(a, b, (x, y) => x + y);
export const sub = (a: XValue, b: XValue) => ap(a, b, (x, y) => x - y);
export const mul = (a: XValue, b: XValue) => ap(a, b, (x, y) => x * y);

export function iferror(v: XValue, fallback: XValue): XValue {
  return isErr(v) ? fallback : v;
}

/** Excel MOD: result has the sign of the divisor. */
export function xmod(a: number, b: number): XValue {
  if (b === 0) return new ExcelError('#DIV/0!');
  return a - b * Math.floor(a / b);
}

/** Excel ROUND: half away from zero on the 15-significant-digit value; `digits` may be negative. */
export function xround(v: number, digits: number): number {
  const scaled = Number((Math.abs(v) * Math.pow(10, digits)).toPrecision(15));
  const r = Math.floor(scaled + 0.5);
  const out = digits >= 0 ? r / Math.pow(10, digits) : r * Math.pow(10, -digits);
  return v < 0 ? -out : out;
}

/** Case-insensitive equality as used by MATCH(…,0) / COUNTIF on plain text. */
export const textEq = (a: string, b: string): boolean => a.trim().toLowerCase() === b.trim().toLowerCase();
