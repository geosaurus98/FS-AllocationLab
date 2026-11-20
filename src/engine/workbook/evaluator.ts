import { indexToCol, MAX_ROW, parseFormula, type Node } from './parser';
import {
  arr,
  BLANK,
  compare,
  criteria,
  dims,
  E,
  isArr,
  isErr,
  isRef,
  parseDateText,
  serialFromYMD,
  toBool,
  toNum,
  toText,
  wildcardRegex,
  XErr,
  ymdFromSerial,
  type Arr,
  type Ref,
  type Scalar,
  type Value,
} from './values';

/**
 * Lazy evaluator for the recommendations workbook's formulas.
 *
 * Semantics follow legacy (pre-dynamic-array) Excel, as stored in the file and reproduced by
 * LibreOffice: ranges used where a single value is expected are implicitly intersected with the
 * formula's row/column, except inside array (Ctrl+Shift+Enter) formulas and SUMPRODUCT, which work
 * element by element. Circular references are solved by iteration, as Excel's iterative calculation
 * does for the workbook.
 */

export interface WorkbookModel {
  source: string;
  sheets: Record<string, { maxRow: number; maxCol: number; cells: [number, number, string, unknown][] }>;
  names: Record<string, string>;
  iterate: boolean;
  iterateCount: number;
  iterateDelta: number;
}

type CellDef = { k: 'v'; v: Scalar } | { k: 'f'; src: string; array: boolean };

interface Ctx {
  sheet: string;
  row: number;
  col: number;
  array: boolean;
}

const key = (sheet: string, r: number, c: number) => `${sheet}\u0000${r}\u0000${c}`;

function decode(kind: string, payload: unknown): Scalar {
  switch (kind) {
    case 'n':
      return payload as number;
    case 's':
      return payload as string;
    case 'b':
      return payload as boolean;
    case 'e':
      return new XErr(String(payload));
    default:
      return BLANK;
  }
}

export class Workbook {
  private cells = new Map<string, CellDef>();
  private inputs = new Map<string, Scalar>();
  private sheetMax = new Map<string, number>();
  private globalMaxRow = 1;
  private names = new Map<string, string>();
  private astCache = new Map<string, Node>();
  private memo = new Map<string, Scalar>();
  private prev = new Map<string, Scalar>();
  private inProgress = new Set<string>();
  private circular = false;
  private rangeCache = new Map<string, Arr>();
  private matchCache = new Map<string, Map<string, number>>();
  /** Excel serial of TODAY(). */
  today: number;

  /**
   * LibreOffice 24.2 (used to generate golden results) does not implement Excel 365's FILTER, so it
   * returns #NAME?. Tests that compare against LibreOffice set this; the app never does.
   */
  libreOfficeCompat = false;

  constructor(private model: WorkbookModel, asOf?: Date) {
    for (const [name, s] of Object.entries(model.sheets)) {
      this.sheetMax.set(name, s.maxRow);
      this.globalMaxRow = Math.max(this.globalMaxRow, s.maxRow);
      for (const [r, c, k, p] of s.cells) {
        if (k === 'f' || k === 'a') this.cells.set(key(name, r, c), { k: 'f', src: String(p), array: k === 'a' });
        else this.cells.set(key(name, r, c), { k: 'v', v: decode(k, p) });
      }
    }
    for (const [n, t] of Object.entries(model.names)) this.names.set(n.toUpperCase(), t);
    const d = asOf ?? new Date();
    this.today = serialFromYMD(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  // ------------------------------------------------------------------ inputs

  /** Set an input cell (replaces a literal or a formula). `null` clears it. */
  set(sheet: string, a1OrRow: string | number, colOrValue: number | Scalar | null, maybeValue?: Scalar | null): void {
    let r: number, c: number, v: Scalar | null | undefined;
    if (typeof a1OrRow === 'string') {
      const ref = parseFormula(a1OrRow);
      if (ref.t !== 'ref') throw new Error(`Bad address ${a1OrRow}`);
      r = ref.r1;
      c = ref.c1;
      v = colOrValue as Scalar | null;
    } else {
      r = a1OrRow;
      c = colOrValue as number;
      v = maybeValue;
    }
    const k = key(sheet, r, c);
    this.inputs.set(k, v === null || v === undefined ? BLANK : v);
    this.sheetMax.set(sheet, Math.max(this.sheetMax.get(sheet) ?? 0, r));
    this.globalMaxRow = Math.max(this.globalMaxRow, r);
    this.invalidate();
  }

  /** Read the raw (non-evaluated) definition of a cell — for diagnostics. */
  formulaAt(sheet: string, r: number, c: number): string | null {
    const d = this.cells.get(key(sheet, r, c));
    return d && d.k === 'f' ? d.src : null;
  }

  invalidate(): void {
    this.memo.clear();
    this.rangeCache.clear();
    this.matchCache.clear();
  }

  // ------------------------------------------------------------------ evaluation

  /** Value of a cell, recalculating (with iteration for circular references) as needed. */
  get(sheet: string, a1OrRow: string | number, col?: number): Scalar {
    let r: number, c: number;
    if (typeof a1OrRow === 'string') {
      const ref = parseFormula(a1OrRow);
      if (ref.t !== 'ref') throw new Error(`Bad address ${a1OrRow}`);
      r = ref.r1;
      c = ref.c1;
    } else {
      r = a1OrRow;
      c = col!;
    }
    const k = key(sheet, r, c);
    if (this.memo.has(k)) return this.memo.get(k)!;
    this.circular = false;
    let v = this.cell(sheet, r, c);
    if (this.circular && this.model.iterate) {
      // Iterate the whole calculation until the circular values settle (Excel keeps iterating on
      // every recalculation, so a workbook in use sits at the converged fixed point).
      for (let i = 0; i < 1000; i++) {
        const snapshot = new Map(this.memo);
        this.prev = snapshot;
        this.memo.clear();
        this.rangeCache.clear();
        this.matchCache.clear();
        this.circular = false;
        v = this.cell(sheet, r, c);
        let maxChange = 0;
        for (const [kk, nv] of this.memo) {
          const ov = snapshot.get(kk);
          if (typeof nv === 'number' && typeof ov === 'number') maxChange = Math.max(maxChange, Math.abs(nv - ov));
        }
        if (maxChange < 1e-13) break;
      }
    }
    return v;
  }

  private cell(sheet: string, r: number, c: number): Scalar {
    const k = key(sheet, r, c);
    const m = this.memo.get(k);
    if (m !== undefined) return m;
    if (this.inputs.has(k)) return this.inputs.get(k)!;
    const def = this.cells.get(k);
    if (!def) return BLANK;
    if (def.k === 'v') return def.v;
    if (this.inProgress.has(k)) {
      this.circular = true;
      const p = this.prev.get(k);
      return p === undefined ? 0 : p;
    }
    this.inProgress.add(k);
    let out: Scalar;
    try {
      const ast = this.ast(def.src);
      const v = this.ev(ast, { sheet, row: r, col: c, array: def.array });
      out = this.finalScalar(v, { sheet, row: r, col: c, array: def.array });
    } catch (e) {
      if (e instanceof RangeError) throw e;
      out = E.VALUE;
      if (typeof process !== 'undefined' && process.env?.WB_DEBUG) console.error('eval error', sheet, r, c, def.src.slice(0, 120), e);
    } finally {
      this.inProgress.delete(k);
    }
    // a cell that took a stale circular value is recomputed on the next iteration
    this.memo.set(k, out);
    return out;
  }

  private ast(src: string): Node {
    let a = this.astCache.get(src);
    if (!a) {
      a = parseFormula(src);
      this.astCache.set(src, a);
    }
    return a;
  }

  private finalScalar(v: Value, ctx: Ctx): Scalar {
    const out = this.finalRaw(v, ctx);
    return out === BLANK ? 0 : out; // a formula pointing at an empty cell shows 0
  }

  private finalRaw(v: Value, ctx: Ctx): Scalar {
    if (isRef(v)) {
      if (v.r1 === v.r2 && v.c1 === v.c2) return this.cell(v.sheet, v.r1, v.c1);
      if (ctx.array) return this.cell(v.sheet, v.r1, v.c1);
      return this.intersect(v, ctx);
    }
    if (isArr(v)) return v.rows[0]?.[0] ?? BLANK;
    return v === BLANK ? 0 : v; // an empty result displays as 0 — but keep blank for "" semantics below
  }

  // ------------------------------------------------------------------ references

  private clipRows(ref: Ref): number {
    return ref.r2 === MAX_ROW ? Math.min(ref.r2, this.globalMaxRow) : ref.r2;
  }

  private materialize(ref: Ref): Arr {
    const ck = `${ref.sheet}|${ref.r1}|${ref.c1}|${ref.r2}|${ref.c2}`;
    const hit = this.rangeCache.get(ck);
    if (hit) return hit;
    const r2 = this.clipRows(ref);
    const c2 = Math.min(ref.c2, 120);
    const rows: Scalar[][] = [];
    for (let r = ref.r1; r <= r2; r++) {
      const row: Scalar[] = [];
      for (let c = ref.c1; c <= c2; c++) row.push(this.cell(ref.sheet, r, c));
      rows.push(row);
    }
    const a = arr(rows);
    this.rangeCache.set(ck, a);
    return a;
  }

  private intersect(ref: Ref, ctx: Ctx): Scalar {
    if (ref.r1 === ref.r2 && ref.c1 === ref.c2) return this.cell(ref.sheet, ref.r1, ref.c1);
    if (ref.c1 === ref.c2) {
      if (ctx.sheet === ref.sheet || true) {
        if (ctx.row >= ref.r1 && ctx.row <= ref.r2) return this.cell(ref.sheet, ctx.row, ref.c1);
      }
      return E.VALUE;
    }
    if (ref.r1 === ref.r2) {
      if (ctx.col >= ref.c1 && ctx.col <= ref.c2) return this.cell(ref.sheet, ref.r1, ctx.col);
      return E.VALUE;
    }
    return E.VALUE;
  }

  /** Value → scalar or array (refs resolved according to the evaluation mode). */
  private deref(v: Value, ctx: Ctx): Scalar | Arr {
    if (isRef(v)) {
      if (v.r1 === v.r2 && v.c1 === v.c2) return this.cell(v.sheet, v.r1, v.c1);
      return ctx.array ? this.materialize(v) : this.intersect(v, ctx);
    }
    return v;
  }

  /** Value → array (refs always materialised). */
  private toArr(v: Value): Arr {
    if (isRef(v)) return this.materialize(v);
    if (isArr(v)) return v;
    return arr([[v]]);
  }

  // ------------------------------------------------------------------ AST evaluation

  private ev(n: Node, ctx: Ctx): Value {
    switch (n.t) {
      case 'num':
        return n.v;
      case 'str':
        return n.v;
      case 'bool':
        return n.v;
      case 'err':
        return new XErr(n.v);
      case 'ref':
        return { kind: 'ref', sheet: n.sheet ?? ctx.sheet, r1: n.r1, c1: n.c1, r2: n.r2, c2: n.c2 };
      case 'name':
        return this.name(n.name, ctx);
      case 'array':
        return arr(n.rows.map((row) => row.map((x) => this.scalarOf(this.ev(x, ctx), ctx))));
      case 'un': {
        const v = this.deref(this.ev(n.e, ctx), ctx);
        return lift1(v, (x) => {
          const num = toNum(x);
          if (isErr(num)) return num;
          return n.op === '-' ? -num : n.op === '%' ? num / 100 : num;
        });
      }
      case 'bin':
        return this.binary(n, ctx);
      case 'fn':
        return this.call(n.name, n.args, ctx);
    }
  }

  private scalarOf(v: Value, ctx: Ctx): Scalar {
    const d = this.deref(v, ctx);
    return isArr(d) ? d.rows[0]?.[0] ?? BLANK : d;
  }

  private name(name: string, ctx: Ctx): Value {
    const t = this.names.get(name.toUpperCase());
    if (t === undefined) return E.NAME;
    if (t.includes(',') && !/[A-Z]+\(/i.test(t)) return E.VALUE; // union reference
    return this.ev(this.ast(t.replace(/^=/, '')), ctx);
  }

  private binary(n: Extract<Node, { t: 'bin' }>, ctx: Ctx): Value {
    if (n.op === ':') {
      const a = this.ev(n.a, ctx);
      const b = this.ev(n.b, ctx);
      if (!isRef(a) || !isRef(b)) return E.VALUE;
      return { kind: 'ref', sheet: a.sheet, r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
    }
    const a = this.deref(this.ev(n.a, ctx), ctx);
    const b = this.deref(this.ev(n.b, ctx), ctx);
    return lift2(a, b, (x, y) => binop(n.op, x, y));
  }

  // ------------------------------------------------------------------ functions

  private call(name: string, args: (Node | null)[], ctx: Ctx): Value {
    const ev = (i: number, c: Ctx = ctx): Value => (args[i] ? this.ev(args[i]!, c) : BLANK);
    const val = (i: number, c: Ctx = ctx) => this.deref(ev(i, c), c);
    const has = (i: number) => i < args.length && args[i] !== null;
    const scalarFn = (f: (...xs: Scalar[]) => Scalar): Value => liftN(args.map((_, i) => val(i)), f);

    switch (name) {
      // ---------------- logical
      case 'IF': {
        const c = val(0);
        if (isArr(c)) {
          const a = has(1) ? this.deref(ev(1), ctx) : true;
          const b = args.length > 2 ? (has(2) ? this.deref(ev(2), ctx) : 0) : false;
          return liftN([c, a, b], (cc, aa, bb) => {
            const t = toBool(cc);
            if (isErr(t)) return t;
            return t ? aa : bb;
          });
        }
        const t = toBool(c);
        if (isErr(t)) return t;
        if (t) return has(1) ? ev(1) : args.length > 1 ? 0 : true;
        return args.length > 2 ? (has(2) ? ev(2) : 0) : false;
      }
      case 'IFERROR': {
        const v = ev(0);
        const d = this.deref(v, ctx);
        if (isArr(d)) {
          if (!d.rows.some((r) => r.some(isErr))) return d;
          const alt = this.deref(ev(1), ctx);
          return lift2(d, alt, (x, y) => (isErr(x) ? y : x));
        }
        if (isErr(d)) return ev(1);
        return isRef(v) && !(v.r1 === v.r2 && v.c1 === v.c2) && ctx.array ? v : d;
      }
      case 'AND':
      case 'OR': {
        let any = false;
        let result = name === 'AND';
        for (let i = 0; i < args.length; i++) {
          const v = ev(i);
          const items: Scalar[] = isRef(v) || isArr(v) ? this.toArr(v).rows.flat().filter((x) => x !== BLANK && typeof x !== 'string') : [this.scalarOf(v, ctx)];
          for (const x of items) {
            const b = toBool(x);
            if (isErr(b)) return b;
            any = true;
            result = name === 'AND' ? result && b : result || b;
          }
        }
        return any ? result : E.VALUE;
      }
      case 'NOT':
        return scalarFn((x) => {
          const b = toBool(x);
          return isErr(b) ? b : !b;
        });
      case 'ISNUMBER':
        return scalarFn((x) => typeof x === 'number');
      case 'ISERROR':
        return scalarFn((x) => isErr(x));
      case 'ISNA':
        return scalarFn((x) => isErr(x) && x.code === '#N/A');
      case 'ISTEXT':
        return scalarFn((x) => typeof x === 'string');
      case 'ISBLANK':
        return scalarFn((x) => x === BLANK);
      case 'TRUE':
        return true;
      case 'FALSE':
        return false;

      // ---------------- aggregates
      case 'SUM':
      case 'MIN':
      case 'MAX':
      case 'COUNT': {
        const nums: number[] = [];
        for (let i = 0; i < args.length; i++) {
          if (!has(i)) continue;
          const v = ev(i);
          if (isRef(v) || isArr(v)) {
            for (const x of this.toArr(v).rows.flat()) {
              if (isErr(x)) {
                if (name === 'COUNT') continue;
                return x;
              }
              if (typeof x === 'number') nums.push(x);
            }
          } else {
            const x = v as Scalar;
            if (name === 'COUNT') {
              if (typeof x === 'number' || (typeof x === 'string' && typeof toNum(x) === 'number')) nums.push(0);
              continue;
            }
            const num = toNum(x);
            if (isErr(num)) return num;
            nums.push(num);
          }
        }
        if (name === 'COUNT') return nums.length;
        if (name === 'SUM') return nums.reduce((a, b) => a + b, 0);
        if (nums.length === 0) return 0;
        return name === 'MIN' ? Math.min(...nums) : Math.max(...nums);
      }
      case 'SUMPRODUCT': {
        const actx = { ...ctx, array: true };
        const arrays = args.map((_, i) => {
          const d = this.deref(ev(i, actx), actx);
          return isArr(d) ? d : arr([[d]]);
        });
        const { h, w } = dims(arrays[0]!);
        let total = 0;
        for (const a of arrays) {
          const d = dims(a);
          if (d.h !== h || d.w !== w) return E.VALUE;
        }
        for (let r = 0; r < h; r++) {
          for (let c = 0; c < w; c++) {
            let p = 1;
            for (const a of arrays) {
              const x = a.rows[r]![c]!;
              if (isErr(x)) return x;
              p *= typeof x === 'number' ? x : arrays.length === 1 ? 0 : typeof x === 'boolean' ? 0 : 0;
            }
            total += p;
          }
        }
        return total;
      }
      case 'SUMIF':
      case 'COUNTIF': {
        const range = this.toArr(ev(0));
        const rangeRef = ev(0);
        const sumRef = name === 'SUMIF' && has(2) ? ev(2) : rangeRef;
        const crit = val(1);
        const sumArr = this.shaped(sumRef, range);
        return lift1(crit, (cr) => {
          const pred = criteria(cr);
          let total = 0;
          range.rows.forEach((row, r) =>
            row.forEach((x, c) => {
              if (!pred(x)) return;
              if (name === 'COUNTIF') total += 1;
              else {
                const s = sumArr.rows[r]?.[c];
                if (typeof s === 'number') total += s;
              }
            }),
          );
          return total;
        });
      }
      case 'SUMIFS':
      case 'COUNTIFS': {
        const first = name === 'SUMIFS' ? 1 : 0;
        const pairs: { range: Arr; crit: Scalar | Arr }[] = [];
        for (let i = first; i + 1 < args.length + 1 && i < args.length; i += 2) pairs.push({ range: this.toArr(ev(i)), crit: val(i + 1) });
        const base = pairs[0]!.range;
        const sumArr = name === 'SUMIFS' ? this.shaped(ev(0), base) : null;
        return liftN(pairs.map((p) => p.crit), (...crs) => {
          const preds = crs.map((c) => criteria(c));
          let total = 0;
          base.rows.forEach((row, r) =>
            row.forEach((_, c) => {
              for (let k = 0; k < pairs.length; k++) if (!preds[k]!(pairs[k]!.range.rows[r]?.[c] ?? BLANK)) return;
              if (!sumArr) total += 1;
              else {
                const s = sumArr.rows[r]?.[c];
                if (typeof s === 'number') total += s;
              }
            }),
          );
          return total;
        });
      }
      case 'LARGE':
      case 'SMALL': {
        const a = this.toArr(this.deref(ev(0), ctx));
        const nums = a.rows.flat().filter((x): x is number => typeof x === 'number');
        const errv = a.rows.flat().find(isErr);
        if (errv) return errv;
        return lift1(val(1), (kk) => {
          const k = toNum(kk);
          if (isErr(k)) return k;
          const sorted = [...nums].sort((x, y) => (name === 'LARGE' ? y - x : x - y));
          const v = sorted[Math.ceil(k) - 1];
          return v === undefined ? E.NUM : v;
        });
      }
      case 'AGGREGATE': {
        // only AGGREGATE(3, 5, ref) — COUNTA ignoring hidden rows — is used
        const a = this.toArr(ev(2));
        return a.rows.flat().filter((x) => x !== BLANK).length;
      }

      // ---------------- lookup & reference
      case 'INDEX': {
        const src = ev(0);
        const rowV = has(1) ? val(1) : 0;
        const colV = has(2) ? val(2) : null;
        if (isArr(rowV) || (colV !== null && isArr(colV))) {
          return liftN([rowV, colV ?? BLANK], (rr, cc) => this.scalarOf(this.index(src, rr, colV === null ? null : cc, ctx), ctx));
        }
        return this.index(src, rowV as Scalar, colV as Scalar | null, ctx);
      }
      case 'MATCH': {
        const look = this.toArr(ev(1));
        const type = has(2) ? toNum(this.scalarOf(ev(2), ctx)) : 1;
        if (isErr(type)) return type;
        const flat = vector(look);
        if (!flat) return E.NA;
        const lookRef = ev(1);
        return lift1(val(0), (v) => this.match(v, flat, type, isRef(lookRef) ? lookRef : null));
      }
      case 'VLOOKUP': {
        const table = this.toArr(ev(1));
        const colN = toNum(this.scalarOf(ev(2), ctx));
        if (isErr(colN)) return colN;
        const approx = has(3) ? toBool(this.scalarOf(ev(3), ctx)) : true;
        if (isErr(approx)) return approx;
        const firstCol = table.rows.map((r) => r[0] ?? BLANK);
        const tableRef = ev(1);
        return lift1(val(0), (v) => {
          if (colN < 1 || colN > (table.rows[0]?.length ?? 0)) return E.REF;
          const i = this.match(v, firstCol, approx ? 1 : 0, isRef(tableRef) ? { ...tableRef, c2: tableRef.c1 } : null);
          if (isErr(i)) return i;
          return table.rows[(i as number) - 1]![colN - 1] ?? BLANK;
        });
      }
      case 'LOOKUP': {
        const vec = vector(this.toArr(ev(1)));
        const res = has(2) ? vector(this.toArr(ev(2))) : vec;
        if (!vec || !res) return E.NA;
        return lift1(val(0), (v) => {
          const i = this.match(v, vec, 1, null);
          if (isErr(i)) return i;
          return res[(i as number) - 1] ?? E.NA;
        });
      }
      case 'OFFSET': {
        const base = ev(0);
        if (!isRef(base)) return E.VALUE;
        const ro = toNum(this.scalarOf(ev(1), ctx));
        const co = toNum(this.scalarOf(ev(2), ctx));
        if (isErr(ro)) return ro;
        if (isErr(co)) return co;
        const h = has(3) ? toNum(this.scalarOf(ev(3), ctx)) : base.r2 - base.r1 + 1;
        const w = has(4) ? toNum(this.scalarOf(ev(4), ctx)) : base.c2 - base.c1 + 1;
        if (isErr(h)) return h;
        if (isErr(w)) return w;
        if (h < 1 || w < 1) return E.REF;
        const r1 = base.r1 + Math.trunc(ro);
        const c1 = base.c1 + Math.trunc(co);
        if (r1 < 1 || c1 < 1) return E.REF;
        return { kind: 'ref', sheet: base.sheet, r1, c1, r2: r1 + Math.trunc(h) - 1, c2: c1 + Math.trunc(w) - 1 };
      }
      case 'ROW': {
        if (!has(0)) return ctx.row;
        const v = ev(0);
        if (!isRef(v)) return E.VALUE;
        if (v.r1 === v.r2 || !ctx.array) return v.r1;
        const rows: Scalar[][] = [];
        for (let r = v.r1; r <= this.clipRows(v); r++) rows.push([r]);
        return arr(rows);
      }
      case 'ROWS': {
        const v = ev(0);
        if (isRef(v)) return v.r2 - v.r1 + 1;
        if (isArr(v)) return v.rows.length;
        return 1;
      }
      case 'FILTER': {
        if (this.libreOfficeCompat) return E.NAME;
        const actx = { ...ctx, array: true };
        const a = this.toArr(this.deref(ev(0, actx), actx));
        const inc = this.toArr(this.deref(ev(1, actx), actx));
        const out: Scalar[][] = [];
        a.rows.forEach((row, r) => {
          const f = inc.rows[r]?.[0] ?? inc.rows[0]?.[r];
          if (isErr(f)) {
            out.push([f]);
            return;
          }
          const b = toBool(f ?? false);
          if (b === true) out.push(row);
        });
        if (out.length === 0) return has(2) ? val(2) : new XErr('#CALC!');
        return arr(out);
      }

      // ---------------- text
      case 'TEXTJOIN': {
        const delim = toText(this.scalarOf(ev(0), ctx));
        const ignore = toBool(this.scalarOf(ev(1), ctx));
        if (isErr(delim)) return delim;
        if (isErr(ignore)) return ignore;
        const parts: string[] = [];
        for (let i = 2; i < args.length; i++) {
          const v = ev(i);
          const items = isRef(v) || isArr(v) ? this.toArr(v).rows.flat() : [v as Scalar];
          for (const x of items) {
            if (isErr(x)) return x;
            const t = toText(x) as string;
            if (ignore && t === '') continue;
            parts.push(t);
          }
        }
        return parts.join(delim);
      }
      case 'LEFT':
      case 'RIGHT':
        return scalarFn((s, nn = 1) => {
          const t = toText(s);
          if (isErr(t)) return t;
          const k = nn === BLANK ? 1 : toNum(nn);
          if (isErr(k)) return k;
          if (k < 0) return E.VALUE;
          return name === 'LEFT' ? t.slice(0, Math.floor(k)) : Math.floor(k) === 0 ? '' : t.slice(-Math.floor(k));
        });
      case 'MID':
        return scalarFn((s, st, nn) => {
          const t = toText(s);
          const a = toNum(st);
          const b = toNum(nn);
          if (isErr(t)) return t;
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          if (a < 1 || b < 0) return E.VALUE;
          return t.substr(Math.floor(a) - 1, Math.floor(b));
        });
      case 'LEN':
        return scalarFn((s) => {
          const t = toText(s);
          return isErr(t) ? t : t.length;
        });
      case 'TRIM':
        return scalarFn((s) => {
          const t = toText(s);
          return isErr(t) ? t : t.replace(/ +/g, ' ').trim();
        });
      case 'UPPER':
      case 'LOWER':
        return scalarFn((s) => {
          const t = toText(s);
          return isErr(t) ? t : name === 'UPPER' ? t.toUpperCase() : t.toLowerCase();
        });
      case 'SUBSTITUTE':
        return scalarFn((s, from, to) => {
          const t = toText(s), f = toText(from), r = toText(to);
          if (isErr(t)) return t;
          if (isErr(f)) return f;
          if (isErr(r)) return r;
          return f === '' ? t : t.split(f).join(r);
        });
      case 'SEARCH':
      case 'FIND':
        return scalarFn((find, within, start = 1) => {
          const f = toText(find), w = toText(within);
          const s = start === BLANK ? 1 : toNum(start);
          if (isErr(f)) return f;
          if (isErr(w)) return w;
          if (isErr(s)) return s;
          if (s < 1 || s > w.length + 1) return E.VALUE;
          if (name === 'FIND') {
            const i = w.indexOf(f, s - 1);
            return i < 0 ? E.VALUE : i + 1;
          }
          const re = new RegExp(wildcardRegex(f).source.replace(/^\^/, '').replace(/\$$/, ''), 'is');
          const m = re.exec(w.slice(s - 1));
          return m ? m.index + s : E.VALUE;
        });
      case 'TEXT':
        return scalarFn((v, fmt) => {
          const f = toText(fmt);
          if (isErr(f)) return f;
          return formatText(v, f);
        });
      case 'VALUE':
        return scalarFn((v) => (typeof v === 'number' ? v : toNum(v)));
      case 'HYPERLINK':
        return has(1) ? val(1) : val(0);

      // ---------------- maths
      case 'ABS':
        return scalarFn((x) => num1(x, Math.abs));
      case 'LOG10':
        return scalarFn((x) => num1(x, (v) => (v <= 0 ? E.NUM : Math.log10(v))));
      case 'MOD':
        return scalarFn((a, b) => {
          const x = toNum(a), y = toNum(b);
          if (isErr(x)) return x;
          if (isErr(y)) return y;
          if (y === 0) return E.DIV0;
          return x - y * Math.floor(x / y);
        });
      case 'ROUND':
        return scalarFn((a, b) => {
          const x = toNum(a), d = toNum(b);
          if (isErr(x)) return x;
          if (isErr(d)) return d;
          return roundHalfAway(x, Math.trunc(d));
        });
      case 'INT':
        return scalarFn((x) => num1(x, Math.floor));

      // ---------------- dates
      case 'TODAY':
        return this.today;
      case 'DATEVALUE':
        return scalarFn((t) => {
          if (typeof t !== 'string') return E.VALUE;
          const d = parseDateText(t);
          return d === null ? E.VALUE : d;
        });
      case 'YEAR':
        return scalarFn((x) => num1(x, (v) => ymdFromSerial(v).y));
      case 'EDATE':
        return scalarFn((s, m) => {
          const a = toNum(s), b = toNum(m);
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          const { y, m: mo, d } = ymdFromSerial(a);
          const total = y * 12 + (mo - 1) + Math.trunc(b);
          const ny = Math.floor(total / 12), nm = (total % 12) + 1;
          const last = ymdFromSerial(serialFromYMD(ny, nm + 1, 1) - 1).d;
          return serialFromYMD(ny, nm, Math.min(d, last));
        });

      // ---------------- not needed by the engine
      case 'CELL': {
        const info = toText(this.scalarOf(ev(0), ctx));
        const target = ev(1);
        if (info !== 'address' || !isRef(target)) return E.VALUE;
        const book = `[${this.model.source}]${target.sheet}`;
        const sheetPart = /[^A-Za-z0-9_]/.test(target.sheet) ? `'${book}'` : book;
        return `${sheetPart}!$${indexToCol(target.c1)}$${target.r1}`;
      }
      case 'GETHYPERLINKADDRESS':
        return E.NAME;
      default:
        return E.NAME;
    }
  }

  /** Resize `ref` to the shape of `like` anchored at its top-left (SUMIF sum_range behaviour). */
  private shaped(v: Value, like: Arr): Arr {
    if (!isRef(v)) return this.toArr(v);
    const { h, w } = dims(like);
    return this.materialize({ ...v, r2: v.r1 + h - 1, c2: v.c1 + w - 1 });
  }

  private index(src: Value, rowV: Scalar, colV: Scalar | null, ctx: Ctx): Value {
    let r = toNum(rowV);
    let c = colV === null ? null : toNum(colV);
    if (isErr(r)) return r;
    if (c !== null && isErr(c)) return c;
    if (isRef(src)) {
      const h = this.clipRows(src) - src.r1 + 1;
      const w = src.c2 - src.c1 + 1;
      if (c === null) {
        if (h === 1 && w > 1) {
          c = r;
          r = 1;
        } else c = w === 1 ? 1 : 0;
      }
      r = Math.trunc(r);
      c = Math.trunc(c as number);
      if (r < 0 || c < 0 || r > h || c > w) return E.REF;
      const r1 = r === 0 ? src.r1 : src.r1 + r - 1;
      const r2 = r === 0 ? src.r2 : r1;
      const c1 = c === 0 ? src.c1 : src.c1 + c - 1;
      const c2 = c === 0 ? src.c2 : c1;
      return { kind: 'ref', sheet: src.sheet, r1, c1, r2, c2 };
    }
    const a = this.toArr(this.deref(src, ctx));
    const { h, w } = dims(a);
    if (c === null) {
      if (h === 1) {
        c = r;
        r = 1;
      } else c = 1;
    }
    r = Math.trunc(r);
    c = Math.trunc(c as number);
    if (r < 0 || c < 0 || r > h || c > w) return E.REF;
    if (r === 0) return arr(a.rows.map((row) => [row[c! - 1]!]));
    if (c === 0) return arr([a.rows[r - 1]!]);
    return a.rows[r - 1]![c - 1]!;
  }

  private match(v: Scalar, flat: Scalar[], type: number, ref: Ref | null): number | XErr {
    if (isErr(v)) return v;
    if (type === 0) {
      if (typeof v === 'string' && /[*?~]/.test(v)) {
        const re = wildcardRegex(v);
        const i = flat.findIndex((x) => typeof x === 'string' && re.test(x));
        return i < 0 ? E.NA : i + 1;
      }
      // cached exact-match index for reference lookups
      if (ref) {
        const ck = `${ref.sheet}|${ref.r1}|${ref.c1}|${ref.r2}|${ref.c2}`;
        let idx = this.matchCache.get(ck);
        if (!idx) {
          idx = new Map();
          flat.forEach((x, i) => {
            const k = matchKey(x);
            if (k !== null && !idx!.has(k)) idx!.set(k, i + 1);
          });
          this.matchCache.set(ck, idx);
        }
        const k = matchKey(v);
        if (k === null) return E.NA;
        return idx.get(k) ?? E.NA;
      }
      const i = flat.findIndex((x) => x !== BLANK && compare(x, v) === 0 && typeof x === typeof v);
      return i < 0 ? E.NA : i + 1;
    }
    // approximate: binary-search semantics on (assumed) sorted data
    let lo = 0, hi = flat.length - 1, ans = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      let m = mid;
      while (m <= hi && (flat[m] === BLANK || typeof flat[m] !== typeof v)) m++;
      if (m > hi) {
        hi = mid - 1;
        continue;
      }
      const cmp = compare(flat[m]!, v);
      if (isErr(cmp)) return cmp;
      if (type > 0 ? cmp <= 0 : cmp >= 0) {
        ans = m;
        lo = m + 1;
      } else hi = mid - 1;
    }
    return ans < 0 ? E.NA : ans + 1;
  }
}

function matchKey(x: Scalar): string | null {
  if (x === BLANK) return null;
  if (typeof x === 'number') return `n:${x}`;
  if (typeof x === 'string') return `s:${x.toLowerCase()}`;
  if (typeof x === 'boolean') return `b:${x}`;
  return null;
}

function vector(a: Arr): Scalar[] | null {
  const { h, w } = dims(a);
  if (h === 1) return a.rows[0]!;
  if (w === 1) return a.rows.map((r) => r[0]!);
  return null;
}

function num1(x: Scalar, f: (n: number) => number | XErr): Scalar {
  const n = toNum(x);
  return isErr(n) ? n : f(n);
}

export function roundHalfAway(v: number, digits: number): number {
  const scaled = Number((Math.abs(v) * Math.pow(10, digits)).toPrecision(15));
  const r = Math.floor(scaled + 0.5);
  const out = digits >= 0 ? r / Math.pow(10, digits) : r * Math.pow(10, -digits);
  return v < 0 ? -out : out;
}

function binop(op: string, x: Scalar, y: Scalar): Scalar {
  switch (op) {
    case '+':
    case '-':
    case '*':
    case '/':
    case '^': {
      const a = toNum(x);
      if (isErr(a)) return a;
      const b = toNum(y);
      if (isErr(b)) return b;
      if (op === '+') return a + b;
      if (op === '-') return a - b;
      if (op === '*') return a * b;
      if (op === '/') return b === 0 ? E.DIV0 : a / b;
      return Math.pow(a, b);
    }
    case '&': {
      const a = toText(x);
      if (isErr(a)) return a;
      const b = toText(y);
      if (isErr(b)) return b;
      return a + b;
    }
    default: {
      const c = compare(x, y);
      if (isErr(c)) return c;
      switch (op) {
        case '=':
          return c === 0;
        case '<>':
          return c !== 0;
        case '<':
          return c < 0;
        case '>':
          return c > 0;
        case '<=':
          return c <= 0;
        case '>=':
          return c >= 0;
      }
      return E.VALUE;
    }
  }
}

function lift1(v: Scalar | Arr | Value, f: (x: Scalar) => Scalar): Scalar | Arr {
  if (isArr(v as Value)) return arr((v as Arr).rows.map((r) => r.map(f)));
  return f(v as Scalar);
}

function lift2(a: Scalar | Arr, b: Scalar | Arr, f: (x: Scalar, y: Scalar) => Scalar): Scalar | Arr {
  return liftN([a, b], f);
}

/** Element-wise application with Excel broadcasting (1-row/1-column arrays expand; out of range → #N/A). */
function liftN(vals: (Scalar | Arr | Value)[], f: (...xs: Scalar[]) => Scalar): Scalar | Arr {
  const arrays = vals.filter((v) => isArr(v as Value)) as Arr[];
  if (arrays.length === 0) return f(...(vals as Scalar[]));
  const h = Math.max(...arrays.map((a) => dims(a).h));
  const w = Math.max(...arrays.map((a) => dims(a).w));
  const rows: Scalar[][] = [];
  for (let r = 0; r < h; r++) {
    const row: Scalar[] = [];
    for (let c = 0; c < w; c++) {
      const xs = vals.map((v) => {
        if (!isArr(v as Value)) return v as Scalar;
        const a = v as Arr;
        const d = dims(a);
        const rr = d.h === 1 ? 0 : r;
        const cc = d.w === 1 ? 0 : c;
        if (rr >= d.h || cc >= d.w) return E.NA;
        return a.rows[rr]![cc]!;
      });
      row.push(f(...xs));
    }
    rows.push(row);
  }
  return arr(rows);
}

/** TEXT() for the formats the workbook uses. */
export function formatText(v: Scalar, fmt: string): Scalar {
  if (isErr(v)) return v;
  const n = typeof v === 'number' ? v : toNum(v);
  if (isErr(n)) return typeof v === 'string' ? v : n;
  const f = fmt.trim();
  const pct = f.endsWith('%');
  const core = pct ? f.slice(0, -1) : f;
  const x = pct ? n * 100 : n;
  const m = /^(\$?)(#,##)?0(?:\.(0+))?$/.exec(core);
  if (m) {
    const decimals = m[3]?.length ?? 0;
    let s = roundHalfAway(Math.abs(x), decimals).toFixed(decimals);
    if (m[2]) {
      const [i, d] = s.split('.');
      s = i!.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d ? `.${d}` : '');
    }
    return `${x < 0 && Number(s.replace(/,/g, '')) !== 0 ? '-' : ''}${m[1]}${s}${pct ? '%' : ''}`;
  }
  if (/^0$/.test(f)) return String(roundHalfAway(n, 0));
  const date = /^(d{1,2})\/(m{1,2})\/(y{2,4})$|^(y{4})-(m{2})-(d{2})$/i.exec(f);
  if (date) {
    const { y, m: mo, d } = ymdFromSerial(n);
    const p2 = (k: number) => String(k).padStart(2, '0');
    return date[4] ? `${y}-${p2(mo)}-${p2(d)}` : `${date[1]!.length === 2 ? p2(d) : d}/${date[2]!.length === 2 ? p2(mo) : mo}/${date[3]!.length === 2 ? String(y).slice(-2) : y}`;
  }
  return toText(v);
}
