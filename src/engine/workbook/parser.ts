/**
 * Excel formula parser (the subset used by the recommendations workbook).
 * Produces an AST; references are kept in A1 form with optional sheet.
 */

export type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'err'; v: string }
  | { t: 'ref'; sheet: string | null; r1: number; c1: number; r2: number; c2: number }
  | { t: 'name'; name: string; sheet: string | null }
  | { t: 'array'; rows: Node[][] }
  | { t: 'fn'; name: string; args: (Node | null)[] }
  | { t: 'un'; op: '-' | '+' | '%'; e: Node }
  | { t: 'bin'; op: string; a: Node; b: Node };

export const MAX_ROW = 1048576;
export const MAX_COL = 16384;

export function colToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function indexToCol(n: number): string {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

type Tok =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'err'; v: string }
  | { k: 'ref'; sheet: string | null; text: string }
  | { k: 'id'; v: string; sheet: string | null }
  | { k: 'op'; v: string }
  | { k: 'eof' };

const ERRORS = ['#N/A', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#NULL!', '#GETTING_DATA'];

function lex(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  const peekRef = (s: string, at: number): RegExpExecArray | null =>
    /^(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)(?![A-Za-z0-9_.(])/.exec(s.slice(at));
  while (i < src.length) {
    const ch = src[i]!;
    if (ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t') {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j];
        j++;
      }
      out.push({ k: 'str', v: s });
      i = j + 1;
      continue;
    }
    if (ch === '#') {
      const e = ERRORS.find((x) => src.startsWith(x, i));
      if (!e) throw new Error(`Unknown error literal at ${i}: ${src.slice(i, i + 10)}`);
      out.push({ k: 'err', v: e });
      i += e.length;
      continue;
    }
    // sheet-qualified: 'Sheet Name'!ref or Sheet!ref
    let sheet: string | null = null;
    if (ch === "'") {
      const j = src.indexOf("'!", i + 1);
      if (j < 0) throw new Error('Unterminated sheet name');
      sheet = src.slice(i + 1, j).replace(/''/g, "'");
      i = j + 2;
    } else {
      const m = /^([A-Za-z_][A-Za-z0-9_.]*)!/.exec(src.slice(i));
      if (m) {
        sheet = m[1]!;
        i += m[0].length;
      }
    }
    const c = src[i];
    if (c === undefined) break;
    const rm = peekRef(src, i);
    if (rm && (sheet !== null || !/^[A-Za-z_]/.test(src.slice(i)) || /^\$?[A-Za-z]{1,3}\$?\d+/.test(src.slice(i)) || /^\$?[A-Za-z]{1,3}:/.test(src.slice(i)))) {
      // a cell/column/row reference — but not a function call like LOG10( (excluded by lookahead)
      out.push({ k: 'ref', sheet, text: rm[1]! });
      i += rm[1]!.length;
      continue;
    }
    if (/[0-9.]/.test(c) && sheet === null) {
      const m = /^(\d+\.?\d*(?:[eE][+-]?\d+)?|\.\d+(?:[eE][+-]?\d+)?)/.exec(src.slice(i))!;
      out.push({ k: 'num', v: Number(m[1]) });
      i += m[1]!.length;
      continue;
    }
    if (/[A-Za-z_\\]/.test(c)) {
      const m = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/.exec(src.slice(i))!;
      out.push({ k: 'id', v: m[0], sheet });
      i += m[0].length;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === '<>' || two === '<=' || two === '>=') {
      out.push({ k: 'op', v: two });
      i += 2;
      continue;
    }
    if ('+-*/^&=<>%(),:{};'.includes(c)) {
      out.push({ k: 'op', v: c });
      i++;
      continue;
    }
    throw new Error(`Unexpected character "${c}" at ${i} in ${src.slice(0, 80)}`);
  }
  out.push({ k: 'eof' });
  return out;
}

function parseRefText(text: string, sheet: string | null): Node {
  const strip = (s: string) => s.replace(/\$/g, '');
  const parts = text.split(':').map(strip);
  const cell = /^([A-Za-z]{1,3})(\d+)$/;
  if (parts.length === 1) {
    const m = cell.exec(parts[0]!)!;
    const r = Number(m[2]);
    const c = colToIndex(m[1]!);
    return { t: 'ref', sheet, r1: r, c1: c, r2: r, c2: c };
  }
  const [a, b] = parts as [string, string];
  const ma = cell.exec(a);
  const mb = cell.exec(b);
  if (ma && mb) {
    const r1 = Number(ma[2]), r2 = Number(mb[2]), c1 = colToIndex(ma[1]!), c2 = colToIndex(mb[1]!);
    return { t: 'ref', sheet, r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
  }
  if (/^[A-Za-z]+$/.test(a) && /^[A-Za-z]+$/.test(b)) {
    const c1 = colToIndex(a), c2 = colToIndex(b);
    return { t: 'ref', sheet, r1: 1, c1: Math.min(c1, c2), r2: MAX_ROW, c2: Math.max(c1, c2) };
  }
  if (/^\d+$/.test(a) && /^\d+$/.test(b)) {
    const r1 = Number(a), r2 = Number(b);
    return { t: 'ref', sheet, r1: Math.min(r1, r2), c1: 1, r2: Math.max(r1, r2), c2: MAX_COL };
  }
  throw new Error(`Bad reference ${text}`);
}

export function parseFormula(src: string): Node {
  const toks = lex(src);
  let p = 0;
  const peek = () => toks[p]!;
  const next = () => toks[p++]!;
  const isOp = (v: string) => {
    const t = peek();
    return t.k === 'op' && t.v === v;
  };
  const expect = (v: string) => {
    const t = next();
    if (t.k !== 'op' || t.v !== v) throw new Error(`Expected ${v} in ${src.slice(0, 120)}`);
  };

  // precedence: comparison < & < +- < */ < ^ < unary/percent < range ':'
  function comparison(): Node {
    let a = concat();
    while (peek().k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes((peek() as { v: string }).v)) {
      const op = (next() as { v: string }).v;
      a = { t: 'bin', op, a, b: concat() };
    }
    return a;
  }
  function concat(): Node {
    let a = additive();
    while (isOp('&')) {
      next();
      a = { t: 'bin', op: '&', a, b: additive() };
    }
    return a;
  }
  function additive(): Node {
    let a = multiplicative();
    while (isOp('+') || isOp('-')) {
      const op = (next() as { v: string }).v;
      a = { t: 'bin', op, a, b: multiplicative() };
    }
    return a;
  }
  function multiplicative(): Node {
    let a = power();
    while (isOp('*') || isOp('/')) {
      const op = (next() as { v: string }).v;
      a = { t: 'bin', op, a, b: power() };
    }
    return a;
  }
  function power(): Node {
    let a = unary();
    while (isOp('^')) {
      next();
      a = { t: 'bin', op: '^', a, b: unary() };
    }
    return a;
  }
  function unary(): Node {
    if (isOp('-') || isOp('+')) {
      const op = (next() as { v: string }).v as '-' | '+';
      return { t: 'un', op, e: unary() };
    }
    let e = range();
    while (isOp('%')) {
      next();
      e = { t: 'un', op: '%', e };
    }
    return e;
  }
  function range(): Node {
    let a = primary();
    while (isOp(':')) {
      next();
      a = { t: 'bin', op: ':', a, b: primary() };
    }
    return a;
  }
  function primary(): Node {
    const t = next();
    switch (t.k) {
      case 'num':
        return { t: 'num', v: t.v };
      case 'str':
        return { t: 'str', v: t.v };
      case 'err':
        return { t: 'err', v: t.v };
      case 'ref':
        return parseRefText(t.text, t.sheet);
      case 'id': {
        const up = t.v.toUpperCase();
        if (isOp('(')) {
          next();
          const args: (Node | null)[] = [];
          if (isOp(')')) {
            next();
          } else {
            for (;;) {
              if (isOp(',') || isOp(')')) args.push(null);
              else args.push(comparison());
              if (isOp(',')) {
                next();
                continue;
              }
              expect(')');
              break;
            }
          }
          const name = up.replace(/^_XLFN\._XLWS\./, '').replace(/^_XLFN\./, '').replace(/^_XLWS\./, '');
          return { t: 'fn', name, args };
        }
        if (up === 'TRUE' || up === 'FALSE') return { t: 'bool', v: up === 'TRUE' };
        return { t: 'name', name: t.v, sheet: t.sheet };
      }
      case 'op':
        if (t.v === '(') {
          const e = comparison();
          expect(')');
          return e;
        }
        if (t.v === '{') {
          const rows: Node[][] = [[]];
          for (;;) {
            rows[rows.length - 1]!.push(unary());
            if (isOp(',')) {
              next();
              continue;
            }
            if (isOp(';')) {
              next();
              rows.push([]);
              continue;
            }
            expect('}');
            break;
          }
          return { t: 'array', rows };
        }
        break;
      default:
        break;
    }
    throw new Error(`Unexpected token ${JSON.stringify(t)} in ${src.slice(0, 160)}`);
  }
  const e = comparison();
  if (peek().k !== 'eof') throw new Error(`Trailing input in ${src.slice(0, 160)}: ${JSON.stringify(peek())}`);
  return e;
}
