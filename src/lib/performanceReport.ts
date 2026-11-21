/**
 * Custodian performance reports for the Annual Review: ports of the parsing in the workbook's
 * AR_Part1 → Apex_Tables / OA_Tables macros.
 *
 * The adviser pastes the report text into the workbook's "Input" sheet (one line per row in column A);
 * here it is pasted into a text box. Excel turns lines that look like numbers ("$1,234.56", "12.34%",
 * "(1,234.56)") into numbers when pasted, and the macros rely on that, so `excelValue` does the same.
 *
 * Deliberate differences from the macros (all stop wrong figures appearing silently):
 *  - Apex sector returns: when a heading or its data line is missing, the row is left blank; the macro
 *    reused the previous heading's figures.
 *  - Apex current allocation: the percentage is the last "%" figure on the line (the macro took the last
 *    word, which is the dollar amount when the line also has one).
 *  - OneAnswer withdrawals are always shown, and counted, as a negative amount whatever sign the report uses.
 */

export type ReportCustodian = 'Apex' | 'OA';

export interface Table1Row {
  text: string;
  negative: boolean;
}

export interface SectorReturn {
  heading: string;
  /** Text for column 2 (the return %), or null to leave the template's "%". */
  percent: string | null;
  /** Text for column 3 (the dollar return), or null to leave the template's "$". */
  dollars: string | null;
}

export type AllocationClass = 'Cash' | 'Fixed Interest' | 'Australasian Equities' | 'Global Equities' | 'Property' | 'Multi Sector';
export const ALLOCATION_CLASSES: AllocationClass[] = ['Cash', 'Fixed Interest', 'Australasian Equities', 'Global Equities', 'Property', 'Multi Sector'];

export interface PerformanceReport {
  custodian: ReportCustodian;
  /** "1 July 2024" style dates from the reporting period. */
  openingDate: string | null;
  closingDate: string | null;
  /** Since-inception returns (the last two rows of Table1); null when the report has no inception section. */
  inception: { start: Date; end: Date; days: number; arithmetic: number; geometric: number } | null;
  /** Table1 values, top to bottom from the bookmarked cell. */
  table1: Table1Row[];
  dollarGain: string | null;
  percentGain: string | null;
  /** Table2: return per sector. */
  sectors: SectorReturn[];
  /** Table3: current allocation per class as a fraction; null when not in the report. */
  current: Record<AllocationClass, number | null>;
  /** Labels the parser looked for and did not find. */
  missing: string[];
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const MON = MONTHS.map((m) => m.slice(0, 3).toLowerCase());

/** The value Excel stores when this text is pasted into a cell. */
export function excelValue(text: string): number | string {
  const s = text.trim();
  if (s === '') return '';
  const m = /^(\()?(-)?(\$)?(-)?(\d{1,3}(?:,\d{3})+|\d+)?(\.\d+)?(%)?(\))?$/.exec(s.replace(/\s/g, ''));
  if (!m || (!m[5] && !m[6]) || !!m[1] !== !!m[8]) return s;
  let n = Number(`${(m[5] ?? '0').replace(/,/g, '')}${m[6] ?? ''}`);
  if (m[1] || m[2] || m[4]) n = -n;
  if (m[7]) n /= 100;
  return n;
}

/** VBA Val(): the leading number in a string ("1234.5abc" → 1234.5, "$5" → 0). */
function vbaVal(s: string): number {
  const m = /^\s*([-+]?(\d+\.?\d*|\.\d+))/.exec(s);
  return m ? Number(m[1]) : 0;
}

/** VBA IsNumeric for a word: numbers with optional sign, thousands separators and a currency sign. */
function isNumericWord(w: string): boolean {
  return /^[-+]?\$?[-+]?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?$/.test(w) && /\d/.test(w);
}

/** Text dates in reports: "1 Jul 2024", "01/07/2024", "1 July 2024". */
export function parseReportDate(s: string): Date | null {
  const t = s.trim();
  let m = /^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/.exec(t);
  if (m) {
    const mi = MON.indexOf(m[2]!.slice(0, 3).toLowerCase());
    if (mi >= 0) return new Date(Date.UTC(+m[3]!, mi, +m[1]!));
  }
  m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/.exec(t);
  if (m) return new Date(Date.UTC(+m[3]!, +m[2]! - 1, +m[1]!));
  return null;
}

/** ConvertAndFormatDate: "1 Jul 2024" → "1 July 2024"; other formats are returned unchanged. */
export function longReportDate(s: string): string {
  const parts = s.trim().split(' ');
  if (parts.length === 3) {
    const mi = MON.indexOf(parts[1]!.slice(0, 3).toLowerCase());
    if (mi >= 0) return `${parts[0]} ${MONTHS[mi]} ${parts[2]}`;
  }
  return s.trim();
}

const SHORT_MON = MONTHS.map((m) => m.slice(0, 3));
/** Format(d, "dd MMM yy") */
export const ddMmmYy = (d: Date) => `${String(d.getUTCDate()).padStart(2, '0')} ${SHORT_MON[d.getUTCMonth()]} ${String(d.getUTCFullYear()).slice(-2)}`;

const group = (n: number, dp: number) => {
  const [i, d] = Math.abs(n).toFixed(dp).split('.');
  return i!.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d ? `.${d}` : '');
};
/** Format(x, "$#,##0.00") */
export const dollars2 = (n: number) => `${n < 0 && Math.abs(n) >= 0.005 ? '-' : ''}$${group(n, 2)}`;
/** Format(x, "$#,##0") */
export const dollars0 = (n: number) => `${n < 0 && Math.abs(n) >= 0.5 ? '-' : ''}$${group(n, 0)}`;
/** Format(x, "0.00%") */
export const pct2 = (n: number) => `${(n * 100).toFixed(2)}%`;

/** Report text → column A of the Input sheet. Tab-separated pastes put later fields in other columns. */
export function reportLines(text: string): string[] {
  return text.split(/\r?\n/).map((l) => (l.split('\t')[0] ?? '').trim());
}

/** Range.Find(LookAt:=xlWhole): first cell equal to `label`, case-insensitive unless asked. */
function findLine(lines: string[], label: string, matchCase = false, occurrence = 1): number {
  let n = 0;
  for (let i = 0; i < lines.length; i++) {
    const a = lines[i]!;
    if (matchCase ? a === label : a.toLowerCase() === label.toLowerCase()) {
      n++;
      if (n === occurrence) return i;
    }
  }
  return -1;
}

const dateRange = (line: string | undefined): [string, string] | null => {
  if (!line) return null;
  const parts = line.split(' to ');
  return parts.length === 2 ? [parts[0]!.trim(), parts[1]!.trim()] : null;
};

function inceptionFrom(range: [string, string] | null, pct: number | null): PerformanceReport['inception'] {
  if (!range || pct === null) return null;
  const start = parseReportDate(range[0]);
  const end = parseReportDate(range[1]);
  if (!start || !end) return null;
  const days = Math.round((end.getTime() - start.getTime()) / 86400000);
  if (days <= 0) return null;
  return { start, end, days, arithmetic: (pct / days) * 365, geometric: (1 + pct) ** (365 / days) - 1 };
}

const emptyCurrent = (): Record<AllocationClass, number | null> => ({ Cash: null, 'Fixed Interest': null, 'Australasian Equities': null, 'Global Equities': null, Property: null, 'Multi Sector': null });

// ---------------------------------------------------------------- Apex

export function parseApexReport(text: string): PerformanceReport {
  const lines = reportLines(text);
  const missing: string[] = [];
  const num = (i: number): number | null => {
    const v = excelValue(lines[i] ?? '');
    return typeof v === 'number' ? v : null;
  };
  const out: PerformanceReport = { custodian: 'Apex', openingDate: null, closingDate: null, inception: null, table1: [], dollarGain: null, percentGain: null, sectors: [], current: emptyCurrent(), missing };

  const inc = findLine(lines, 'Net Return Since Inception');
  if (inc < 0) missing.push('Net Return Since Inception');
  else out.inception = inceptionFrom(dateRange(lines[inc + 1]), num(inc + 9));

  const npr = findLine(lines, 'Net Portfolio Return');
  if (npr < 0) missing.push('Net Portfolio Return');
  else {
    for (let i = 1; i <= 8; i++) {
      const v = num(npr + i);
      if (v === null) {
        out.table1.push({ text: (lines[npr + i] ?? '').trim(), negative: false });
        continue;
      }
      let t = i <= 7 ? dollars2(v) : `${(v * 100).toFixed(2)}%`;
      if (i === 3 && v < 0) t = `(${dollars2(Math.abs(v))})`;
      out.table1.push({ text: t, negative: v < 0 });
      if (i === 7) out.dollarGain = t;
      if (i === 8) out.percentGain = t;
    }
  }

  const per = findLine(lines, 'Net Return for Period');
  const range = per >= 0 ? dateRange(lines[per + 1]) : null;
  if (!range) missing.push('Net Return for Period');
  else {
    out.openingDate = longReportDate(range[0]);
    out.closingDate = longReportDate(range[1]);
  }

  const headings: [string, string][] = [
    ['Australasian Equities', 'Australasian Equities'],
    ['Global Equities', 'Global Equities'],
    ['Property', 'Property'],
    ['New Zealand Fixed Interest', 'New Zealand Fixed Interest'],
    ['Global Fixed Interest', 'International Fixed Interest'],
    ['Multi Sector', 'New Zealand Diversified'],
  ];
  for (const [heading, search] of headings) {
    const h = findLine(lines, search);
    let parts: string[] | null = null;
    if (h >= 0) {
      for (let r = h + 1; r < lines.length; r++) {
        const p = lines[r]!.split(' ');
        if (isNumericWord(p[0] ?? '') && p.length >= 8) {
          parts = p;
          break;
        }
      }
    }
    if (!parts) {
      out.sectors.push({ heading, percent: null, dollars: null });
      continue;
    }
    const n = parts.length;
    const sum = vbaVal(parts[n - 3]!.replace(/,/g, '')) + vbaVal(parts[n - 1]!.replace(/,/g, ''));
    out.sectors.push({ heading, percent: parts[n - 2]!, dollars: dollars0(sum) });
  }

  const rc = findLine(lines, 'Report Class Current % NZD Amount');
  if (rc < 0) missing.push('Report Class Current % NZD Amount');
  else {
    let property = 0;
    let fi = 0;
    let anyProperty = false;
    let anyFi = false;
    for (let i = rc + 1; i < lines.length; i++) {
      const line = lines[i]!;
      if (line.startsWith('Total')) break;
      const words = line.split(' ');
      const headingWords: string[] = [];
      for (const w of words) {
        if (isNumericWord(w.replace(/,/g, '')) || w.includes('%')) break;
        headingWords.push(w);
      }
      let heading = headingWords.join(' ').trim();
      if (!heading) continue;
      if (heading.includes('Cash and Cash Equivalents')) heading = 'Cash';
      if (heading.includes('New Zealand Diversified')) heading = 'Multi Sector';
      const pctWord = [...words].reverse().find((w) => w.includes('%'));
      if (!pctWord) continue;
      const v = Number(pctWord.replace('%', '').replace(/,/g, '')) / 100;
      if (!Number.isFinite(v)) continue;
      if (heading.includes('Property')) {
        property += v;
        anyProperty = true;
      } else if (heading.includes('Fixed Interest')) {
        fi += v;
        anyFi = true;
      } else {
        const cls = ALLOCATION_CLASSES.find((c) => c.includes(heading) || heading.includes(c));
        if (cls) out.current[cls] = v;
      }
    }
    if (anyProperty) out.current.Property = property;
    if (anyFi) out.current['Fixed Interest'] = fi;
  }
  return out;
}

// ---------------------------------------------------------------- OneAnswer

/** ExtractFirstDollarAmount */
function firstDollar(s: string): number {
  const m = /-?\$[\d,]+\.?\d*/.exec(s);
  return m ? Number(m[0].replace(/[$,]/g, '')) || 0 : 0;
}

export function parseOAReport(text: string): PerformanceReport {
  const lines = reportLines(text);
  const missing: string[] = [];
  const out: PerformanceReport = { custodian: 'OA', openingDate: null, closingDate: null, inception: null, table1: [], dollarGain: null, percentGain: null, sectors: [], current: emptyCurrent(), missing };

  const inc = findLine(lines, 'Performance Summary from Inception');
  if (inc < 0) missing.push('Performance Summary from Inception');
  else {
    const matches = [...(lines[inc + 17] ?? '').matchAll(/\b\d+(\.\d+)?%/g)];
    const last = matches.at(-1)?.[0];
    out.inception = inceptionFrom(dateRange(lines[inc + 1]), last ? Number(last.slice(0, -1)) / 100 : null);
  }

  const ps = findLine(lines, 'Performance Summary from');
  if (ps < 0) missing.push('Performance Summary from');
  else {
    const range = dateRange(lines[ps + 1]);
    if (range) {
      out.openingDate = longReportDate(range[0]);
      out.closingDate = longReportDate(range[1]);
    }
    const v = (o: number) => firstDollar(lines[ps + o] ?? '');
    const opening = v(4);
    const contributions = v(5);
    const withdrawals = -Math.abs(v(6));
    const fourth = v(7) + v(8);
    const firstFour = opening + contributions + withdrawals + v(7);
    const closing = v(17);
    const gain = closing - firstFour;
    const pm = /\b\d+\.?\d*%/.exec(lines[ps + 19] ?? '');
    out.table1 = [
      { text: dollars2(opening), negative: opening < 0 },
      { text: dollars2(contributions), negative: contributions < 0 },
      { text: withdrawals < 0 ? `(${dollars2(-withdrawals)})` : dollars2(0), negative: withdrawals < 0 },
      { text: dollars2(fourth), negative: fourth < 0 },
      { text: dollars2(firstFour), negative: firstFour < 0 },
      { text: dollars2(closing), negative: closing < 0 },
      { text: dollars2(gain), negative: gain < 0 },
      { text: pm ? pm[0] : '', negative: false },
    ];
    out.dollarGain = dollars2(gain);
    out.percentGain = pm ? pm[0] : null;
  }

  // Table2: the dollar return from the "Totals" line under each heading (column 2 is not filled by the macro)
  for (const [heading, search] of [
    ['Australasian Equities', 'Australasian Equities'],
    ['Global Equities', 'Global Equities'],
    ['Property', 'Property'],
    ['New Zealand Fixed Interest', 'New Zealand Fixed Interest'],
    ['Global Fixed Interest', 'International Fixed Interest'],
    ['Multi Sector', 'Multi Sector'],
  ] as const) {
    const h = findLine(lines, search, true);
    let dollars: string | null = null;
    if (h >= 0) {
      for (let r = h + 1; r < lines.length; r++) {
        if (lines[r]!.startsWith('Totals')) {
          const parts = lines[r]!.split(' ');
          dollars = dollars2(vbaVal(parts[parts.length - 1]!.replace(/,/g, '')));
          break;
        }
      }
    }
    out.sectors.push({ heading, percent: null, dollars });
  }

  // Table3: the CLASS TOTAL percentage under the second occurrence of each class heading
  const classTotal = (heading: string): number | null => {
    const at = findLine(lines, heading, false, 2);
    if (at < 0) return null;
    for (let r = at + 1; r < lines.length; r++) {
      if (lines[r]!.startsWith('CLASS TOTAL')) {
        const p = lines[r]!.split(' ').filter((w) => w.includes('%')).at(-1);
        return p ? vbaVal(p.replace('%', '')) / 100 : null;
      }
    }
    return null;
  };
  for (const cls of ALLOCATION_CLASSES) {
    if (cls === 'Fixed Interest') {
      const nz = classTotal('New Zealand Fixed Interest') ?? 0;
      const intl = classTotal('International Fixed Interest') ?? 0;
      out.current[cls] = nz + intl > 0 ? nz + intl : null;
    } else {
      out.current[cls] = classTotal(cls);
    }
  }
  return out;
}

/** Which custodian a pasted report is from: OneAnswer reports start with "Printed". */
export function parsePerformanceReport(text: string): PerformanceReport {
  const first = reportLines(text).find((l) => l !== '') ?? '';
  return first.startsWith('Printed') ? parseOAReport(text) : parseApexReport(text);
}
