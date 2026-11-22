// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { adviceTables, soaTable } from '../engine/workbook/documentTables';
import { Workbook, type WorkbookModel } from '../engine/workbook/evaluator';
import { applyScenario, type Scenario } from '../engine/workbook/scenario';
import { buildAnnualReview, buildStatementOfAdvice, emptyReviewDetails } from './annualReview';
import { DocxDocument } from './docx';
import type { LetterClient } from './letters';
import { excelValue, parseApexReport, parseOAReport, parsePerformanceReport } from './performanceReport';
import { entryText, readZip, writeZip } from './zip';

// ---------------------------------------------------------------- synthetic reports (shaped as the macros expect)

const APEX = [
  'Portfolio Performance Report',
  'Portfolio: DEMO-1',
  'Net Return Since Inception',
  '1 Jul 2019 to 30 Jun 2025',
  'x', 'x', 'x', 'x', 'x', 'x', 'x',
  '45.67%',
  'Net Portfolio Return',
  '$500,000.00',
  '$20,000.00',
  '-$10,000.00',
  '$5,000.00',
  '$30,000.00',
  '$545,000.00',
  '$35,000.00',
  '7.00%',
  'Net Return for Period',
  '1 Jul 2024 to 30 Jun 2025',
  'Australasian Equities',
  'Code Name Value',
  '1 Fund A 10,000.00 500.00 1,000.00 12.34% 300.00',
  'Global Equities',
  '1 Fund B 20,000.00 500.00 2,000.00 9.50% -100.00',
  'Property',
  '1 Fund C 5,000.00 0.00 -200.00 -1.10% 150.00',
  'New Zealand Fixed Interest',
  '1 Fund D 50,000.00 0.00 1,500.00 4.00% 0.00',
  'International Fixed Interest',
  '1 Fund E 30,000.00 0.00 900.00 3.10% 0.00',
  'Report Class Current % NZD Amount',
  'Cash and Cash Equivalents 2.00% 10,900.00',
  'New Zealand Fixed Interest 30.00% 163,500.00',
  'International Fixed Interest 10.00% 54,500.00',
  'Australasian Equities 20.00% 109,000.00',
  'Global Equities 30.00% 163,500.00',
  'Property 8.00% 43,600.00',
  'Total 100.00% 545,000.00',
].join('\n');

const OA = [
  'Printed 21/11/2026 - DEMO-2 - Sample',
  'Performance Summary from Inception',
  '1 Jul 2018 to 30 Jun 2025',
  ...Array(15).fill('x'),
  'Total return since inception 52.30%',
  'Performance Summary from',
  '1 Jul 2024 to 30 Jun 2025',
  'x',
  'x',
  'Opening value $400,000.00',
  'Contributions $10,000.00',
  'Withdrawals $8,000.00',
  'Income $6,000.00',
  'Fees $0.00',
  ...Array(7).fill('x'),
  'x',
  'Closing value $430,000.00',
  'x',
  'Return for period 5.75%',
  'Australasian Equities',
  'Totals 1 2 3 4 5 6 7 4,321.00',
  'Global Equities',
  'Totals 1 2 3 4 5 6 7 -1,234.50',
  'Cash',
  'Cash',
  'CLASS TOTAL 1,000.00 3.00%',
  'Australasian Equities',
  'CLASS TOTAL 2,000.00 25.00%',
  'Global Equities',
  'CLASS TOTAL 2,000.00 32.00%',
  'New Zealand Fixed Interest',
  'New Zealand Fixed Interest',
  'CLASS TOTAL 2,000.00 30.00%',
  'International Fixed Interest',
  'International Fixed Interest',
  'CLASS TOTAL 2,000.00 10.00%',
].join('\n');

describe('excel paste values', () => {
  it('converts number-like text as Excel does', () => {
    expect(excelValue('$1,234.50')).toBe(1234.5);
    expect(excelValue('-$1,234.50')).toBe(-1234.5);
    expect(excelValue('($1,234.50)')).toBe(-1234.5);
    expect(excelValue('12.5%')).toBeCloseTo(0.125);
    expect(excelValue('1 Jul 2024 to 30 Jun 2025')).toBe('1 Jul 2024 to 30 Jun 2025');
    expect(excelValue('12,34')).toBe('12,34');
  });
});

describe('Apex performance report', () => {
  const r = parseApexReport(APEX);
  it('reads the period, inception and Table1', () => {
    expect(r.missing).toEqual([]);
    expect(r.openingDate).toBe('1 July 2024');
    expect(r.closingDate).toBe('30 June 2025');
    expect(r.inception!.days).toBe(2191);
    expect(r.inception!.arithmetic).toBeCloseTo((0.4567 / 2191) * 365, 10);
    expect(r.table1.map((x) => x.text)).toEqual(['$500,000.00', '$20,000.00', '($10,000.00)', '$5,000.00', '$30,000.00', '$545,000.00', '$35,000.00', '7.00%']);
    expect(r.table1[2]!.negative).toBe(true);
    expect(r.dollarGain).toBe('$35,000.00');
    expect(r.percentGain).toBe('7.00%');
  });
  it('reads sector returns and the current allocation', () => {
    expect(r.sectors.map((s) => [s.heading, s.percent, s.dollars])).toEqual([
      ['Australasian Equities', '12.34%', '$1,300'],
      ['Global Equities', '9.50%', '$1,900'],
      ['Property', '-1.10%', '-$50'],
      ['New Zealand Fixed Interest', '4.00%', '$1,500'],
      ['Global Fixed Interest', '3.10%', '$900'],
      ['Multi Sector', null, null],
    ]);
    expect(r.current.Cash).toBeCloseTo(0.02);
    expect(r.current['Fixed Interest']).toBeCloseTo(0.4);
    expect(r.current['Australasian Equities']).toBeCloseTo(0.2);
    expect(r.current.Property).toBeCloseTo(0.08);
    expect(r.current['Multi Sector']).toBeNull();
  });
});

describe('OneAnswer performance report', () => {
  const r = parsePerformanceReport(OA);
  it('is detected and parsed', () => {
    expect(r.custodian).toBe('OA');
    expect(r).toEqual(parseOAReport(OA));
    expect(r.closingDate).toBe('30 June 2025');
    expect(r.inception!.days).toBeGreaterThan(1825);
    expect(r.table1.map((x) => x.text)).toEqual(['$400,000.00', '$10,000.00', '($8,000.00)', '$6,000.00', '$408,000.00', '$430,000.00', '$22,000.00', '5.75%']);
    expect(r.sectors.find((s) => s.heading === 'Global Equities')!.dollars).toBe('-$1,234.50');
    expect(r.current.Cash).toBeCloseTo(0.03);
    expect(r.current['Fixed Interest']).toBeCloseTo(0.4);
    expect(r.current.Property).toBeNull();
  });
});

// ---------------------------------------------------------------- synthetic Annual Review template

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const run = (t: string) => `<w:r><w:t xml:space="preserve">${t}</w:t></w:r>`;
const p = (...x: string[]) => `<w:p>${x.join('')}</w:p>`;
const cell = (t: string, inner = '') => `<w:tc>${p(inner, run(t))}</w:tc>`;
const bm = (id: number, name: string) => `<w:bookmarkStart w:id="${id}" w:name="${name}"/><w:bookmarkEnd w:id="${id}"/>`;
const tbl = (rows: string[][], mark?: { row: number; col: number; id: number; name: string }) =>
  `<w:tbl><w:tblPr/>${rows.map((r, i) => `<w:tr>${r.map((t, j) => cell(t, mark && mark.row === i && mark.col === j ? bm(mark.id, mark.name) : '')).join('')}</w:tr>`).join('')}</w:tbl>`;

async function arTemplate(): Promise<Uint8Array> {
  const body = [
    p(run('Annual Review for Client_Name – Year_Date')),
    p(run('Strategy: Strategy_Name. I am / We are pleased; I / We review my / our advice.')),
    p(run('Insert previous Blurb')),
    p(run('RWT rwtXXX%. PIR pirXXX%. Talk to your accountant.')),
    p(run('holder6 objectives and holder2.')),
    `<w:p><w:bookmarkStart w:id="1" w:name="AuthorityToAct"/>${run('Authority to act wording')}<w:bookmarkEnd w:id="1"/></w:p>`,
    p(run('Period Ending: Short_Date')),
    `<w:p><w:bookmarkStart w:id="2" w:name="OpeningPV"/>${run('Opening value')}<w:bookmarkEnd w:id="2"/></w:p>`,
    tbl([['Opening', '$'], ['Contributions', '$'], ['Withdrawals', '$'], ['Income', '$'], ['Growth', '$'], ['Closing', '$'], ['Gain', '$'], ['Return', '%'], ['arith', ''], ['geo', '']], { row: 0, col: 1, id: 3, name: 'Table1' }),
    p(run('You gained dollarGain (percentGain).')),
    tbl([['Sector', 'Return', 'Gain'], ['Australasian Equities', '%', '$'], ['Global Equities', '%', '$'], ['Property', '%', '$'], ['New Zealand Fixed Interest', '%', '$'], ['Global Fixed Interest', '%', '$'], ['Multi Sector', '%', '$']], { row: 1, col: 0, id: 4, name: 'Table2' }),
    tbl([['Class', 'Target', 'Range', 'Current'], ['Cash', '%', '', '%'], ['Fixed Interest', '%', '', '%'], ['Australasian Equities', '%', '', '%'], ['Global Equities', '%', '', '%'], ['Property', '%', '', '%'], ['Multi Sector', '%', '', '%'], ['Total', '', '', '']], { row: 0, col: 0, id: 5, name: 'Table3' }),
    `<w:p>${bm(6, 'OpeningComments')}</w:p>`,
    p(run('Sig1')),
    p(run('and')),
    `<w:p>${bm(7, 'Sig2')}${run('Sig2')}</w:p>`,
    p(run('Investmentrationale')),
    p(run('Portfolioforadvice')),
    p(run('InsertTable')),
    p(run('ChangeTable')),
  ].join('');
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>`;
  return writeZip([
    { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>' },
    { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
    { name: 'word/document.xml', data: doc },
    { name: 'word/settings.xml', data: `<?xml version="1.0" encoding="UTF-8"?><w:settings ${NS}><w:zoom w:percent="100"/><w:compat/></w:settings>` },
  ]);
}

async function commentary(): Promise<Uint8Array> {
  const doc = `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${p(run('Markets rose this year.'))}${p(run('Outlook remains uncertain.'))}<w:sectPr/></w:body></w:document>`;
  return writeZip([
    { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>' },
    { name: 'word/document.xml', data: doc },
  ]);
}

const model: WorkbookModel = JSON.parse(readFileSync(resolve(__dirname, '../engine/workbook/model.json'), 'utf8'));
const fixtures: { scenario: Scenario }[] = JSON.parse(readFileSync(resolve(__dirname, '../engine/workbook/__golden_rounds__/rounds.json'), 'utf8'));
function workbook(): Workbook {
  const wb = new Workbook(model, new Date(2026, 10, 20));
  wb.libreOfficeCompat = true;
  applyScenario(wb, fixtures[0]!.scenario);
  wb.set('Portfolio', 'BD1', 'Apex');
  return wb;
}

const client: LetterClient = { fullName: 'Ms A Sample', firstName: 'Alex', street: '', suburb: '', city: '', clientType: 'Individual', fileName: 'Sample', leadAdviser: '', secondaryAdviser: '' };
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

describe('Annual Review', () => {
  it('fills the template from the report and the workbook', async () => {
    const wb = workbook();
    const out = await buildAnnualReview({
      template: await arTemplate(),
      report: parseApexReport(APEX),
      targets: { Cash: 2, 'Fixed Interest': 40, 'Australasian Equities': 20, 'Global Equities': 30, Property: 8, 'Multi Sector': 0 },
      tables: adviceTables(wb, { clientName: 'Ms A Sample', custodian: 'Apex' }),
      client,
      review: { ...emptyReviewDetails(), strategyName: 'Balanced', blurb: 'Previous blurb.', signatories: 'Alex Sample', iOrWe: 'I', rwt: '33', authorityToAct: true },
      commentary: await commentary(),
      today: new Date(2026, 10, 21),
    });
    expect(out.fileName).toBe('Sample Annual Review 211126.docx');
    const doc = await DocxDocument.open(out.bytes);
    const text = doc.text();
    expect(text).toContain('Annual Review for Ms A Sample – November 2026');
    expect(text).toContain('Strategy: Balanced. I am pleased; I review my advice.');
    expect(text).toContain('Previous blurb.');
    expect(text).toContain('RWT 33%. PIR XXX%.');
    expect(text).toContain('Your objectives and your life.');
    expect(text).not.toContain('Authority to act wording');
    expect(text).toContain('Period Ending: 30 June 2025');
    expect(text).toContain('Opening Portfolio Value - 1 July 2024');
    expect(text).toContain('You gained $35,000.00 (7.00%).');
    expect(text).toContain('Markets rose this year.');
    expect(text).toContain('Alex Sample');
    expect(text).not.toMatch(/\band\n/); // the empty Sig2 paragraph and the one before it are removed
    for (const ph of ['Investmentrationale', 'Portfolioforadvice', 'InsertTable', 'ChangeTable']) expect(text).not.toContain(ph);

    const tables = Array.from(doc.doc.getElementsByTagNameNS(W, 'tbl')).filter((t) => !t.parentNode || (t.parentNode as Element).localName === 'body');
    const grid = (t: Element) => doc.tableRows(t).map((r) => doc.rowCells(r).map((c) => doc.cellText(c).trim()));
    const t1 = grid(tables[0]!);
    expect(t1.map((r) => r[1])).toEqual(['$500,000.00', '$20,000.00', '($10,000.00)', '$5,000.00', '$30,000.00', '$545,000.00', '$35,000.00', '7.00%', expect.stringMatching(/%$/), expect.stringMatching(/%$/)]);
    expect(t1[8]![0]).toBe('Net Portfolio Return p.a. from 01 Jul 19 to 30 Jun 25 (arithmetic average)');
    const t2 = grid(tables[1]!);
    expect(t2.map((r) => r[0])).not.toContain('Multi Sector'); // no figures → row removed
    expect(t2[1]).toEqual(['Australasian Equities', '12.34%', '$1,300']);
    const t3 = grid(tables[2]!);
    expect(t3.map((r) => r[0])).toEqual(['Class', 'Cash', 'Fixed Interest', 'Australasian Equities', 'Global Equities', 'Property', 'Total']);
    expect(t3[2]).toEqual(['Fixed Interest', '40.00%', '', '40.00%']);
    expect(t3[6]).toEqual(['Total', '100.00%', '', '100.00%']);
    expect(t3[1]![2]).toBe('+/- 10% Income Assets');
    expect(t3[3]![2]).toBe('+/- 10% Growth Assets');
    // XXX and the missing accountant are highlighted; Word updates the table of contents on open
    expect(doc.doc.getElementsByTagNameNS(W, 'highlight').length).toBeGreaterThanOrEqual(2);
    const settings = await entryText(readZip(out.bytes).find((e) => e.name === 'word/settings.xml')!);
    expect(settings).toMatch(/<w:updateFields w:val="true"\/><w:compat/);
  });
});

describe('Statement of Advice', () => {
  it('lists held and recommended holdings with Data sheet details', async () => {
    const wb = workbook();
    const t = soaTable(wb, 'Apex');
    expect(t.rows[0]!.map((c) => c.text)).toEqual(['Investments', 'Manager or Company Name', 'Website', 'Description', 'Broker', 'Brokerage', 'Buy', 'Sell', 'Cost p.a', 'Fact Sheet', 'Product Disclosure']);
    expect(t.rows.length).toBeGreaterThan(10);
    const names = t.rows.slice(1).map((r) => r[0]!.text);
    expect(names).not.toContain('NZD');
    for (const r of t.rows.slice(1)) {
      expect(r).toHaveLength(11);
      if (r[10]!.link) expect(r[10]!.text).toBe('Click Here');
      if (r[2]!.link) expect(r[2]!.text).not.toMatch(/^https?:/);
    }
    // into a template: landscape section for the table, portrait afterwards
    const body = `${p(run('Appendix for strategy_name clients'))}<w:p>${bm(1, 'SOA_Table')}</w:p>${p()}${p(run('After the table'))}`;
    const tpl = await writeZip([
      { name: '[Content_Types].xml', data: '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>' },
      { name: 'word/document.xml', data: `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>` },
    ]);
    const out = await buildStatementOfAdvice({ template: tpl, table: t, strategyName: 'Balanced', clientName: 'Sample', today: new Date(2026, 10, 21) });
    expect(out.fileName).toBe('SOA Sample Appendices 211126.docx');
    const doc = await DocxDocument.open(out.bytes);
    expect(doc.text()).toContain('Appendix for Balanced clients');
    const sects = Array.from(doc.doc.getElementsByTagNameNS(W, 'sectPr'));
    expect(sects).toHaveLength(2);
    expect(sects[0]!.getElementsByTagNameNS(W, 'pgSz')[0]!.getAttribute('w:orient')).toBe('landscape');
    expect(sects[1]!.getElementsByTagNameNS(W, 'pgSz')[0]!.getAttribute('w:orient')).toBeNull();
  });
});
