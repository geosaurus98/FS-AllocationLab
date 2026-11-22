// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import type { AdviceTables } from '../engine/workbook/documentTables';
import { DocxDocument } from './docx';
import { buildLetter, ddmmyyyy, longDate, type LetterClient } from './letters';
import { entryText, readZip, writeZip } from './zip';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const run = (t: string, bold = false) => `<w:r>${bold ? '<w:rPr><w:b/></w:rPr>' : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const para = (...runs: string[]) => `<w:p>${runs.join('')}</w:p>`;
const sigTable = (name: string) =>
  `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>${[`[sig ${name}]`, name, 'Financial Adviser'].map((t) => `<w:tr><w:tc>${para(run(t))}</w:tc></w:tr>`).join('')}</w:tbl>`;

/** A synthetic template shaped like the firm's: placeholders, some split across runs. */
async function template(): Promise<Uint8Array> {
  const body = [
    para(run('Full_'), run('Name', true)),
    para(run('Street_Add')),
    para(run('Suburb'), run(' '), run('City_Post')),
    para(run('Long_Date')),
    para(run('Dear First_Name,')),
    para(run('Your portfolio and your goals. We considered youre needs.')),
    para(run('Investment'), run('rationale')),
    para(run('Portfolioforadvice')),
    `<w:p><w:bookmarkStart w:id="1" w:name="Cost1"/>${run('Costs heading')}<w:bookmarkEnd w:id="1"/></w:p>`,
    para(run('InsertTable')),
    para(run('ChangeTable')),
    `<w:p><w:bookmarkStart w:id="2" w:name="AdvisorTable"/><w:bookmarkEnd w:id="2"/></w:p>`,
    `<w:p><w:bookmarkStart w:id="3" w:name="John"/></w:p>${sigTable('John Example')}<w:p><w:bookmarkEnd w:id="3"/></w:p>`,
    `<w:p><w:bookmarkStart w:id="4" w:name="Hayden"/></w:p>${sigTable('Hayden Example')}<w:p><w:bookmarkEnd w:id="4"/></w:p>`,
    para(run('Kind regards')),
  ].join('');
  const doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${body}<w:sectPr/></w:body></w:document>`;
  return writeZip([
    {
      name: '[Content_Types].xml',
      data: '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    },
    { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' },
    { name: 'word/document.xml', data: doc },
  ]);
}

const tables: AdviceTables = {
  rationale: [
    { kind: 'title', text: 'Investment Rationale' },
    { kind: 'heading', text: 'Global Equities' },
    { kind: 'bullet', text: 'Add a new investment in Example Fund.', factSheet: 'https://example.com/fact.pdf' },
    { kind: 'bullet', text: 'Exit existing investment in Other Fund.' },
  ],
  portfolio: { widths: [170, 81, 81, 81, 71], fontSize: 10, center: true, edgeBorders: false, rows: [[{ text: '' }, { text: 'Portfolio' }, { text: 'Changes' }, { text: 'Recomm. Portfolio' }, { text: '% Allocation' }], [{ text: 'Portfolio Total', bold: true }, { text: '$100' }, { text: '$0' }, { text: '$100' }, { text: '100.00' }]] },
  costs: { widths: [163, 68, 69, 50, 50, 50, 75], fontSize: 8, center: true, edgeBorders: true, rows: [[{ text: '' }, { text: 'Broker' }, { text: 'Brokerage' }, { text: 'Buy' }, { text: 'Sell' }, { text: 'Cost p.a.' }, { text: 'Product Disclosure' }], [{ text: 'Example Fund' }, { text: '-' }, { text: '-' }, { text: '0.00%' }, { text: '0.00%' }, { text: '0.20%' }, { text: 'Click here', link: 'https://example.com/pds.pdf' }]] },
  summary: { widths: [300, 50, 70, 80], fontSize: 10, center: false, edgeBorders: true, summaryBorders: true, rows: [[{ text: 'Sample Client' }, { text: 'Action' }, { text: 'Changes' }, { text: 'Recomm. Portfolio' }], [{ text: '' }, { text: '' }, { text: '' }, { text: '' }], [{ text: 'Totals', bold: true }, { text: '' }, { text: '$0' }, { text: '$100' }]] },
};

const client: LetterClient = { fullName: 'Ms A Sample', firstName: 'Alex', street: '1 Example Street', suburb: 'Sampleton', city: 'Wellington 6011', clientType: 'Individual', fileName: 'Sample', leadAdviser: 'John', secondaryAdviser: 'Hayden' };
const today = new Date(2026, 10, 21);

describe('Letter of Advice from a template', () => {
  it('fills text, tables, rationale and signatures', async () => {
    const out = await buildLetter({ kind: 'loa', template: await template(), tables, client, today });
    expect(out.fileName).toBe('Sample LOA 21112026.docx');
    expect(out.notes).toEqual([]);
    const doc = await DocxDocument.open(out.bytes);
    const text = doc.text();
    expect(text).toContain('Ms A Sample');
    expect(text).not.toContain('Full_');
    expect(text).toContain('1 Example Street');
    expect(text).toContain('Sampleton Wellington 6011');
    expect(text).toContain('21 November 2026');
    expect(text).toContain('Dear Alex,');
    expect(text).toContain('Your portfolio and your goals. We considered your needs.');
    expect(text).toContain('Add a new investment in Example Fund. Link to the fact sheet here.');
    for (const p of ['Investmentrationale', 'Portfolioforadvice', 'InsertTable', 'ChangeTable']) expect(text).not.toContain(p);
    // adviser source tables deleted; signature table built from their first columns
    const tbls = doc.doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'tbl');
    expect(tbls.length).toBe(4); // portfolio, costs, summary, signature
    expect(text).toContain('[sig John Example]');
    expect(text.match(/John Example/g)).toHaveLength(2); // once in the signature table: "[sig John Example]" and "John Example"
    expect(doc.hasBookmark('John')).toBe(false);
    expect(doc.hasBookmark('Hayden')).toBe(false);
    // hyperlinks registered as external relationships
    const rels = await entryText(readZip(out.bytes).find((e) => e.name === 'word/_rels/document.xml.rels')!);
    expect(rels).toContain('https://example.com/fact.pdf');
    expect(rels).toContain('https://example.com/pds.pdf');
    expect(rels).toContain('numbering.xml');
    const ct = await entryText(readZip(out.bytes).find((e) => e.name === '[Content_Types].xml')!);
    expect(ct).toContain('/word/numbering.xml');
  });

  it('uses trust wording and removes the cost section when there are no costs', async () => {
    const out = await buildLetter({ kind: 'loa', template: await template(), tables: { ...tables, costs: null }, client: { ...client, clientType: 'Trust' }, today });
    const doc = await DocxDocument.open(out.bytes);
    const text = doc.text();
    expect(text).toContain("The Trustees' portfolio and the Trustees' goals. We considered the Beneficiaries' needs.");
    expect(text).not.toContain('Costs heading');
    expect(out.notes.some((n) => n.includes('Cost1'))).toBe(true);
  });

  it('reports missing placeholders', async () => {
    const out = await buildLetter({ kind: 'email', template: await template(), tables, client, today });
    expect(out.fileName).toBe(`EMAIL Sample ${ddmmyyyy(today)}.docx`);
    expect(out.notes.some((n) => n.includes('Portfolio_for_advice'))).toBe(true);
  });
});

describe('without a template', () => {
  it('produces a plain document with the tables', async () => {
    const out = await buildLetter({ kind: 'loa', template: null, tables, client, today });
    const doc = await DocxDocument.open(out.bytes);
    const text = doc.text();
    expect(text).toContain('Letter of Advice tables: Ms A Sample');
    expect(text).toContain(longDate(today));
    expect(text).toContain('Investment Rationale');
    expect(text).toContain('Click here');
  });
});

describe('text replacement', () => {
  it('matches whole words, keeps case, and works across runs', async () => {
    const d = DocxDocument.blank([{ text: 'your yours Your YOUR' }]);
    expect(d.replaceText('your', 'the Trustees’')).toBe(3);
    expect(d.text()).toBe('the Trustees’ yours The Trustees’ THE TRUSTEES’');
  });
});
