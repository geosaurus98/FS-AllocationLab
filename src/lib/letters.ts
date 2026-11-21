/**
 * Advice documents built from the firm's Word templates: ports of the workbook macros
 * Letter_Of_Advice, Email_Advice and Follow_Up_Letter.
 *
 * The macros read client details from the "Master Portfolio Risk Profiles" workbook and the templates
 * from the P: drive. AllocationLab has neither: the adviser types the details and chooses the
 * template file. Nothing is uploaded; the filled document is downloaded by the browser.
 */

import type { AdviceTables, DocTable } from '../engine/workbook/documentTables';
import { DocxDocument } from './docx';

export type LetterKind = 'loa' | 'email' | 'followup';

export const CLIENT_TYPES = ['Individual', 'Joint', 'Trust', 'Estate', 'Charity', 'Company'] as const;
export type ClientType = (typeof CLIENT_TYPES)[number];

/** Signature bookmarks in the Letter of Advice template; deleted once the signature table is built. */
export const ADVISER_BOOKMARKS = ['John', 'Hayden', 'Claudia'] as const;
const DELETE_AFTER_SIGNATURE = ['John', 'Hayden', 'Claudia', 'BKMRK'];

export interface LetterClient {
  /** Corrospondence_Name, e.g. "Mr J and Mrs A Smith". */
  fullName: string;
  firstName: string;
  street: string;
  suburb: string;
  city: string;
  clientType: ClientType;
  /** ExcelssName: used in the file name. */
  fileName: string;
  leadAdviser: string;
  secondaryAdviser: string;
}

interface KindSpec {
  label: string;
  placeholders: { rationale: string; portfolio: string; costs: string; summary: string };
  layout: { portfolio: Partial<DocTable>; costs: Partial<DocTable>; summary: Partial<DocTable> };
  fileName: (c: LetterClient, ddmmyyyy: string) => string;
}

export const LETTER_KINDS: Record<LetterKind, KindSpec> = {
  loa: {
    label: 'Letter of Advice',
    placeholders: { rationale: 'Investmentrationale', portfolio: 'Portfolioforadvice', costs: 'InsertTable', summary: 'ChangeTable' },
    layout: {
      portfolio: { widths: [170, 81, 81, 81, 71], fontSize: 10, center: true },
      costs: { widths: [163, 68, 69, 50, 50, 50, 75], fontSize: 8, center: true },
      summary: { widths: [300, 50, 70, 80], fontSize: 10, center: false },
    },
    fileName: (c, d) => `${c.fileName} LOA ${d}`,
  },
  email: {
    label: 'Email Advice',
    placeholders: { rationale: 'Investmentrationale', portfolio: 'Portfolio_for_advice', costs: 'Costs_For_Advice', summary: 'ChangeTable' },
    layout: {
      portfolio: { widths: [315, 85, 85, 85, 71], fontSize: 10, center: false },
      costs: { widths: [305, 70, 70, 55, 55, 55, 80], fontSize: 10, center: false },
      summary: { widths: [460, 80, 80, 100], fontSize: 10, center: false },
    },
    fileName: (c, d) => `EMAIL ${c.fileName} ${d}`,
  },
  followup: {
    label: 'Follow-up letter',
    placeholders: { rationale: 'Investmentrationale', portfolio: 'Portfolioforadvice', costs: 'InsertTable', summary: 'ChangeTable' },
    layout: {
      portfolio: { widths: [170, 81, 81, 81, 71], fontSize: 10, center: false },
      costs: { widths: [163, 68, 69, 50, 50, 50, 75], fontSize: 8, center: false },
      summary: { widths: [300, 50, 70, 80], fontSize: 10, center: false },
    },
    fileName: (c, d) => `${c.fileName} Follow-up ${d}`,
  },
};

/** Email Advice: holder0–holder4 wording by client type. */
const HOLDER_WORDS: Record<string, [string, string, string, string, string]> = {
  trust: ["the Beneficiaries'", "the Trustees' overall", 'the Trustees’', 'the Trustees', 'The Trustees’'],
  estate: ['the Estate’s', 'the Estate’s', 'the Executor’s', 'the Executor', 'The Executor’s'],
  charity: ['the Charity’s', 'the Charity’s', 'the Trustees’', 'the Trustees', 'The Trustees’'],
  joint: ['your', 'your', 'your', 'you both', 'Your'],
  company: ['the Company’s', 'the Company’s', 'the Directors’', 'the Directors', 'The Directors’'],
  individual: ['your', 'your', 'your', 'you', 'Your'],
};

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
export const longDate = (d: Date) => `${d.getDate()} ${MONTHS[d.getMonth()]} ${d.getFullYear()}`;
export const ddmmyyyy = (d: Date) => `${String(d.getDate()).padStart(2, '0')}${String(d.getMonth() + 1).padStart(2, '0')}${d.getFullYear()}`;

export interface LetterResult {
  bytes: Uint8Array;
  fileName: string;
  /** Things the adviser should know: placeholders or bookmarks not in the template, etc. */
  notes: string[];
}

/** Placeholder paragraphs for the plain document produced when no template is loaded. */
function blankTemplate(kind: LetterKind, tables: AdviceTables): DocxDocument {
  const p = LETTER_KINDS[kind].placeholders;
  return DocxDocument.blank([
    { text: `${LETTER_KINDS[kind].label} tables: Full_Name`, style: 'title' },
    { text: 'Long_Date', style: 'note' },
    { text: 'Produced by AllocationLab without the firm template. Copy the tables into the letter and check every figure before use.', style: 'note' },
    { text: p.rationale },
    { text: 'Portfolio for advice', style: 'heading' },
    { text: p.portfolio },
    ...(tables.costs ? [{ text: 'Costs for advice', style: 'heading' as const }, { text: p.costs }] : []),
    { text: 'Summary of changes', style: 'heading' },
    { text: p.summary },
  ]);
}

export async function buildLetter(args: {
  kind: LetterKind;
  /** The Word template (.docx); null for a plain document with just the tables. */
  template: Uint8Array | ArrayBuffer | null;
  tables: AdviceTables;
  client: LetterClient;
  today: Date;
}): Promise<LetterResult> {
  const { kind, tables, client, today } = args;
  const spec = LETTER_KINDS[kind];
  const doc = args.template ? await DocxDocument.open(args.template) : blankTemplate(kind, tables);
  const notes: string[] = [];

  // ---- text (ReplaceTextInWord: whole words, case-insensitive)
  if (kind === 'loa' || !args.template) {
    doc.replaceText('Full_Name', client.fullName || 'Full_Name');
    doc.replaceText('Long_Date', longDate(today));
  }
  if (kind === 'loa') {
    doc.replaceText('First_Name', client.firstName);
    doc.replaceText('Street_Add', client.street || 'Street');
    doc.replaceText('Suburb', client.suburb || 'Suburb');
    doc.replaceText('City_Post', client.city || 'City');
    if (client.clientType === 'Trust' || client.clientType === 'Estate') {
      doc.replaceText('your', "the Trustees'");
      doc.replaceText('youre', "the Beneficiaries'");
    } else {
      doc.replaceText('youre', 'your');
    }
    if (args.template) {
      const sig = doc.signature(client.leadAdviser, client.secondaryAdviser, DELETE_AFTER_SIGNATURE);
      if (!sig.built) notes.push('The template has no "AdvisorTable" bookmark, so no signature table was added.');
      for (const m of sig.missing) notes.push(`No signature bookmark "${m}" with a table in the template.`);
    }
  } else if (kind === 'email') {
    doc.replaceText('First_Name', client.firstName);
    const words = HOLDER_WORDS[client.clientType.toLowerCase()] ?? HOLDER_WORDS.individual!;
    words.forEach((w, i) => doc.replaceText(`holder${i}`, w));
  }

  // ---- tables and rationale (PerformTaskAndPaste)
  const p = spec.placeholders;
  if (!doc.insertRationaleAt(p.rationale, tables.rationale)) notes.push(`"${p.rationale}" was not found, so the investment rationale was not added.`);
  if (!doc.insertTableAt(p.portfolio, { ...tables.portfolio, ...spec.layout.portfolio })) notes.push(`"${p.portfolio}" was not found, so the portfolio table was not added.`);
  if (tables.costs) {
    if (!doc.insertTableAt(p.costs, { ...tables.costs, ...spec.layout.costs })) notes.push(`"${p.costs}" was not found, so the costs table was not added.`);
  } else {
    for (const b of ['Cost1', 'Cost2', 'Cost3']) doc.deleteBookmark(b);
    notes.push('No holding has a cost line, so the cost section (bookmarks Cost1–Cost3) was removed.');
  }
  if (!doc.insertTableAt(p.summary, { ...tables.summary, ...spec.layout.summary })) notes.push(`"${p.summary}" was not found, so the summary table was not added.`);

  return { bytes: await doc.save(), fileName: `${spec.fileName(client, ddmmyyyy(today)).trim()}.docx`, notes };
}
