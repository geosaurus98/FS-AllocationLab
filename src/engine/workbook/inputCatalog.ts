import type { WorkbookModel } from './evaluator';

/**
 * Asset Allocation input cells that the Settings page exposes (beyond the profile table and the
 * property model, which have their own editors). Defaults come from the workbook itself.
 */

export type InputKind = 'pct' | 'money' | 'number' | 'text' | 'yesno';

export interface InputField {
  cell: string;
  label: string;
  kind: InputKind;
  hint?: string;
  options?: string[];
}

export interface InputGroup {
  title: string;
  fields: InputField[];
  /** Model fund lists: name cell column + distribution column for a row range. */
  lists?: { title: string; nameCol: string; distCol: string | null; extraCol?: { col: string; label: string }; rows: [number, number] }[];
}

export const INPUT_GROUPS: InputGroup[] = [
  {
    title: 'Fixed interest',
    fields: [
      { cell: 'C15', label: 'Managed funds share of fixed interest', kind: 'pct' },
      { cell: 'J15', label: 'Bond rounding', kind: 'money' },
      { cell: 'M15', label: 'Minimum change for funds', kind: 'money', hint: "Don't change if less than this amount" },
      { cell: 'M16', label: 'Maximum exposure to one bank', kind: 'pct' },
      { cell: 'M17', label: 'Reduce bonds?', kind: 'yesno', options: ['Yes', 'No'] },
      { cell: 'M28', label: 'Minimum new bond', kind: 'money' },
      { cell: 'M29', label: 'Minimum new term deposit (over 1 year)', kind: 'money' },
      { cell: 'M30', label: 'Minimum new term deposit (under 1 year)', kind: 'money' },
      { cell: 'I29', label: 'Add small additionals to', kind: 'text' },
    ],
    lists: [{ title: 'Fixed-interest managed funds', nameCol: 'P', distCol: null, rows: [15, 29] }],
  },
  {
    title: 'Australasian equities',
    fields: [
      { cell: 'D35', label: 'NZ share of Australasian equities', kind: 'pct' },
      { cell: 'D36', label: 'Australian share', kind: 'pct' },
      { cell: 'D37', label: 'Trans-Tasman share', kind: 'pct' },
      { cell: 'J35', label: 'Equal-weight AUS and NZ direct shares', kind: 'yesno', options: ['Yes', 'No'] },
      { cell: 'J44', label: 'NZ direct shares distribution', kind: 'pct' },
      { cell: 'J55', label: 'AUS direct shares distribution', kind: 'pct' },
      { cell: 'M35', label: 'Fund trade minimum', kind: 'money' },
      { cell: 'M36', label: 'Share trade minimum', kind: 'money' },
      { cell: 'M37', label: 'Minimum number of direct shares', kind: 'number' },
      { cell: 'M38', label: 'Minimum amount per direct share', kind: 'money' },
    ],
    lists: [
      { title: 'NZ funds', nameCol: 'B', distCol: 'C', rows: [41, 47] },
      { title: 'Australian funds', nameCol: 'B', distCol: 'C', rows: [52, 61] },
      { title: 'Trans-Tasman funds', nameCol: 'B', distCol: 'C', rows: [66, 70] },
      { title: 'Direct shares (AUS and NZ) to add', nameCol: 'P', distCol: null, rows: [35, 70] },
    ],
  },
  {
    title: 'Global equities',
    fields: [
      { cell: 'C74', label: 'Gold (whole portfolio)', kind: 'pct' },
      { cell: 'C75', label: 'Passive share', kind: 'pct' },
      { cell: 'C76', label: 'Active share', kind: 'pct' },
      { cell: 'C77', label: 'Prioritise', kind: 'yesno', options: ['Hedged', 'Active/Passive'] },
      { cell: 'G74', label: "Active group 1 ('1's) share", kind: 'pct' },
      { cell: 'G75', label: "Active group 2 ('2's) share", kind: 'pct' },
      { cell: 'G76', label: "Active group 3 ('3's) share", kind: 'pct' },
      { cell: 'C101', label: 'Direct shares share', kind: 'pct' },
      { cell: 'J101', label: 'Total hedged', kind: 'pct' },
      { cell: 'M101', label: 'Hedged prioritised within passive', kind: 'yesno', options: ['Yes', 'No'] },
      { cell: 'Q73', label: 'Gold limit', kind: 'money' },
      { cell: 'Q74', label: 'Let gold run', kind: 'yesno', options: ['Yes', 'No'] },
      { cell: 'Q75', label: 'Fund trade minimum', kind: 'money' },
      { cell: 'Q76', label: 'Share trade minimum', kind: 'money' },
    ],
    lists: [
      { title: 'Active (group number in place of %)', nameCol: 'B', distCol: null, extraCol: { col: 'C', label: 'Group' }, rows: [80, 99] },
      { title: 'Passive', nameCol: 'B', distCol: 'C', rows: [105, 114] },
      { title: 'Gold assets (distribution set from holdings at import)', nameCol: 'J', distCol: null, rows: [74, 77] },
    ],
  },
];

export type InputValue = number | string | null;

/** Literal value of an Asset Allocation cell in the workbook model (null if empty or a formula). */
export function workbookInputDefault(model: WorkbookModel, a1: string): InputValue {
  const m = /^([A-Z]+)(\d+)$/.exec(a1);
  if (!m) return null;
  const col = m[1]!.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const row = Number(m[2]);
  const cell = model.sheets['Asset Allocation']!.cells.find(([r, c]) => r === row && c === col);
  if (!cell) return null;
  const [, , k, p] = cell;
  if (k === 'n') return p as number;
  if (k === 's') return p as string;
  return null;
}

export function isFormulaCell(model: WorkbookModel, a1: string): boolean {
  const m = /^([A-Z]+)(\d+)$/.exec(a1)!;
  const col = m[1]!.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const row = Number(m[2]);
  const cell = model.sheets['Asset Allocation']!.cells.find(([r, c]) => r === row && c === col);
  return !!cell && (cell[2] === 'f' || cell[2] === 'a');
}
