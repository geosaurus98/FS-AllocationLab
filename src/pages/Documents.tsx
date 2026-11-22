import { useMemo, useState } from 'react';
import { Button, Card, Field, inputCls, Notice, cx } from '../components/ui';
import { adviceTables, soaTable, type DocTable } from '../engine/workbook/documentTables';
import type { Workbook } from '../engine/workbook/evaluator';
import { BLANK } from '../engine/workbook/values';
import { buildAnnualReview, buildStatementOfAdvice, emptyReviewDetails, type ReviewDetails } from '../lib/annualReview';
import { DocxDocument } from '../lib/docx';
import { ADVISER_BOOKMARKS, buildLetter, CLIENT_TYPES, type LetterClient, type LetterKind } from '../lib/letters';
import { ALLOCATION_CLASSES, parsePerformanceReport, pct2, type AllocationClass } from '../lib/performanceReport';

export { emptyReviewDetails };
export const emptyLetterClient = (): LetterClient => ({
  fullName: '',
  firstName: '',
  street: '',
  suburb: '',
  city: '',
  clientType: 'Individual',
  fileName: '',
  leadAdviser: '',
  secondaryAdviser: '',
});

type DocKind = LetterKind | 'annual' | 'soa';
const KIND_LABEL: Record<DocKind, string> = { loa: 'Letter of Advice', email: 'Email Advice', followup: 'Follow-up letter', annual: 'Annual Review', soa: 'Statement of Advice appendices' };
const TEMPLATE_FILE: Partial<Record<DocKind, string>> = { loa: 'Letter of Advice Template.docx', email: 'Email Advice Template.docx', annual: 'Annual Review Template.docx', soa: 'Appendices Template.docx' };

/** Portfolio Cash_Target … Multi_Sector_Target (percent). */
const TARGET_CELLS: Record<AllocationClass, string> = { Cash: 'D11', 'Fixed Interest': 'D72', 'Australasian Equities': 'D120', 'Global Equities': 'D168', Property: 'D192', 'Multi Sector': 'D201' };

function download(bytes: Uint8Array, name: string) {
  const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/**
 * Documents: letters, the annual review and the Statement of Advice appendices, filled from the firm's
 * Word templates with this recommendation's tables. Client details are typed here (the macros read them
 * from the Master Portfolio Risk Profiles workbook). Templates and reports stay in memory; nothing is uploaded.
 */
export function Documents({
  workbook,
  custodian,
  client,
  onClient,
  review,
  onReview,
  version,
}: {
  workbook: Workbook;
  custodian: 'Apex' | 'OA';
  client: LetterClient;
  onClient: (c: LetterClient) => void;
  review: ReviewDetails;
  onReview: (r: ReviewDetails) => void;
  /** Changes whenever the workbook is recalculated, so the tables are rebuilt. */
  version: unknown;
}) {
  const [kind, setKind] = useState<DocKind>('loa');
  const [template, setTemplate] = useState<{ name: string; bytes: ArrayBuffer; bookmarks: string[] } | null>(null);
  const [templateError, setTemplateError] = useState<string | null>(null);
  const [reportText, setReportText] = useState('');
  const [commentary, setCommentary] = useState<{ name: string; bytes: ArrayBuffer } | null>(null);
  const [busy, setBusy] = useState(false);
  const [outcome, setOutcome] = useState<{ fileName: string; notes: string[] } | null>(null);
  const [error, setError] = useState<string | null>(null);

  const tables = useMemo(() => adviceTables(workbook, { clientName: client.fullName || 'Client', custodian }), [workbook, custodian, client.fullName, version]);
  const soa = useMemo(() => (kind === 'soa' ? soaTable(workbook, custodian) : null), [kind, workbook, custodian, version]);
  const report = useMemo(() => (reportText.trim() ? parsePerformanceReport(reportText) : null), [reportText]);
  const targets = useMemo(() => {
    const out = {} as Record<AllocationClass, number | null>;
    for (const c of ALLOCATION_CLASSES) {
      const v = workbook.get('Portfolio', TARGET_CELLS[c]);
      out[c] = typeof v === 'number' ? v : null;
    }
    return out;
  }, [workbook, version]);
  const workbookStrategy = useMemo(() => {
    const v = workbook.get('Portfolio', 'BE1');
    return v === BLANK ? '' : String(v);
  }, [workbook, version]);
  const strategy = review.strategyName || workbookStrategy;

  const set = (patch: Partial<LetterClient>) => onClient({ ...client, ...patch });
  const setR = (patch: Partial<ReviewDetails>) => onReview({ ...review, ...patch });

  const loadTemplate = async (f: File) => {
    setTemplateError(null);
    setOutcome(null);
    try {
      const bytes = await f.arrayBuffer();
      const doc = await DocxDocument.open(bytes);
      setTemplate({ name: f.name, bytes, bookmarks: doc.bookmarkNames() });
    } catch (e) {
      setTemplate(null);
      setTemplateError(`${f.name}: ${e instanceof Error ? e.message : String(e)}`);
    }
  };

  const create = async () => {
    setBusy(true);
    setError(null);
    setOutcome(null);
    try {
      const today = new Date();
      const c = { ...client, fileName: client.fileName || client.fullName || 'Client' };
      let out: { bytes: Uint8Array; fileName: string; notes: string[] };
      if (kind === 'annual') {
        if (!template) throw new Error('Choose the Annual Review template first.');
        out = await buildAnnualReview({ template: template.bytes, report, targets, tables, client: c, review: { ...review, strategyName: strategy }, commentary: commentary?.bytes ?? null, today });
      } else if (kind === 'soa') {
        out = await buildStatementOfAdvice({ template: template?.bytes ?? null, table: soa!, strategyName: strategy, clientName: client.fullName, today });
      } else {
        out = await buildLetter({ kind, template: template?.bytes ?? null, tables, client: c, today });
      }
      download(out.bytes, out.fileName);
      setOutcome({ fileName: out.fileName, notes: out.notes });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const adviserOptions = Array.from(new Set([...ADVISER_BOOKMARKS, ...(template?.bookmarks.filter((b) => !/^(AdvisorTable|Cost\d|BKMRK|Table\d|Sig\d|OpeningPV|ClosingPV|EndDate|AlloDate|OpeningComments|AuthorityToAct|SOA_Table)$/i.test(b)) ?? [])]));
  const year = new Date().getFullYear();
  const usesAddress = kind === 'loa';
  const usesAdvisers = kind === 'loa' || kind === 'annual';
  const needsTemplate = kind === 'annual';

  return (
    <div className="space-y-4">
      <Notice tone="warning">Draft documents for adviser review. Check every figure, the wording and the signatures before anything is sent to a client.</Notice>
      <div className="grid gap-4 lg:grid-cols-2">
        <Card title="Document">
          <div className="space-y-3">
            <Field label="Document type">
              <select className={inputCls} value={kind} onChange={(e) => { setKind(e.target.value as DocKind); setOutcome(null); setTemplate(null); }}>
                {(Object.keys(KIND_LABEL) as DocKind[]).map((k) => (
                  <option key={k} value={k}>{KIND_LABEL[k]}</option>
                ))}
              </select>
            </Field>
            <Field
              label={kind === 'followup' ? "Client's follow-up letter (.docx)" : 'Word template (.docx)'}
              hint={
                kind === 'followup'
                  ? `The macro opens the client's latest letter in Letters ${year}\\Follow-up Letters ${year}.`
                  : `P:\\DATA\\Current (${year})\\Templates ${year}\\Macro Word Templates\\${TEMPLATE_FILE[kind]}.${needsTemplate ? '' : ' Without a template you get a plain document with the tables.'}`
              }
            >
              <input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" className="block w-full text-sm" onChange={(e) => { const f = e.target.files?.[0]; if (f) void loadTemplate(f); e.target.value = ''; }} />
            </Field>
            {template && (
              <div className="flex items-center justify-between gap-2 rounded-md border border-slate-200 bg-slate-50 px-3 py-2 text-sm">
                <span>{template.name}{template.bookmarks.length > 0 && <span className="text-slate-500"> · bookmarks: {template.bookmarks.join(', ')}</span>}</span>
                <Button onClick={() => setTemplate(null)}>Remove</Button>
              </div>
            )}
            {templateError && <Notice tone="danger">{templateError}</Notice>}
            <Button variant="primary" onClick={() => void create()} disabled={busy || (needsTemplate && !template)}>
              {busy ? 'Creating…' : template ? `Create ${KIND_LABEL[kind]}` : needsTemplate ? 'Choose the template first' : 'Create plain document with the tables'}
            </Button>
            {error && <Notice tone="danger">{error}</Notice>}
            {outcome && (
              <Notice tone={outcome.notes.length ? 'warning' : 'success'}>
                Downloaded <strong>{outcome.fileName}</strong>.
                {outcome.notes.length > 0 && (
                  <ul className="mt-1 list-disc pl-5">
                    {outcome.notes.map((n) => <li key={n}>{n}</li>)}
                  </ul>
                )}
              </Notice>
            )}
          </div>
        </Card>
        <Card title="Client details" actions={<span className="text-xs text-slate-500">From the Master Portfolio Risk Profiles workbook</span>}>
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Correspondence name"><input className={inputCls} value={client.fullName} onChange={(e) => set({ fullName: e.target.value })} placeholder="Mr J and Mrs A Smith" /></Field>
            {kind !== 'soa' && kind !== 'annual' && <Field label="First name(s)"><input className={inputCls} value={client.firstName} onChange={(e) => set({ firstName: e.target.value })} placeholder="John and Anne" /></Field>}
            {kind !== 'soa' && (
              <Field label="Client type">
                <select className={inputCls} value={client.clientType} onChange={(e) => set({ clientType: e.target.value as LetterClient['clientType'] })}>
                  {CLIENT_TYPES.map((t) => <option key={t}>{t}</option>)}
                </select>
              </Field>
            )}
            {kind !== 'soa' && <Field label="Name for the file" hint="ExcelssName, e.g. the surname."><input className={inputCls} value={client.fileName} onChange={(e) => set({ fileName: e.target.value })} /></Field>}
            {(kind === 'soa' || kind === 'annual') && (
              <Field label="Strategy name" hint={workbookStrategy ? `Blank uses the workbook's Risk_Profile: ${workbookStrategy}.` : 'Risk_Profile in the master file.'}>
                <input className={inputCls} value={review.strategyName} placeholder={workbookStrategy} onChange={(e) => setR({ strategyName: e.target.value })} />
              </Field>
            )}
            {usesAddress && (
              <>
                <Field label="Street"><input className={inputCls} value={client.street} onChange={(e) => set({ street: e.target.value })} /></Field>
                <Field label="Suburb"><input className={inputCls} value={client.suburb} onChange={(e) => set({ suburb: e.target.value })} /></Field>
                <Field label="City and postcode"><input className={inputCls} value={client.city} onChange={(e) => set({ city: e.target.value })} /></Field>
                <div />
              </>
            )}
            {usesAdvisers && (
              <>
                <Field label="Lead adviser" hint="Signature bookmark in the template.">
                  <input className={inputCls} list="adviser-bookmarks" value={client.leadAdviser} onChange={(e) => set({ leadAdviser: e.target.value })} />
                </Field>
                <Field label="Secondary adviser">
                  <input className={inputCls} list="adviser-bookmarks" value={client.secondaryAdviser} onChange={(e) => set({ secondaryAdviser: e.target.value })} />
                </Field>
                <datalist id="adviser-bookmarks">{adviserOptions.map((b) => <option key={b} value={b} />)}</datalist>
              </>
            )}
          </div>
        </Card>
      </div>

      {kind === 'annual' && (
        <>
          <div className="grid gap-4 lg:grid-cols-2">
            <Card title="Review details" actions={<span className="text-xs text-slate-500">From the master file</span>}>
              <div className="grid gap-3 sm:grid-cols-2">
                <Field label="Signatories" hint="Comma separated, up to four (Sig1–Sig4)."><input className={inputCls} value={review.signatories} onChange={(e) => setR({ signatories: e.target.value })} /></Field>
                <Field label="I or We">
                  <select className={inputCls} value={review.iOrWe} onChange={(e) => setR({ iOrWe: e.target.value as 'I' | 'We' })}>
                    <option>I</option>
                    <option>We</option>
                  </select>
                </Field>
                <Field label="RWT rate" hint="Blank leaves XXX highlighted."><input className={inputCls} value={review.rwt} onChange={(e) => setR({ rwt: e.target.value })} /></Field>
                <Field label="PIR rate" hint="Blank leaves XXX highlighted."><input className={inputCls} value={review.pir} onChange={(e) => setR({ pir: e.target.value })} /></Field>
                <Field label="Accountant" hint="Blank leaves “accountant” highlighted."><input className={inputCls} value={review.accountant} onChange={(e) => setR({ accountant: e.target.value })} /></Field>
                <label className="flex items-center gap-2 self-end pb-2 text-sm"><input type="checkbox" checked={review.authorityToAct} onChange={(e) => setR({ authorityToAct: e.target.checked })} /> Authority to act given (removes that section)</label>
                <div className="sm:col-span-2"><Field label="Primary objective"><input className={inputCls} value={review.primaryObjective} onChange={(e) => setR({ primaryObjective: e.target.value })} /></Field></div>
                <div className="sm:col-span-2"><Field label="Previous blurb" hint="Replaces “Insert previous Blurb”."><textarea className={cx(inputCls, 'h-20')} value={review.blurb} onChange={(e) => setR({ blurb: e.target.value })} /></Field></div>
                <div className="sm:col-span-2">
                  <Field label="Commentary insert (.docx)" hint="P:\DATA\Current (year)\Commentary for Portfolio year\Annual Review Commentary year\“Month Year insert.docx”, placed at the OpeningComments bookmark. Pictures are not copied.">
                    <input type="file" accept=".docx" className="block w-full text-sm" onChange={async (e) => { const f = e.target.files?.[0]; if (f) setCommentary({ name: f.name, bytes: await f.arrayBuffer() }); e.target.value = ''; }} />
                  </Field>
                  {commentary && <p className="mt-1 text-xs text-slate-600">{commentary.name} <button type="button" className="ml-2 underline" onClick={() => setCommentary(null)}>remove</button></p>}
                </div>
              </div>
            </Card>
            <Card title="Performance report" actions={<span className="text-xs text-slate-500">{report ? (report.custodian === 'OA' ? 'OneAnswer report' : 'Apex report') : 'Paste the report text'}</span>}>
              <p className="mb-2 text-xs text-slate-600">Paste the custodian performance report as text, one line per row, as it is pasted into the workbook's Input sheet. OneAnswer reports start with “Printed”.</p>
              <textarea aria-label="Performance report" className={cx(inputCls, 'h-40 font-mono text-xs')} value={reportText} onChange={(e) => setReportText(e.target.value)} />
              {report && <ReportSummary report={report} targets={targets} />}
            </Card>
          </div>
        </>
      )}

      {kind === 'soa' && soa && (
        <Card title="Appendix product table" actions={<span className="text-xs text-slate-500">“SOA_Table” · landscape section</span>}>
          <TablePreview table={soa} />
        </Card>
      )}

      {kind !== 'soa' && (
        <>
          <Card title="Summary of changes" actions={<span className="text-xs text-slate-500">“ChangeTable”</span>}>
            <TablePreview table={tables.summary} />
          </Card>
          <Card title="Portfolio for advice" actions={<span className="text-xs text-slate-500">“Portfolioforadvice”</span>}>
            <TablePreview table={tables.portfolio} />
          </Card>
          <Card title="Costs for advice" actions={<span className="text-xs text-slate-500">“InsertTable”</span>}>
            {tables.costs ? <TablePreview table={tables.costs} /> : <p className="text-sm text-slate-500">No holding has a cost line; the Cost1–Cost3 sections are removed from the letter.</p>}
          </Card>
          <Card title="Investment rationale" actions={<span className="text-xs text-slate-500">“Investmentrationale”</span>}>
            <div className="space-y-1 text-sm">
              {tables.rationale.map((r, i) =>
                r.kind === 'bullet' ? (
                  <p key={i} className="ml-8 list-item list-disc">
                    {r.text}{' '}
                    {r.factSheet && <>Link to the fact sheet <a className="text-sky-700 underline" href={r.factSheet} target="_blank" rel="noreferrer noopener">here</a>.</>}
                  </p>
                ) : (
                  <p key={i} className={cx('ml-3 font-semibold', r.kind === 'heading' && 'mt-2 underline')}>{r.text}</p>
                ),
              )}
            </div>
          </Card>
        </>
      )}
    </div>
  );
}

function ReportSummary({ report, targets }: { report: ReturnType<typeof parsePerformanceReport>; targets: Record<AllocationClass, number | null> }) {
  const LABELS = ['Opening value', 'Contributions', 'Withdrawals', report.custodian === 'OA' ? 'Income and fees' : 'Row 4', report.custodian === 'OA' ? 'Subtotal' : 'Row 5', 'Closing value', 'Gain ($)', 'Return (%)'];
  return (
    <div className="mt-3 space-y-3 text-xs">
      {report.missing.length > 0 && <Notice tone="warning">Not found in the report: {report.missing.join('; ')}.</Notice>}
      <p>
        Period: <strong>{report.openingDate ?? '—'}</strong> to <strong>{report.closingDate ?? '—'}</strong>
        {report.inception && <> · since inception ({report.inception.days} days): arithmetic {pct2(report.inception.arithmetic)}, geometric {pct2(report.inception.geometric)}{report.inception.days < 1825 && ' (under 5 years: rows removed)'}</>}
      </p>
      <div className="grid gap-3 sm:grid-cols-3">
        <table className="w-full">
          <caption className="text-left font-semibold">Table1</caption>
          <tbody>
            {report.table1.map((r, i) => (
              <tr key={i}><td className="pr-2 text-slate-500">{LABELS[i]}</td><td className={cx('text-right tabular-nums', r.negative && 'text-red-700')}>{r.text}</td></tr>
            ))}
          </tbody>
        </table>
        <table className="w-full">
          <caption className="text-left font-semibold">Table2</caption>
          <tbody>
            {report.sectors.map((s) => (
              <tr key={s.heading}><td className="pr-2 text-slate-500">{s.heading}</td><td className="text-right tabular-nums">{s.percent ?? '—'}</td><td className="text-right tabular-nums">{s.dollars ?? '—'}</td></tr>
            ))}
          </tbody>
        </table>
        <table className="w-full">
          <caption className="text-left font-semibold">Table3 (target / current)</caption>
          <tbody>
            {ALLOCATION_CLASSES.map((c) => (
              <tr key={c}><td className="pr-2 text-slate-500">{c}</td><td className="text-right tabular-nums">{targets[c] === null ? '—' : pct2(targets[c]! / 100)}</td><td className="text-right tabular-nums">{report.current[c] === null ? '—' : pct2(report.current[c]!)}</td></tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function TablePreview({ table }: { table: DocTable }) {
  return (
    <div className="overflow-x-auto">
      <table className="border-collapse text-xs" style={{ fontFamily: 'Verdana, sans-serif' }}>
        <tbody>
          {table.rows.map((row, i) => (
            <tr key={i}>
              {row.map((c, j) => (
                <td
                  key={j}
                  className={cx('px-2 py-0.5', c.bold && 'font-bold', c.align === 'center' ? 'text-center' : 'text-left', c.topBorder && 'border-t border-slate-500', table.columnLines && 'border border-slate-300')}
                  style={{ background: c.fill ? `#${c.fill}` : undefined, color: c.color ? `#${c.color}` : undefined, minWidth: (table.widths[j] ?? 60) * 0.9 }}
                >
                  {c.link ? <a className="text-sky-700 underline" href={c.link} target="_blank" rel="noreferrer noopener">{c.text}</a> : c.text}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
