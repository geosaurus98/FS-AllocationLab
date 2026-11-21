import { useMemo, useState } from 'react';
import { Button, Card, Field, inputCls, Notice, cx } from '../components/ui';
import { adviceTables, type DocTable } from '../engine/workbook/documentTables';
import type { Workbook } from '../engine/workbook/evaluator';
import { DocxDocument } from '../lib/docx';
import { ADVISER_BOOKMARKS, buildLetter, CLIENT_TYPES, LETTER_KINDS, type LetterClient, type LetterKind } from '../lib/letters';

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

/**
 * Documents: the Letter of Advice / Email Advice / Follow-up letter, filled from the firm's Word
 * template with this recommendation's tables. Client details are typed here (the macros read them
 * from the Master Portfolio Risk Profiles workbook). The template stays in memory; nothing is uploaded.
 */
export function Documents({
  workbook,
  custodian,
  client,
  onClient,
  version,
}: {
  workbook: Workbook;
  custodian: 'Apex' | 'OA';
  client: LetterClient;
  onClient: (c: LetterClient) => void;
  /** Changes whenever the workbook is recalculated, so the tables are rebuilt. */
  version: unknown;
}) {
  const [kind, setKind] = useState<LetterKind>('loa');
  const [template, setTemplate] = useState<{ name: string; bytes: ArrayBuffer; bookmarks: string[] } | null>(null);
  const [templateError, setTemplateError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [outcome, setOutcome] = useState<{ fileName: string; notes: string[] } | null>(null);
  const [error, setError] = useState<string | null>(null);

  const tables = useMemo(() => adviceTables(workbook, { clientName: client.fullName || 'Client', custodian }), [workbook, custodian, client.fullName, version]);
  const set = (patch: Partial<LetterClient>) => onClient({ ...client, ...patch });

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
      const out = await buildLetter({ kind, template: template?.bytes ?? null, tables, client: { ...client, fileName: client.fileName || client.fullName || 'Client' }, today: new Date() });
      const url = URL.createObjectURL(new Blob([out.bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = out.fileName;
      a.click();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
      setOutcome({ fileName: out.fileName, notes: out.notes });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const adviserOptions = Array.from(new Set([...ADVISER_BOOKMARKS, ...(template?.bookmarks.filter((b) => !/^(AdvisorTable|Cost\d|BKMRK)$/i.test(b)) ?? [])]));
  const year = new Date().getFullYear();
  const usesAddress = kind === 'loa';

  return (
    <div className="space-y-4">
      <Notice tone="warning">Draft documents for adviser review. Check every figure, the wording and the signatures before anything is sent to a client.</Notice>
      <div className="grid gap-4 lg:grid-cols-2">
        <Card title="Document">
          <div className="space-y-3">
            <Field label="Document type">
              <select className={inputCls} value={kind} onChange={(e) => { setKind(e.target.value as LetterKind); setOutcome(null); }}>
                {(Object.keys(LETTER_KINDS) as LetterKind[]).map((k) => (
                  <option key={k} value={k}>{LETTER_KINDS[k].label}</option>
                ))}
              </select>
            </Field>
            <Field
              label={kind === 'followup' ? "Client's follow-up letter (.docx)" : 'Word template (.docx)'}
              hint={
                kind === 'followup'
                  ? `The macro opens the client's latest letter in Letters ${year}\\Follow-up Letters ${year}.`
                  : `P:\\DATA\\Current (${year})\\Templates ${year}\\Macro Word Templates\\${kind === 'loa' ? 'Letter of Advice Template.docx' : 'Email Advice Template.docx'}. Without a template you get a plain document with the tables.`
              }
            >
              <input type="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" className="block w-full text-sm" onChange={(e) => { const f = e.target.files?.[0]; if (f) void loadTemplate(f); e.target.value = ''; }} />
            </Field>
            {template && (
              <div className="flex items-center justify-between rounded-md border border-slate-200 bg-slate-50 px-3 py-2 text-sm">
                <span>{template.name}{template.bookmarks.length > 0 && <span className="text-slate-500"> · bookmarks: {template.bookmarks.join(', ')}</span>}</span>
                <Button onClick={() => setTemplate(null)}>Remove</Button>
              </div>
            )}
            {templateError && <Notice tone="danger">{templateError}</Notice>}
            <Button variant="primary" onClick={() => void create()} disabled={busy}>
              {busy ? 'Creating…' : template ? `Create ${LETTER_KINDS[kind].label}` : 'Create plain document with the tables'}
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
            <Field label="First name(s)"><input className={inputCls} value={client.firstName} onChange={(e) => set({ firstName: e.target.value })} placeholder="John and Anne" /></Field>
            <Field label="Client type">
              <select className={inputCls} value={client.clientType} onChange={(e) => set({ clientType: e.target.value as LetterClient['clientType'] })}>
                {CLIENT_TYPES.map((t) => <option key={t}>{t}</option>)}
              </select>
            </Field>
            <Field label="Name for the file" hint="ExcelssName, e.g. the surname."><input className={inputCls} value={client.fileName} onChange={(e) => set({ fileName: e.target.value })} /></Field>
            {usesAddress && (
              <>
                <Field label="Street"><input className={inputCls} value={client.street} onChange={(e) => set({ street: e.target.value })} /></Field>
                <Field label="Suburb"><input className={inputCls} value={client.suburb} onChange={(e) => set({ suburb: e.target.value })} /></Field>
                <Field label="City and postcode"><input className={inputCls} value={client.city} onChange={(e) => set({ city: e.target.value })} /></Field>
                <div />
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
                  className={cx('px-2 py-0.5', c.bold && 'font-bold', c.align === 'center' ? 'text-center' : 'text-left', c.topBorder && 'border-t border-slate-500')}
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
