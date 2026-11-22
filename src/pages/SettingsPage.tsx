import { useEffect, useMemo, useState } from 'react';
import type { WorkbookModel } from '../engine/workbook/evaluator';
import { INPUT_GROUPS, isFormulaCell, workbookInputDefault, type InputField, type InputValue } from '../engine/workbook/inputCatalog';
import { PercentInput } from '../components/PercentInput';
import { Button, Card, Field, inputCls, IssueList, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, cx } from '../components/ui';
import { isErr, type XValue } from '../engine/excel';
import { profileColumn, WORKBOOK_PROFILES, type ProfileInputs, type WorkbookProfile } from '../engine/profiles';
import { WORKBOOK_SETTINGS, type FirmSettings, type ModelHolding } from '../engine/settings';
import { validateFirmSettings } from '../engine/settingsValidation';
import { formatDate } from '../lib/format';
import { useEngineSettings } from '../state/EngineSettingsContext';

type RowKey = 'lowerRisk' | 'cash' | 'fi' | 'growth' | 'ae' | 'ge' | 'prop' | 'ms';
const ROWS: { key: RowKey; label: string }[] = [
  { key: 'lowerRisk', label: 'Lower risk (defensive)' },
  { key: 'cash', label: 'Cash' },
  { key: 'fi', label: 'Fixed income' },
  { key: 'growth', label: 'Growth' },
  { key: 'ae', label: 'Australasian equities' },
  { key: 'ge', label: 'Global equities' },
  { key: 'prop', label: 'Property' },
  { key: 'ms', label: 'Multi sector' },
];
type MasterKey = 'cash' | 'fi' | 'ae' | 'ge' | 'prop' | 'ms';
const EDIT_PROFILES = ['Conservative', 'Moderate', 'Balanced', 'Growth', 'Aggressive', 'Custom'] as const;

const clone = <T,>(v: T): T => JSON.parse(JSON.stringify(v)) as T;
const show = (v: XValue) => (isErr(v) ? v.code : `${(v * 100).toFixed(2)}%`);
const toPct = (v: number | null | undefined) => (v === null || v === undefined ? null : Math.round(v * 1e8) / 1e6);

export function SettingsPage() {
  const es = useEngineSettings();
  const [draft, setDraft] = useState<FirmSettings>(() => clone(es.current.settings));
  const [baseVersion, setBaseVersion] = useState(es.current.version);
  const [note, setNote] = useState('');
  const [status, setStatus] = useState<FirmSettings['status']>('approved');
  const [confirmReset, setConfirmReset] = useState(false);
  if (baseVersion !== es.current.version) {
    setBaseVersion(es.current.version);
    setDraft(clone(es.current.settings));
  }
  const issues = useMemo(() => validateFirmSettings(draft), [draft]);
  const errors = issues.filter((i) => i.severity === 'error');
  const dirty = JSON.stringify(draft) !== JSON.stringify(es.current.settings);
  const differsFromWorkbook = JSON.stringify(es.current.settings.tactical) !== JSON.stringify(WORKBOOK_SETTINGS.tactical) || JSON.stringify(es.current.settings.property) !== JSON.stringify(WORKBOOK_SETTINGS.property) || JSON.stringify(es.current.settings.general) !== JSON.stringify(WORKBOOK_SETTINGS.general);

  const set = (f: (d: FirmSettings) => void) => setDraft((d) => {
    const n = clone(d);
    f(n);
    return n;
  });

  return (
    <>
      <PageHeader
        title="Settings"
        subtitle="Firm-level inputs to the recommendations engine — the equivalent of the workbook's Asset Allocation sheet. Changes apply to recommendations once saved; every save is a new version."
        actions={
          <>
            <Button onClick={() => setDraft(clone(es.current.settings))} disabled={!dirty}>Discard changes</Button>
            {confirmReset ? (
              <>
                <Button variant="danger" onClick={() => { es.resetToWorkbook('Reset to workbook values'); setConfirmReset(false); }}>Confirm reset</Button>
                <Button onClick={() => setConfirmReset(false)}>Cancel</Button>
              </>
            ) : (
              <Button variant="danger" onClick={() => setConfirmReset(true)} disabled={!differsFromWorkbook}>Reset to workbook values…</Button>
            )}
          </>
        }
      />

      <div className="mb-6 flex flex-wrap items-center gap-3 text-sm">
        <StatusBadge status={es.current.status === 'approved' ? 'approved' : 'draft'} />
        <span>Version {es.current.version} · saved {formatDate(es.current.savedAt)}</span>
        <span className="text-slate-500">{es.current.note}</span>
        {!differsFromWorkbook && <span className="rounded bg-slate-100 px-1.5 py-0.5 text-xs text-slate-600">Same as workbook</span>}
        {dirty && <span className="rounded bg-amber-100 px-1.5 py-0.5 text-xs font-medium text-amber-900">Unsaved changes</span>}
      </div>

      <Card title="Risk profiles (tactical asset allocation)" className="mb-6">
        <p className="mb-3 text-xs text-slate-600">
          Percentages of the whole portfolio. Shaded cells are calculated the same way as the workbook: equity and property splits for the other profiles are scaled from <strong>Balanced</strong>. Custom values are the defaults for a new Custom client.
        </p>
        <TableWrap>
          <thead>
            <tr>
              <Th />
              {EDIT_PROFILES.map((p) => <Th key={p} align="right">{p}</Th>)}
              <Th align="right">Master Risk Alloc</Th>
            </tr>
          </thead>
          <tbody>
            {ROWS.map((row) => (
              <tr key={row.key} className={row.key === 'lowerRisk' || row.key === 'growth' ? 'bg-slate-50 font-medium' : ''}>
                <Td className="whitespace-nowrap">{row.label}</Td>
                {EDIT_PROFILES.map((p) => {
                  const inputs = draft.tactical[p] as Partial<ProfileInputs>;
                  const literal = row.key in inputs && inputs[row.key] !== null && inputs[row.key] !== undefined;
                  const computed = profileColumn(draft.tactical, p as WorkbookProfile)[row.key];
                  return (
                    <Td key={p} align="right">
                      {literal ? (
                        <PercentInput
                          className="ml-auto w-24"
                          ariaLabel={`${p} ${row.label}`}
                          value={toPct(inputs[row.key] as number)}
                          onChange={(v) => set((d) => { (d.tactical[p] as Record<string, number | null>)[row.key] = v === null ? 0 : v / 100; })}
                        />
                      ) : (
                        <span className="inline-block w-24 rounded bg-slate-100 px-2 py-1 text-slate-600" title="Calculated">{show(computed)}</span>
                      )}
                    </Td>
                  );
                })}
                <Td align="right">
                  {row.key === 'lowerRisk' || row.key === 'growth' ? (
                    <span className="inline-block w-24 rounded bg-slate-100 px-2 py-1 text-slate-600">{show(profileColumn(draft.tactical, 'Master Risk Alloc')[row.key])}</span>
                  ) : (
                    <PercentInput className="ml-auto w-24" ariaLabel={`Master Risk Alloc ${row.label}`} value={toPct(draft.tactical['Master Risk Alloc'][row.key as MasterKey])} onChange={(v) => set((d) => { d.tactical['Master Risk Alloc'][row.key as MasterKey] = v === null ? 0 : v / 100; })} />
                  )}
                </Td>
              </tr>
            ))}
            <tr>
              <Td className="text-xs text-slate-500">Total</Td>
              {[...EDIT_PROFILES, 'Master Risk Alloc' as const].map((p) => {
                const c = profileColumn(draft.tactical, p);
                const parts = [c.cash, c.fi, c.ae, c.ge, c.prop, c.ms];
                const total = parts.some(isErr) ? null : (parts as number[]).reduce((a, x) => a + x, 0);
                return <Td key={p} align="right" className={cx('text-xs', total !== null && Math.abs(total - 1) > 1e-9 ? 'font-semibold text-red-700' : 'text-slate-500')}>{total === null ? '—' : `${(total * 100).toFixed(2)}%`}</Td>;
              })}
            </tr>
          </tbody>
        </TableWrap>
        <p className="mt-2 text-xs text-slate-500">“Current” is not configurable — it uses each client's own current mix. {WORKBOOK_PROFILES.length} profiles in total.</p>
      </Card>

      <Card title="Property model" className="mb-6">
        <h3 className="mb-2 text-sm font-semibold">Size bands (the property pot decides the split)</h3>
        <TableWrap>
          <thead><tr><Th>Band</Th><Th align="right">Pot over</Th><Th align="right">Up to</Th><Th align="right">NZ funds</Th><Th align="right">NZ shares</Th><Th align="right">AUS</Th><Th align="right">International</Th></tr></thead>
          <tbody>
            {draft.property.bands.map((b, i) => (
              <tr key={i}>
                <Td>{i + 1}</Td>
                <Td align="right">{i === 0 ? '0' : <PercentInput className="ml-auto w-28" suffix="$" ariaLabel={`Band ${i + 1} from`} value={b.from} onChange={(v) => set((d) => { d.property.bands[i]!.from = v ?? 0; })} />}</Td>
                <Td align="right">{b.to === null ? 'no limit' : <PercentInput className="ml-auto w-28" suffix="$" ariaLabel={`Band ${i + 1} to`} value={b.to} onChange={(v) => set((d) => { d.property.bands[i]!.to = v; })} />}</Td>
                {(['nzFunds', 'nzShares', 'aus', 'int'] as const).map((k) => (
                  <Td key={k} align="right"><PercentInput className="ml-auto w-24" ariaLabel={`Band ${i + 1} ${k}`} value={toPct(b[k])} onChange={(v) => set((d) => { d.property.bands[i]![k] = v === null ? null : v / 100; })} /></Td>
                ))}
              </tr>
            ))}
          </tbody>
        </TableWrap>
        <div className="mt-4 grid gap-4 lg:grid-cols-2">
          <ModelList title="NZ property shares" list={draft.property.nzShares} onChange={(l) => set((d) => { d.property.nzShares = l; })} />
          <ModelList title="NZ property funds" list={draft.property.nzFunds} onChange={(l) => set((d) => { d.property.nzFunds = l; })} />
          <ModelList title="Australian property (banded)" list={[draft.property.ausBanded]} onChange={(l) => set((d) => { d.property.ausBanded = l[0]!; })} />
          <ModelList title="International property (banded)" list={[draft.property.intBanded]} onChange={(l) => set((d) => { d.property.intBanded = l[0]!; })} />
        </div>
        <div className="mt-4 grid gap-3 sm:grid-cols-3">
          <Field label="Fund trade minimum"><PercentInput suffix="$" ariaLabel="Property fund limit" value={draft.property.fundLimit} onChange={(v) => set((d) => { d.property.fundLimit = v; })} /></Field>
          <Field label="Share trade minimum"><PercentInput suffix="$" ariaLabel="Property share limit" value={draft.property.shareLimit} onChange={(v) => set((d) => { d.property.shareLimit = v; })} /></Field>
        </div>
      </Card>

      <WorkbookInputsCard draft={draft} set={set} />

      <Card title="General" className="mb-6">
        <div className="grid gap-3 sm:grid-cols-4">
          <Field label="Rounding unit" hint="New purchases are rounded to this."><PercentInput suffix="$" ariaLabel="Portfolio rounding" value={draft.general.portfolioRounding} onChange={(v) => set((d) => { d.general.portfolioRounding = v; })} /></Field>
          <Field label="Minimum change" hint="Don't change an investment unless the change exceeds this % of its value."><PercentInput ariaLabel="Portfolio tolerance" value={toPct(draft.general.portfolioTolerance)} onChange={(v) => set((d) => { d.general.portfolioTolerance = v === null ? null : v / 100; })} /></Field>
          <Field label="Asset class tolerance ±"><PercentInput ariaLabel="Asset class tolerance" value={toPct(draft.general.assetClassTolerance)} onChange={(v) => set((d) => { d.general.assetClassTolerance = v === null ? null : v / 100; })} /></Field>
        </div>
      </Card>

      <Card title="Save" className="mb-6">
        <div className="mb-3"><IssueList issues={issues.map((i) => ({ code: 'NO_DATA' as const, severity: i.severity, message: i.message }))} emptyText="All profile and band checks pass." /></div>
        <div className="grid gap-3 sm:grid-cols-[1fr_200px_auto] sm:items-end">
          <Field label="Change note (required)"><input className={inputCls} value={note} onChange={(e) => setNote(e.target.value)} placeholder="e.g. IC meeting 3 Dec — Growth cash 2% → 3%" /></Field>
          <Field label="Status">
            <select className={inputCls} value={status} onChange={(e) => setStatus(e.target.value as FirmSettings['status'])}>
              <option value="approved">Approved</option>
              <option value="draft">Draft (not approved)</option>
            </select>
          </Field>
          <Button variant="primary" disabled={!dirty || errors.length > 0 || !note.trim()} onClick={() => { es.save(draft, note.trim(), status); setNote(''); }}>Save as version {es.current.version + 1}</Button>
        </div>
        {errors.length > 0 && <p className="mt-2 text-xs text-red-700">Fix the errors above before saving.</p>}
      </Card>

      <Card title="Version history">
        {es.history.length === 0 ? <p className="text-sm text-slate-500">No earlier versions.</p> : (
          <ul className="space-y-1 text-sm">
            {[...es.history].reverse().map((h) => (
              <li key={h.version} className="flex flex-wrap gap-3 border-b border-slate-100 py-1">
                <span className="w-12 font-medium">v{h.version}</span>
                <span className="text-slate-500">{formatDate(h.savedAt)}</span>
                <span>{h.note}</span>
              </li>
            ))}
          </ul>
        )}
        <div className="mt-3"><Notice tone="info">Settings are stored in this browser. Use the same browser profile (or export/import, coming with the recommendations screen) to share them.</Notice></div>
      </Card>
    </>
  );
}

function ModelList({ title, list, onChange }: { title: string; list: ModelHolding[]; onChange: (l: ModelHolding[]) => void }) {
  const upd = (i: number, patch: Partial<ModelHolding>) => onChange(list.map((m, j) => (j === i ? { ...m, ...patch } : m)));
  const used = list.filter((m) => m.name);
  const total = used.reduce((a, m) => a + (m.distribution ?? 0), 0);
  return (
    <div>
      <h3 className="mb-1 text-sm font-semibold">{title}</h3>
      <table className="w-full text-sm">
        <tbody>
          {list.map((m, i) => (
            <tr key={i}>
              <td className="py-0.5 pr-2"><input aria-label={`${title} name ${i + 1}`} className={inputCls} value={m.name} placeholder="(empty)" onChange={(e) => upd(i, { name: e.target.value })} /></td>
              <td className="w-28 py-0.5"><PercentInput ariaLabel={`${title} distribution ${i + 1}`} value={toPct(m.distribution)} onChange={(v) => upd(i, { distribution: v === null ? null : v / 100 })} /></td>
            </tr>
          ))}
        </tbody>
      </table>
      {used.length > 0 && <p className={cx('mt-1 text-xs', Math.abs(total - 1) > 1e-9 ? 'text-amber-800' : 'text-slate-500')}>Distribution total {(total * 100).toFixed(2)}%</p>}
    </div>
  );
}

function WorkbookInputsCard({ draft, set }: { draft: FirmSettings; set: (f: (d: FirmSettings) => void) => void }) {
  const [model, setModel] = useState<WorkbookModel | null>(null);
  useEffect(() => {
    import('../engine/workbook/model.json').then((m) => setModel(m.default as unknown as WorkbookModel)).catch(() => setModel(null));
  }, []);
  if (!model) return <Card title="Fixed interest & equities" className="mb-6"><p className="text-sm text-slate-500">Loading workbook inputs…</p></Card>;
  const value = (cell: string): InputValue => (draft.inputs && cell in draft.inputs ? draft.inputs[cell]! : workbookInputDefault(model, cell));
  const setValue = (cell: string, v: InputValue) =>
    set((d) => {
      const inputs = { ...(d.inputs ?? {}) };
      if (v === workbookInputDefault(model, cell)) delete inputs[cell];
      else inputs[cell] = v;
      d.inputs = inputs;
    });
  const changed = Object.keys(draft.inputs ?? {}).length;
  return (
    <Card title="Fixed interest & equities" className="mb-6" actions={changed ? <Button variant="ghost" onClick={() => set((d) => { d.inputs = {}; })}>Revert {changed} change(s) to workbook values</Button> : undefined}>
      <p className="mb-4 text-xs text-slate-600">Inputs from the workbook's Asset Allocation sheet. Changed values are highlighted; everything else uses the workbook's own value.</p>
      <div className="space-y-6">
        {INPUT_GROUPS.map((g) => (
          <div key={g.title}>
            <h3 className="mb-2 text-sm font-semibold text-slate-800">{g.title}</h3>
            <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
              {g.fields.map((f) => <InputFieldControl key={f.cell} field={f} value={value(f.cell)} changed={!!draft.inputs && f.cell in draft.inputs} onChange={(v) => setValue(f.cell, v)} />)}
            </div>
            {g.lists && (
              <div className="mt-3 grid gap-4 lg:grid-cols-2">
                {g.lists.map((l) => (
                  <details key={l.title} className="rounded border border-slate-200 p-2">
                    <summary className="cursor-pointer text-sm font-medium">{l.title}</summary>
                    <table className="mt-2 w-full text-sm">
                      <tbody>
                        {Array.from({ length: l.rows[1] - l.rows[0] + 1 }, (_, i) => l.rows[0] + i).filter((r) => !isFormulaCell(model, `${l.nameCol}${r}`)).map((r) => {
                          const nameCell = `${l.nameCol}${r}`;
                          const second = l.distCol ?? l.extraCol?.col ?? null;
                          const secondCell = second ? `${second}${r}` : null;
                          const nv = value(nameCell);
                          return (
                            <tr key={r}>
                              <td className="py-0.5 pr-2"><input aria-label={`${l.title} ${r}`} className={cx(inputCls, draft.inputs && nameCell in draft.inputs && 'border-amber-400 bg-amber-50')} value={nv === null ? '' : String(nv)} placeholder="(empty)" onChange={(e) => setValue(nameCell, e.target.value || null)} /></td>
                              {secondCell && !isFormulaCell(model, secondCell) && (
                                <td className="w-28 py-0.5">
                                  {l.distCol ? (
                                    <PercentInput ariaLabel={`${l.title} ${r} distribution`} value={typeof value(secondCell) === 'number' ? toPct(value(secondCell) as number) : null} onChange={(v) => setValue(secondCell, v === null ? null : v / 100)} />
                                  ) : (
                                    <PercentInput suffix="" ariaLabel={`${l.title} ${r} ${l.extraCol?.label}`} value={typeof value(secondCell) === 'number' ? (value(secondCell) as number) : null} onChange={(v) => setValue(secondCell, v)} />
                                  )}
                                </td>
                              )}
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </details>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
    </Card>
  );
}

function InputFieldControl({ field, value, changed, onChange }: { field: InputField; value: InputValue; changed: boolean; onChange: (v: InputValue) => void }) {
  const cls = cx(changed && 'rounded ring-2 ring-amber-300');
  let control;
  if (field.kind === 'yesno') {
    control = (
      <select className={inputCls} value={value === null ? '' : String(value)} onChange={(e) => onChange(e.target.value)}>
        {(field.options ?? ['Yes', 'No']).map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    );
  } else if (field.kind === 'text') {
    control = <input className={inputCls} value={value === null ? '' : String(value)} onChange={(e) => onChange(e.target.value || null)} />;
  } else {
    const pct = field.kind === 'pct';
    control = (
      <PercentInput
        ariaLabel={field.label}
        suffix={field.kind === 'money' ? '$' : pct ? '%' : ''}
        value={typeof value === 'number' ? (pct ? toPct(value) : value) : null}
        onChange={(v) => onChange(v === null ? null : pct ? v / 100 : v)}
      />
    );
  }
  return (
    <div className={cls}>
      <Field label={field.label} hint={field.hint}>{control}</Field>
    </div>
  );
}
