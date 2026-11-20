import { useMemo, useState } from 'react';
import sampleClients from '../data/sample-clients.csv?raw';
import sampleHoldings from '../data/sample-holdings.csv?raw';
import { SplitBar, StackedAllocationChart } from '../components/charts';
import { PercentInput } from '../components/PercentInput';
import { Button, Card, EmptyState, Field, inputCls, IssueList, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, ValidationBadge, cx } from '../components/ui';
import { toCsv } from '../lib/csv';
import { formatPct, formatPp } from '../lib/format';
import {
  aggregateClient,
  allClientNumbers,
  analyseClient,
  classLabels,
  fmtMoney,
  parseClientsCsv,
  parseHoldingsCsv,
  type ClientAnalysis,
  type ParseOutcome,
} from '../lib/holdings';
import { ASSET_CLASSES, profile } from '../lib/reference';
import type { AssetClassId, BaselineKind, ValidationStatus } from '../lib/types';
import { useApp } from '../state/AppContext';
import { useClientData, type LoadedFile } from '../state/ClientDataContext';
import { findBaseline } from '../state/state';

function download(filename: string, text: string) {
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const CLIENT_TEMPLATE = toCsv([
  ['Client number', 'Risk Profile', 'Growth Target %', 'Monthly Fees', 'Monthly Withdrawals', 'Cash Cover Months', 'Drift Tolerance pp', 'Notes'],
  ['EXAMPLE-1', 'Balanced 60/40', 60, 450, 2500, '', '', 'optional columns may be left blank'],
]);
const HOLDINGS_TEMPLATE = toCsv([
  ['Client number', 'Asset Name', 'Asset Value', 'Asset Class', 'Asset Sub-Class'],
  ['EXAMPLE-1', 'Example Cash Fund', '20000.00', 'Cash', 'New Zealand'],
]);

const money = (v: number | null | undefined) => (v === null || v === undefined || !Number.isFinite(v) ? '—' : `$${fmtMoney(v)}`);

function statusOf(a: ClientAnalysis): ValidationStatus {
  if (a.issues.some((i) => i.severity === 'error')) return a.client ? 'invalid' : 'incomplete';
  if (a.issues.some((i) => i.severity === 'warning')) return 'warnings';
  return 'valid';
}

export function ClientHoldings() {
  const { state } = useApp();
  const cd = useClientData();
  const [selected, setSelected] = useState<string | null>(null);
  const [confirmClear, setConfirmClear] = useState(false);

  const clientRows = cd.clients?.outcome.rows ?? [];
  const holdingRows = cd.holdings?.outcome.rows ?? [];
  const labels = useMemo(() => classLabels(holdingRows, cd.classMap), [holdingRows, cd.classMap]);

  const analyses = useMemo(() => {
    const numbers = allClientNumbers(clientRows, holdingRows);
    return numbers.map((n) => {
      const client = clientRows.find((c) => c.clientNumber === n) ?? null;
      const agg = aggregateClient(n, holdingRows, cd.classMap);
      const base = client ? findBaseline(state, client.profileId, cd.settings.referenceKind) : undefined;
      const reference = base && base.status !== 'missing' ? base.weights : null;
      const refLabel = client ? `${profile(client.profileId).label} ${cd.settings.referenceKind} allocation` : 'Profile allocation';
      return {
        analysis: analyseClient({
          client,
          aggregate: agg,
          reference,
          referenceLabel: refLabel,
          multiSectorGrowthFraction: state.multiSector.growthFraction,
          defaultCoverMonths: cd.settings.defaultCoverMonths,
          defaultTolerance: cd.settings.defaultTolerance,
        }),
        base,
      };
    });
  }, [clientRows, holdingRows, cd.classMap, cd.settings, state]);

  const current = analyses.find((a) => a.analysis.aggregate.clientNumber === selected) ?? null;
  const anyLoaded = Boolean(cd.clients || cd.holdings);
  const syntheticOnly = (!cd.clients || cd.clients.synthetic) && (!cd.holdings || cd.holdings.synthetic);

  return (
    <>
      <PageHeader
        title="Client Holdings"
        subtitle="Compare each client's current holdings with their risk profile's model allocation and their growth asset target, and check that cash covers monthly fees and withdrawals. Descriptive only — no trades or recommendations are generated."
        actions={
          anyLoaded &&
          (confirmClear ? (
            <>
              <Button variant="danger" onClick={() => { cd.clearAll(); setSelected(null); setConfirmClear(false); }}>Clear client data</Button>
              <Button onClick={() => setConfirmClear(false)}>Cancel</Button>
            </>
          ) : (
            <Button variant="danger" onClick={() => setConfirmClear(true)}>Clear client data…</Button>
          ))
        }
      />

      <div className="mb-6">
        <Notice tone="warning">
          <strong>Client data handling.</strong> Files loaded here are held in this browser tab's memory only. They are not saved to browser storage or sent anywhere, and are cleared when the page is reloaded or closed. This prototype has no logins or access control: do not load real client information until the firm's privacy and compliance process has approved it.
          {anyLoaded && (syntheticOnly ? <> Currently loaded: <strong>synthetic sample data</strong>.</> : <> Currently loaded: <strong>user-supplied files</strong>.</>)}
        </Notice>
      </div>

      <div className="mb-6 grid gap-6 lg:grid-cols-2">
        <FileCard
          title="1. Client details"
          description="Client number, Risk Profile, Growth Target %, Monthly Fees, Monthly Withdrawals. Optional: Cash Cover Months, Drift Tolerance pp, Notes."
          file={cd.clients}
          onText={(name, text, synthetic) => cd.setClients({ name, outcome: parseClientsCsv(text), loadedAt: new Date().toISOString(), synthetic })}
          onClear={() => cd.setClients(null)}
          sample={sampleClients}
          template={CLIENT_TEMPLATE}
          templateName="client-details-template.csv"
          unit="client"
        />
        <FileCard
          title="2. Current holdings"
          description="Client number, Asset Name, Asset Value, Asset Class, Asset Sub-Class. One row per holding; values in NZD ($ and commas allowed)."
          file={cd.holdings}
          onText={(name, text, synthetic) => cd.setHoldings({ name, outcome: parseHoldingsCsv(text), loadedAt: new Date().toISOString(), synthetic })}
          onClear={() => cd.setHoldings(null)}
          sample={sampleHoldings}
          template={HOLDINGS_TEMPLATE}
          templateName="holdings-template.csv"
          unit="holding"
        />
      </div>

      <div className="mb-6 grid gap-6 lg:grid-cols-[minmax(0,1fr)_minmax(0,1.3fr)]">
        <SettingsCard />
        <Card title="Asset class mapping">
          {labels.length === 0 ? (
            <p className="text-sm text-slate-500">Load holdings to see the asset class labels they use.</p>
          ) : (
            <>
              <TableWrap>
                <thead>
                  <tr><Th>Label in file</Th><Th align="right">Holdings</Th><Th align="right">Value</Th><Th>Maps to</Th></tr>
                </thead>
                <tbody>
                  {labels.map((l) => (
                    <tr key={l.key} className={l.mapped ? '' : 'bg-amber-50/60'}>
                      <Td className="font-medium">{l.label}</Td>
                      <Td align="right">{l.count}</Td>
                      <Td align="right">{money(l.value)}</Td>
                      <Td>
                        <select
                          aria-label={`Map ${l.label}`}
                          className={cx(inputCls, 'max-w-[15rem]', !l.mapped && 'border-amber-400')}
                          value={l.mapped ?? ''}
                          onChange={(e) => cd.setClassMap({ ...cd.classMap, [l.key]: (e.target.value || null) as AssetClassId | null })}
                        >
                          <option value="">Unmapped — flag &amp; exclude</option>
                          {ASSET_CLASSES.map((ac) => <option key={ac.id} value={ac.id}>{ac.label}</option>)}
                        </select>
                        {l.isDefault && l.mapped && <span className="ml-2 text-xs text-slate-500">standard name</span>}
                      </Td>
                    </tr>
                  ))}
                </tbody>
              </TableWrap>
              <p className="mt-2 text-xs text-slate-500">Only exact standard names are mapped automatically. Other labels (e.g. “Diversified”) stay unmapped until you choose — their value is shown as unmapped rather than guessed. Mappings are remembered in this browser; client data is not.</p>
            </>
          )}
        </Card>
      </div>

      <Card title={`Clients (${analyses.length})`} className="mb-6">
        {analyses.length === 0 ? (
          <EmptyState title="No client data loaded">Load client details and holdings above, or use the synthetic samples to try the workflow.</EmptyState>
        ) : (
          <TableWrap>
            <thead>
              <tr>
                <Th>Client</Th>
                <Th>Risk profile</Th>
                <Th align="right">Total value</Th>
                <Th align="right">Growth actual / target</Th>
                <Th align="right">Classes outside tolerance</Th>
                <Th align="right">Cash cover (months)</Th>
                <Th>Status</Th>
              </tr>
            </thead>
            <tbody>
              {analyses.map(({ analysis: a }) => {
                const n = a.aggregate.clientNumber;
                const st = statusOf(a);
                const short = a.cash.shortfall !== null && a.cash.shortfall > 0;
                return (
                  <tr key={n} className={cx('cursor-pointer hover:bg-slate-50', selected === n && 'bg-sky-50')} onClick={() => setSelected(n)}>
                    <Td className="font-medium"><button type="button" className="underline-offset-2 hover:underline" onClick={() => setSelected(n)}>{n}</button></Td>
                    <Td className="whitespace-nowrap">{a.client ? profile(a.client.profileId).label : <span className="text-orange-700">no details</span>}</Td>
                    <Td align="right">{a.aggregate.holdingCount ? money(a.aggregate.totalValue) : <span className="text-orange-700">no holdings</span>}</Td>
                    <Td align="right">
                      {a.aggregate.totalValue > 0 ? formatPct(a.split.growth, 1) : '—'} / {a.client?.growthTarget != null ? formatPct(a.client.growthTarget, 1) : '—'}
                      {a.growthGap !== null && <div className={cx('text-xs', Math.abs(a.growthGap) > a.tolerance ? 'text-amber-800' : 'text-slate-500')}>{formatPp(a.growthGap)}</div>}
                    </Td>
                    <Td align="right" className={a.breaches.length ? 'font-semibold text-amber-800' : ''}>{a.drift ? `${a.breaches.length} (±${a.tolerance} pp)` : '—'}</Td>
                    <Td align="right" className={short ? 'font-semibold text-amber-800' : ''}>
                      {a.cash.monthsCovered === null ? '—' : a.cash.monthsCovered === Infinity ? 'no outflows' : a.cash.monthsCovered.toFixed(1)}
                      {a.cash.monthsCovered !== null && <span className="text-xs font-normal text-slate-500"> / {a.cash.coverMonths}</span>}
                    </Td>
                    <Td><ValidationBadge status={st} count={a.issues.filter((i) => i.severity !== 'info').length} /></Td>
                  </tr>
                );
              })}
            </tbody>
          </TableWrap>
        )}
        {analyses.length > 0 && <p className="mt-2 text-xs text-slate-500">Select a client for details. Drift is measured against the profile's {cd.settings.referenceKind} allocation in Assumptions; check its status (synthetic / illustrative / approved) before relying on it.</p>}
      </Card>

      {current && <ClientDetail analysis={current.analysis} baseStatus={current.base?.status} baseVersion={current.base?.version} kind={cd.settings.referenceKind} holdings={holdingRows.filter((h) => h.clientNumber === current.analysis.aggregate.clientNumber)} />}
    </>
  );
}

function FileCard<T>({ title, description, file, onText, onClear, sample, template, templateName, unit }: {
  title: string;
  description: string;
  file: LoadedFile<T> | null;
  onText: (name: string, text: string, synthetic: boolean) => void;
  onClear: () => void;
  sample: string;
  template: string;
  templateName: string;
  unit: string;
}) {
  const o: ParseOutcome<T> | undefined = file?.outcome;
  return (
    <Card title={title} actions={file && <Button variant="ghost" onClick={onClear}>Remove</Button>}>
      <p className="mb-3 text-xs text-slate-600">{description}</p>
      <div className="flex flex-wrap items-center gap-2">
        <input
          type="file"
          accept=".csv,text/csv"
          aria-label={`${title} file`}
          className="max-w-full text-sm file:mr-2 file:rounded file:border file:border-slate-300 file:bg-white file:px-2 file:py-1"
          onChange={async (e) => {
            const f = e.target.files?.[0];
            if (!f) return;
            if (f.size > 5_000_000) { alert('File is larger than 5 MB; this prototype only accepts smaller CSV files.'); return; }
            onText(f.name, await f.text(), false);
            e.target.value = '';
          }}
        />
        <Button variant="ghost" onClick={() => onText('synthetic sample', sample, true)}>Load synthetic sample</Button>
        <Button variant="ghost" onClick={() => download(templateName, template)}>Template</Button>
      </div>
      {file && o && (
        <div className="mt-3 space-y-2 text-sm">
          <div className="flex flex-wrap items-center gap-2">
            <span className="font-medium">{file.name}</span>
            {file.synthetic && <StatusBadge status="synthetic" />}
            <span className="text-slate-600">{o.rows.length} {unit}(s) loaded</span>
            {o.errors.length > 0 && <span className="text-red-700">{o.errors.length} row(s) rejected</span>}
          </div>
          {o.fatal.length > 0 && <Notice tone="danger">{o.fatal.join(' ')}</Notice>}
          {o.errors.length > 0 && (
            <details className="rounded border border-red-200 bg-red-50 px-2 py-1 text-red-900" open={o.errors.length <= 5}>
              <summary className="cursor-pointer text-sm">Rejected rows (not used)</summary>
              <ul className="ml-4 list-disc text-xs">{o.errors.map((e) => <li key={e.line}>Line {e.line}: {e.message}</li>)}</ul>
            </details>
          )}
          {o.warnings.length > 0 && (
            <details className="rounded border border-amber-200 bg-amber-50 px-2 py-1 text-amber-900" open={o.warnings.length <= 5}>
              <summary className="cursor-pointer text-sm">Warnings ({o.warnings.length})</summary>
              <ul className="ml-4 list-disc text-xs">{o.warnings.map((e, i) => <li key={i}>Line {e.line}: {e.message}</li>)}</ul>
            </details>
          )}
        </div>
      )}
    </Card>
  );
}

function SettingsCard() {
  const cd = useClientData();
  const s = cd.settings;
  return (
    <Card title="Analysis settings">
      <div className="grid gap-3 sm:grid-cols-3">
        <Field label="Compare against">
          <select className={inputCls} value={s.referenceKind} onChange={(e) => cd.setSettings({ ...s, referenceKind: e.target.value as BaselineKind })}>
            <option value="strategic">Strategic</option>
            <option value="target">Target</option>
          </select>
        </Field>
        <Field label="Default cash cover" hint="Months of fees + withdrawals.">
          <PercentInput ariaLabel="Default cash cover months" suffix="mo" value={s.defaultCoverMonths} onChange={(v) => v !== null && v >= 0 && v <= 120 && cd.setSettings({ ...s, defaultCoverMonths: v })} />
        </Field>
        <Field label="Default drift tolerance" hint="± pp per asset class.">
          <PercentInput ariaLabel="Default drift tolerance" suffix="pp" value={s.defaultTolerance} onChange={(v) => v !== null && v >= 0 && v <= 100 && cd.setSettings({ ...s, defaultTolerance: v })} />
        </Field>
      </div>
      <p className="mt-3 text-xs text-slate-500">Defaults (12 months cover, ±2 pp) are working assumptions, not approved policy — the ±2 pp figure mirrors the asset-class tolerance in the existing allocation workbook. Per-client values in the client details file override them.</p>
    </Card>
  );
}

function ClientDetail({ analysis: a, baseStatus, baseVersion, kind, holdings }: {
  analysis: ClientAnalysis;
  baseStatus?: Parameters<typeof StatusBadge>[0]['status'];
  baseVersion?: number;
  kind: BaselineKind;
  holdings: { line: number; assetName: string; value: number; assetClass: string; subClass: string }[];
}) {
  const agg = a.aggregate;
  const c = a.client;
  const p = c ? profile(c.profileId) : null;
  return (
    <Card title={<span className="flex flex-wrap items-center gap-2">Client {agg.clientNumber}{p && <span className="font-normal text-slate-600">· {p.label}</span>}<ValidationBadge status={statusOf(a)} /></span>}>
      <div className="grid gap-6 xl:grid-cols-[minmax(0,1fr)_340px]">
        <div className="space-y-6">
          <TableWrap>
            <thead>
              <tr>
                <Th>Asset class</Th>
                <Th align="right">Value</Th>
                <Th align="right">Actual</Th>
                <Th align="right"><span className="flex flex-col items-end gap-1">{kind === 'strategic' ? 'Strategic' : 'Target'}{baseStatus && <span className="flex items-center gap-1 normal-case"><StatusBadge status={baseStatus} /><span className="text-[10px]">v{baseVersion}</span></span>}</span></Th>
                <Th align="right">Drift</Th>
              </tr>
            </thead>
            <tbody>
              {ASSET_CLASSES.map((ac) => {
                const d = a.drift?.[ac.id] ?? null;
                const breach = a.breaches.includes(ac.id);
                return (
                  <tr key={ac.id} className={breach ? 'bg-amber-50' : ''}>
                    <Td className="font-medium">{ac.label}</Td>
                    <Td align="right">{money(agg.valueByClass[ac.id])}</Td>
                    <Td align="right">{formatPct(agg.weights[ac.id])}</Td>
                    <Td align="right">{a.reference ? (a.reference[ac.id] === null ? <span className="text-orange-700">missing</span> : formatPct(a.reference[ac.id])) : '—'}</Td>
                    <Td align="right" className={cx(breach && 'font-semibold text-amber-800')}>{a.drift ? <>{breach && <span aria-hidden>! </span>}{formatPp(d)}</> : '—'}</Td>
                  </tr>
                );
              })}
              {agg.unmappedValue !== 0 && (
                <tr className="bg-amber-50/60">
                  <Td className="font-medium text-amber-900">Unmapped ({agg.unmapped.map((u) => u.label).join(', ')})</Td>
                  <Td align="right">{money(agg.unmappedValue)}</Td>
                  <Td align="right">{formatPct(agg.totalValue > 0 ? (agg.unmappedValue / agg.totalValue) * 100 : null)}</Td>
                  <Td colSpan={2} className="text-xs text-amber-900">Map these labels above to include them.</Td>
                </tr>
              )}
              <tr className="bg-slate-50 font-semibold">
                <Td>Total</Td>
                <Td align="right">{money(agg.totalValue)}</Td>
                <Td align="right">{agg.totalValue > 0 ? formatPct(100) : '—'}</Td>
                <Td colSpan={2} className="text-right text-xs font-normal text-slate-600">Tolerance ±{a.tolerance} pp{c?.driftTolerance != null ? ' (client-specific)' : ''}</Td>
              </tr>
            </tbody>
          </TableWrap>

          {agg.totalValue > 0 && (
            <StackedAllocationChart
              rows={[
                ...(a.reference ? [{ key: 'ref', label: kind === 'strategic' ? 'Strategic' : 'Target', weights: a.reference, tag: baseStatus }] : []),
                { key: 'act', label: 'Current holdings', weights: agg.weights, tag: `client ${agg.clientNumber}` },
              ]}
            />
          )}

          <details className="rounded border border-slate-200">
            <summary className="cursor-pointer px-3 py-2 text-sm font-medium">Holdings ({holdings.length})</summary>
            <div className="px-3 pb-3">
              <TableWrap>
                <thead><tr><Th>Asset name</Th><Th>Asset class</Th><Th>Sub-class</Th><Th align="right">Value</Th></tr></thead>
                <tbody>
                  {holdings.map((h) => (
                    <tr key={h.line}><Td>{h.assetName || '—'}</Td><Td>{h.assetClass}</Td><Td>{h.subClass || '—'}</Td><Td align="right" className={h.value < 0 ? 'text-red-700' : ''}>{money(h.value)}</Td></tr>
                  ))}
                </tbody>
              </TableWrap>
            </div>
          </details>
        </div>

        <div className="space-y-6">
          <div>
            <h3 className="mb-2 text-sm font-semibold text-slate-800">Findings</h3>
            <IssueList issues={a.issues.map((i) => ({ code: 'NO_DATA' as const, severity: i.severity, message: i.message }))} emptyText="Holdings are within tolerance and cash covers the required period." />
          </div>

          <div>
            <h3 className="mb-2 text-sm font-semibold text-slate-800">Growth / defensive</h3>
            <SplitBar growth={a.split.growth} defensive={a.split.defensive} unclassified={a.split.unclassified} />
            <dl className="mt-2 grid grid-cols-2 gap-y-1 text-sm">
              <dt className="text-slate-600">Client growth target</dt><dd className="text-right tabular-nums">{formatPct(c?.growthTarget ?? null)}</dd>
              <dt className="text-slate-600">Actual growth</dt><dd className="text-right tabular-nums">{agg.totalValue > 0 ? formatPct(a.split.growth) : '—'}</dd>
              <dt className="text-slate-600">Actual − target</dt><dd className="text-right tabular-nums">{formatPp(a.growthGap)}</dd>
              <dt className="text-slate-600">Profile model − target</dt><dd className="text-right tabular-nums">{formatPp(a.referenceGrowthGap)}</dd>
            </dl>
          </div>

          <div>
            <h3 className="mb-2 text-sm font-semibold text-slate-800">Cash coverage</h3>
            <dl className="grid grid-cols-2 gap-y-1 text-sm">
              <dt className="text-slate-600">Monthly fees</dt><dd className="text-right tabular-nums">{money(c?.monthlyFees)}</dd>
              <dt className="text-slate-600">Monthly withdrawals</dt><dd className="text-right tabular-nums">{money(c?.monthlyWithdrawals)}</dd>
              <dt className="text-slate-600">Monthly outflow</dt><dd className="text-right font-medium tabular-nums">{money(a.cash.monthlyOutflow)}</dd>
              <dt className="text-slate-600">Cover required</dt><dd className="text-right tabular-nums">{a.cash.coverMonths} months{c?.cashCoverMonths != null ? ' (client)' : ''}</dd>
              <dt className="text-slate-600">Cash required</dt><dd className="text-right tabular-nums">{money(a.cash.requiredCash)}</dd>
              <dt className="text-slate-600">Cash held</dt><dd className="text-right tabular-nums">{money(a.cash.actualCash)}</dd>
              <dt className="text-slate-600">Months covered</dt><dd className={cx('text-right tabular-nums', a.cash.shortfall ? 'font-semibold text-amber-800' : '')}>{a.cash.monthsCovered === null ? '—' : a.cash.monthsCovered === Infinity ? 'no outflows' : a.cash.monthsCovered.toFixed(1)}</dd>
              <dt className="text-slate-600">Shortfall</dt><dd className={cx('text-right tabular-nums', a.cash.shortfall ? 'font-semibold text-amber-800' : '')}>{money(a.cash.shortfall)}</dd>
              <dt className="text-slate-600">Cash needed (% of portfolio)</dt><dd className="text-right tabular-nums">{formatPct(a.cash.requiredCashPct)}</dd>
              <dt className="text-slate-600">Profile cash weight</dt><dd className="text-right tabular-nums">{formatPct(a.cash.referenceCashPct)}</dd>
            </dl>
          </div>

          {c?.notes && (
            <div>
              <h3 className="mb-1 text-sm font-semibold text-slate-800">Notes</h3>
              <p className="whitespace-pre-wrap text-sm text-slate-700">{c.notes}</p>
            </div>
          )}
        </div>
      </div>
    </Card>
  );
}
