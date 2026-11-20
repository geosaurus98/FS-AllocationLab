import { useMemo, useState } from 'react';
import sampleClients from '../data/sample-clients.csv?raw';
import sampleHoldings from '../data/sample-holdings.csv?raw';
import { SplitBar, StackedAllocationChart } from '../components/charts';
import { PercentInput } from '../components/PercentInput';
import { Button, Card, EmptyState, Field, inputCls, IssueList, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, ValidationBadge, cx } from '../components/ui';
import { toCsv } from '../lib/csv';
import { formatPct, formatPp } from '../lib/format';
import { buildSecurityMaster, normName } from '../lib/datasheet';
import {
  aggregateClient,
  allClientNumbers,
  analyseClient,
  classifyHoldings,
  classLabels,
  fmtMoney,
  parseClientsCsv,
  parseHoldingsCsv,
  SOURCE_LABEL,
  type ClassifiedHolding,
  type ClientAnalysis,
  type ClientRecord,
  type ParseOutcome,
} from '../lib/holdings';
import { ASSET_CLASSES, PROFILES, profile } from '../lib/reference';
import type { AssetClassId, BaselineKind, ProfileId, ValidationStatus } from '../lib/types';
import { readXlsxSheet } from '../lib/xlsx';
import { parseCsv, detectDelimiter } from '../lib/csv';
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
  ['Client number', 'Risk Profile', 'Growth Target %', 'Annual Fee %', 'Monthly Fees', 'Monthly Withdrawals', 'Cash Cover Months', 'Drift Tolerance pp', 'Notes'],
  ['EXAMPLE-1', 'Balanced 60/40', 60, 1, '', 2500, '', '', 'use Annual Fee % and/or a fixed Monthly Fees amount'],
]);
const HOLDINGS_TEMPLATE = toCsv([
  ['Client number', 'Asset Name', 'Asset Value', 'Asset Class', 'Asset Sub-Class'],
  ['EXAMPLE-1', 'Example Cash Fund', '20000.00', '', ''],
  ['EXAMPLE-1', 'Example holding with class supplied', '5000.00', 'Global Equities', 'International'],
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

  const clientRows = useMemo(() => {
    const manual = new Set(cd.manualClients.map((c) => c.clientNumber));
    return [...(cd.clients?.outcome.rows ?? []).filter((c) => !manual.has(c.clientNumber)), ...cd.manualClients];
  }, [cd.clients, cd.manualClients]);
  const rawHoldings = cd.holdings?.outcome.rows;
  const holdingRows = useMemo(() => classifyHoldings(rawHoldings ?? [], cd.datasheet?.master ?? null, cd.overrides), [rawHoldings, cd.datasheet, cd.overrides]);
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
  const anyLoaded = Boolean(cd.clients || cd.holdings || cd.manualClients.length);
  const syntheticOnly = (!cd.clients || cd.clients.synthetic) && (!cd.holdings || cd.holdings.synthetic) && cd.manualClients.length === 0;

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

      <DatasheetCard />

      <div className="mb-6 grid gap-6 lg:grid-cols-2">
        <div className="space-y-6">
        <FileCard
          title="2. Client details"
          description="Client number, Risk Profile, Growth Target %, Annual Fee % and/or Monthly Fees ($), Monthly Withdrawals. Optional: Cash Cover Months, Drift Tolerance pp, Notes."
          file={cd.clients}
          onText={(name, text, synthetic) => cd.setClients({ name, outcome: parseClientsCsv(text), loadedAt: new Date().toISOString(), synthetic })}
          onClear={() => cd.setClients(null)}
          sample={sampleClients}
          template={CLIENT_TEMPLATE}
          templateName="client-details-template.csv"
          unit="client"
        />
        <ManualClientCard />
        </div>
        <FileCard
          title="3. Current holdings"
          description="Client number (or Portfolio Id), Asset Name, Asset Value (NZD). Asset Class / Sub-Class are optional — when absent, holdings are classified from the investment datasheet by Apex, FS or OA name. CSV or tab-separated (paste from Excel works)."
          file={cd.holdings}
          onText={(name, text, synthetic) => cd.setHoldings({ name, outcome: parseHoldingsCsv(text), loadedAt: new Date().toISOString(), synthetic })}
          onClear={() => cd.setHoldings(null)}
          sample={sampleHoldings}
          template={HOLDINGS_TEMPLATE}
          templateName="holdings-template.csv"
          unit="holding"
        />
      </div>

      {holdingRows.length > 0 && <ClassificationCard holdings={holdingRows} />}

      <div className="mb-6 grid gap-6 lg:grid-cols-[minmax(0,1fr)_minmax(0,1.3fr)]">
        <SettingsCard />
        <Card title="Asset class label mapping">
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
              <p className="mt-2 text-xs text-slate-500">Asset class labels come from the holdings file or the datasheet. Only exact standard names are mapped automatically. Other labels (e.g. “Diversified”) stay unmapped until you choose — their value is shown as unmapped rather than guessed. Mappings are remembered in this browser; client data is not.</p>
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
  const [paste, setPaste] = useState<string | null>(null);
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
        <Button variant="ghost" onClick={() => setPaste(paste === null ? '' : null)}>{paste === null ? 'Paste…' : 'Close paste'}</Button>
        <Button variant="ghost" onClick={() => onText('synthetic sample', sample, true)}>Load synthetic sample</Button>
        <Button variant="ghost" onClick={() => download(templateName, template)}>Template</Button>
      </div>
      {paste !== null && (
        <div className="mt-3">
          <textarea aria-label={`${title} paste`} className={`${inputCls} font-mono text-xs`} rows={6} value={paste} onChange={(e) => setPaste(e.target.value)} placeholder="Paste rows including the header row (copied from Excel or CSV)" />
          <div className="mt-2 flex gap-2">
            <Button variant="primary" disabled={!paste.trim()} onClick={() => { onText('pasted data', paste, false); setPaste(null); }}>Load pasted data</Button>
          </div>
        </div>
      )}
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
  holdings: ClassifiedHolding[];
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
                <thead><tr><Th>Asset name</Th><Th>Asset class</Th><Th>Sub-class</Th><Th>Type</Th><Th>Classified by</Th><Th align="right">Value</Th></tr></thead>
                <tbody>
                  {[...holdings].sort((x, y) => (x.assetClass || 'zzz').localeCompare(y.assetClass || 'zzz') || y.value - x.value).map((h) => (
                    <tr key={h.line} className={!h.assetClass ? 'bg-amber-50/60' : ''}>
                      <Td className="max-w-md">{h.assetName || '—'}</Td>
                      <Td className="whitespace-nowrap">{h.assetClass || <span className="text-amber-800">unclassified</span>}</Td>
                      <Td>{h.subClass || '—'}</Td>
                      <Td className="text-xs">{h.type || '—'}</Td>
                      <Td className="whitespace-nowrap text-xs text-slate-600">{SOURCE_LABEL[h.source]}{h.match?.status === 'matched' && <> · {h.match.via}{h.match.stripped ? ' (detail suffix removed)' : ''}</>}</Td>
                      <Td align="right" className={h.value < 0 ? 'text-red-700' : ''}>{money(h.value)}</Td>
                    </tr>
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
              <dt className="text-slate-600">Monthly fees{c?.annualFeePct != null && <span className="text-xs"> ({c.annualFeePct}% p.a.{c.monthlyFees ? ' + fixed' : ''})</span>}</dt><dd className="text-right tabular-nums">{money(a.cash.monthlyFees)}</dd>
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

function DatasheetCard() {
  const cd = useClientData();
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const d = cd.datasheet;
  return (
    <Card title="1. Investment datasheet" className="mb-6" actions={d && <Button variant="ghost" onClick={() => cd.setDatasheet(null)}>Remove</Button>}>
      <p className="mb-3 text-xs text-slate-600">
        Load the firm's investment datasheet (<code>Datasheet.xlsx</code>, or the <em>Data</em> sheet of the recommendations workbook, or a CSV export). Holdings without an Asset Class are matched to it by Apex, FS or OA name to get their Asset Class and Sub-Asset Class. Matching is exact (ignoring case and spacing); term-deposit and bond detail suffixes like “(188,000.00; 3.53%; 18 Months; 24/12/2026)” are removed before a second attempt. Nothing is fuzzy-matched or guessed.
      </p>
      <div className="flex flex-wrap items-center gap-2">
        <input
          type="file"
          accept=".xlsx,.xlsm,.csv,.tsv,.txt"
          aria-label="Investment datasheet file"
          disabled={busy}
          className="max-w-full text-sm file:mr-2 file:rounded file:border file:border-slate-300 file:bg-white file:px-2 file:py-1"
          onChange={async (e) => {
            const f = e.target.files?.[0];
            e.target.value = '';
            if (!f) return;
            setError(null);
            setBusy(true);
            try {
              let rows: string[][];
              let sheetName = '—';
              if (/\.(xlsx|xlsm)$/i.test(f.name)) {
                const sheet = await readXlsxSheet(await f.arrayBuffer(), 'Data');
                rows = sheet.rows;
                sheetName = sheet.sheetName;
              } else {
                const text = await f.text();
                rows = parseCsv(text, detectDelimiter(text));
              }
              cd.setDatasheet({ name: f.name, sheetName, master: buildSecurityMaster(rows), loadedAt: new Date().toISOString() });
            } catch (err) {
              setError(err instanceof Error ? err.message : String(err));
            } finally {
              setBusy(false);
            }
          }}
        />
        {busy && <span className="text-sm text-slate-500">Reading…</span>}
      </div>
      {error && <div className="mt-3"><Notice tone="danger">{error}</Notice></div>}
      {d && (
        <p className="mt-3 text-sm">
          <span className="font-medium">{d.name}</span>
          {d.sheetName !== '—' && <span className="text-slate-600"> · sheet “{d.sheetName}”</span>}
          <span className="text-slate-600"> · {d.master.records.length} investments</span>
          {d.master.records.filter((r) => !r.assetClass).length > 0 && <span className="text-slate-500"> ({d.master.records.filter((r) => !r.assetClass).length} without an asset class)</span>}
        </p>
      )}
      {!d && <p className="mt-2 text-xs text-slate-500">Optional if your holdings file already has Asset Class. Held in memory only; not saved.</p>}
    </Card>
  );
}

function ClassificationCard({ holdings }: { holdings: ClassifiedHolding[] }) {
  const cd = useClientData();
  const bySource = new Map<string, { count: number; value: number }>();
  for (const h of holdings) {
    const e = bySource.get(h.source) ?? { count: 0, value: 0 };
    e.count++;
    e.value += h.value;
    bySource.set(h.source, e);
  }
  const unresolved = new Map<string, { name: string; value: number; count: number; reason: string }>();
  for (const h of holdings) {
    if (h.assetClass) continue;
    const k = normName(h.assetName);
    const e = unresolved.get(k) ?? { name: h.assetName, value: 0, count: 0, reason: SOURCE_LABEL[h.source] };
    e.value += h.value;
    e.count++;
    unresolved.set(k, e);
  }
  const overridden = Object.entries(cd.overrides);
  const setOv = (key: string, label: string) => {
    const next = { ...cd.overrides };
    if (label) next[key] = label;
    else delete next[key];
    cd.setOverrides(next);
  };
  return (
    <Card title="Holding classification" className="mb-6">
      <div className="flex flex-wrap gap-2 text-sm">
        {[...bySource.entries()].map(([src, v]) => (
          <span key={src} className={cx('rounded border px-2 py-1', src === 'datasheet' || src === 'file' || src === 'override' ? 'border-slate-200 bg-slate-50' : 'border-amber-300 bg-amber-50 text-amber-900')}>
            {SOURCE_LABEL[src as ClassifiedHolding['source']]}: <strong>{v.count}</strong> · {money(v.value)}
          </span>
        ))}
      </div>
      {!cd.datasheet && holdings.some((h) => h.source === 'not-found') && (
        <p className="mt-3 text-sm text-amber-800">No investment datasheet is loaded, and some holdings have no Asset Class. Load the datasheet above to classify them.</p>
      )}
      {unresolved.size > 0 && (
        <div className="mt-4">
          <h3 className="mb-2 text-sm font-semibold text-slate-800">Unclassified holdings ({unresolved.size}) — excluded from asset-class weights until assigned</h3>
          <TableWrap>
            <thead><tr><Th>Asset name</Th><Th>Reason</Th><Th align="right">Value</Th><Th>Assign asset class</Th></tr></thead>
            <tbody>
              {[...unresolved.entries()].map(([k, u]) => (
                <tr key={k}>
                  <Td className="max-w-md">{u.name}{u.count > 1 && <span className="text-xs text-slate-500"> ×{u.count}</span>}</Td>
                  <Td className="text-xs text-amber-800">{u.reason}</Td>
                  <Td align="right">{money(u.value)}</Td>
                  <Td>
                    <select aria-label={`Assign class for ${u.name}`} className={cx(inputCls, 'max-w-[14rem]')} value="" onChange={(e) => setOv(k, e.target.value)}>
                      <option value="">Leave unclassified</option>
                      {ASSET_CLASSES.map((ac) => <option key={ac.id} value={ac.label}>{ac.label}</option>)}
                    </select>
                  </Td>
                </tr>
              ))}
            </tbody>
          </TableWrap>
        </div>
      )}
      {overridden.length > 0 && (
        <details className="mt-4 text-sm">
          <summary className="cursor-pointer font-medium">Manual assignments ({overridden.length}) — remembered in this browser</summary>
          <ul className="mt-2 space-y-1">
            {overridden.map(([k, label]) => (
              <li key={k} className="flex items-center justify-between gap-2 border-b border-slate-100 py-1">
                <span className="truncate">{k} → <strong>{label}</strong></span>
                <Button variant="ghost" onClick={() => setOv(k, '')}>Remove</Button>
              </li>
            ))}
          </ul>
          <p className="mt-1 text-xs text-slate-500">Assignments are stored by investment name only (no client numbers or values) and apply to every client holding that investment. Consider adding these investments to the datasheet instead.</p>
        </details>
      )}
    </Card>
  );
}

function ManualClientCard() {
  const cd = useClientData();
  const [num, setNum] = useState('');
  const [prof, setProf] = useState<ProfileId>('growth');
  const [growth, setGrowth] = useState<number | null>(null);
  const [feePct, setFeePct] = useState<number | null>(null);
  const [feeFixed, setFeeFixed] = useState<number | null>(null);
  const [wd, setWd] = useState<number | null>(null);
  const [cover, setCover] = useState<number | null>(null);
  const [tol, setTol] = useState<number | null>(null);
  const problems: string[] = [];
  if (!num.trim()) problems.push('client / portfolio number');
  if (growth === null || growth < 0 || growth > 100) problems.push('growth target 0–100%');
  if (feePct !== null && (feePct < 0 || feePct > 100)) problems.push('annual fee 0–100%');
  if (feeFixed !== null && feeFixed < 0) problems.push('fixed fee ≥ 0');
  if (wd === null || wd < 0) problems.push('monthly withdrawals ≥ 0');
  const add = () => {
    const rec: ClientRecord = {
      line: 0,
      clientNumber: num.trim(),
      profileId: prof,
      growthTarget: growth,
      annualFeePct: feePct,
      monthlyFees: feeFixed,
      monthlyWithdrawals: wd,
      cashCoverMonths: cover,
      driftTolerance: tol,
      notes: 'Entered manually',
    };
    cd.upsertManualClient(rec);
    setNum('');
  };
  return (
    <Card title="…or enter a client">
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-4">
        <Field label="Client / Portfolio no."><input className={inputCls} value={num} onChange={(e) => setNum(e.target.value)} /></Field>
        <Field label="Risk profile">
          <select className={inputCls} value={prof} onChange={(e) => setProf(e.target.value as ProfileId)}>
            {PROFILES.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
          </select>
        </Field>
        <Field label="Growth target"><PercentInput ariaLabel="Growth target" value={growth} onChange={setGrowth} /></Field>
        <Field label="Annual fee (% p.a.)"><PercentInput ariaLabel="Annual fee percent" value={feePct} onChange={setFeePct} /></Field>
        <Field label="Fixed fees / month"><PercentInput ariaLabel="Fixed monthly fees" suffix="$" value={feeFixed} onChange={setFeeFixed} /></Field>
        <Field label="Withdrawals / month"><PercentInput ariaLabel="Monthly withdrawals" suffix="$" value={wd} onChange={setWd} /></Field>
        <Field label="Cover (optional)"><PercentInput ariaLabel="Cash cover months" suffix="mo" placeholder="default" value={cover} onChange={setCover} /></Field>
        <Field label="Tolerance (optional)"><PercentInput ariaLabel="Drift tolerance" suffix="pp" placeholder="default" value={tol} onChange={setTol} /></Field>
      </div>
      <div className="mt-3 flex flex-wrap items-center gap-2">
        <Button variant="primary" disabled={problems.length > 0} onClick={add}>Add client</Button>
        {problems.length > 0 && <span className="text-xs text-slate-500">Needs: {problems.join(', ')}</span>}
      </div>
      {cd.manualClients.length > 0 && (
        <ul className="mt-3 space-y-1 text-sm">
          {cd.manualClients.map((c) => (
            <li key={c.clientNumber} className="flex items-center justify-between gap-2 border-b border-slate-100 py-1">
              <span><strong>{c.clientNumber}</strong> · {profile(c.profileId).label} · growth {formatPct(c.growthTarget, 0)} · fee {c.annualFeePct !== null ? `${c.annualFeePct}% pa` : ''}{c.monthlyFees !== null ? ` $${fmtMoney(c.monthlyFees)}/mo` : ''}{c.annualFeePct === null && c.monthlyFees === null ? 'none' : ''} · withdrawals {money(c.monthlyWithdrawals)}/mo</span>
              <Button variant="ghost" onClick={() => cd.removeManualClient(c.clientNumber)}>Remove</Button>
            </li>
          ))}
        </ul>
      )}
      <p className="mt-2 text-xs text-slate-500">Manually entered clients override a client details file row with the same number. Held in memory only. No client names are needed — use the portfolio number.</p>
    </Card>
  );
}
