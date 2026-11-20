import { useEffect, useMemo, useState } from 'react';
import { PercentInput } from '../components/PercentInput';
import { Button, Card, EmptyState, Field, inputCls, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, cx } from '../components/ui';
import type { WorkbookModel } from '../engine/workbook/evaluator';
import { dataRowsFromWorkbook, importApexHoldings, type ImportResult, type RawHolding } from '../engine/workbook/importHoldings';
import { RECOMMENDATION_PROFILES, runRecommendation, type ClientInputs, type RecommendationResult } from '../engine/workbook/run';
import { Workbook } from '../engine/workbook/evaluator';
import { BLANK, isErr, type Scalar } from '../engine/workbook/values';
import { toCsv } from '../lib/csv';
import { parseMoney } from '../lib/holdings';
import { useEngineSettings } from '../state/EngineSettingsContext';

/**
 * Recommendations: runs the recommendations workbook's own formulas (extracted from the .xlsm) for one
 * client. Client data is held in this page's memory only.
 */

const SAMPLE = `Portfolio Id\tAsset name\tAsset value NZD
DEMO-1\tNZD Call Account\t40000
DEMO-1\tANZ Bank New Zealand Ltd (100,000.00; 4.10%; 12 Months; 20/08/2027)\t101500
DEMO-1\tHarbour NZ Corporate Bond Fund (PIE) - XR\t60000
DEMO-1\tSimplicity NZ Share Fund (PIE)\t45000
DEMO-1\tVanguard S&P 500 ETF - USD (VOO)\t80000
DEMO-1\tGoodman Property Trust Units - NZD (GMT)\t12000`;

const fmt$ = (v: Scalar | number | null) => (typeof v === 'number' ? v.toLocaleString('en-NZ', { style: 'currency', currency: 'NZD', maximumFractionDigits: 0 }) : v === null || v === BLANK || v === '' ? '—' : isErr(v) ? v.code : String(v));
const fmtPct = (v: Scalar | number | null, d = 1) => (typeof v === 'number' ? `${v.toFixed(d)}%` : v === null || v === BLANK || v === '' ? '—' : isErr(v) ? v.code : String(v));
const n = (v: Scalar) => (typeof v === 'number' ? v : 0);

function parsePaste(text: string): { rows: RawHolding[]; errors: string[]; portfolioIds: string[] } {
  const lines = text.replace(/\r/g, '').split('\n').filter((l) => l.trim());
  const rows: RawHolding[] = [];
  const errors: string[] = [];
  const ids = new Set<string>();
  lines.forEach((line, i) => {
    const cols = line.includes('\t') ? line.split('\t') : line.split(',');
    if (i === 0 && /asset/i.test(line) && !/\d{3,}\.?\d*\s*$/.test(line)) return; // header row
    const [a, b, c] = cols.map((x) => x.trim());
    const [id, name, value] = cols.length >= 3 ? [a!, b!, c!] : ['', a!, b ?? ''];
    const m = parseMoney(value);
    if (!name || !m.ok || m.value === null) {
      errors.push(`Line ${i + 1}: could not read "${line.slice(0, 80)}"`);
      return;
    }
    if (id) ids.add(id);
    rows.push({ name, value: m.value });
  });
  return { rows, errors, portfolioIds: [...ids] };
}

export function Recommendations() {
  const es = useEngineSettings();
  const [model, setModel] = useState<WorkbookModel | null>(null);
  const [loadError, setLoadError] = useState<string | null>(null);
  useEffect(() => {
    import('../engine/workbook/model.json')
      .then((m) => setModel(m.default as unknown as WorkbookModel))
      .catch((e: unknown) => setLoadError(String(e)));
  }, []);

  const [paste, setPaste] = useState('');
  const [imported, setImported] = useState<{ result: ImportResult; parseErrors: string[]; ids: string[] } | null>(null);
  const [profile, setProfile] = useState<ClientInputs['profile']>('Balanced');
  const [growth, setGrowth] = useState<number | null>(60);
  const [customCash, setCustomCash] = useState<number | null>(2);
  const [feePct, setFeePct] = useState<number | null>(null);
  const [feeFixed, setFeeFixed] = useState<number | null>(null);
  const [withdrawals, setWithdrawals] = useState<number | null>(null);
  const [addRemove, setAddRemove] = useState<number | null>(null);
  const [cashOverride, setCashOverride] = useState<number | null>(null);
  const [date, setDate] = useState(() => new Date().toISOString().slice(0, 10));
  const [updateButton, setUpdateButton] = useState(false);
  const [overrides, setOverrides] = useState<Record<string, number | null>>({});
  const [result, setResult] = useState<RecommendationResult | null>(null);
  const [busy, setBusy] = useState(false);
  const [runError, setRunError] = useState<string | null>(null);

  const dataRows = useMemo(() => {
    if (!model) return [];
    const wb = new Workbook(model);
    return dataRowsFromWorkbook((r, c) => wb.get('Data', r, c), model.sheets.Data!.maxRow);
  }, [model]);

  const total = imported?.result.total ?? 0;
  const annualFees = (feePct ?? 0) / 100 * total + (feeFixed ?? 0) + (withdrawals ?? 0) * 12;

  const doImport = (text: string) => {
    const p = parsePaste(text);
    setImported({ result: importApexHoldings(p.rows, dataRows), parseErrors: p.errors, ids: p.portfolioIds });
    setResult(null);
    setOverrides({});
  };

  const run = (ov = overrides) => {
    if (!model || !imported) return;
    setBusy(true);
    setRunError(null);
    setTimeout(() => {
      try {
        const r = runRecommendation(model, es.current.settings, imported.result, {
          profile,
          custom: profile === 'Custom' ? { growth: (growth ?? 0) / 100, cash: (customCash ?? 0) / 100 } : null,
          portfolioDate: date,
          feesAndWithdrawals: Math.round(annualFees * 100) / 100,
          addRemove,
          cashTotalOverride: cashOverride,
          sectorOverrides: {},
          overrides: ov,
          updateSheetButton: updateButton,
        });
        setResult(r);
      } catch (e) {
        setRunError(e instanceof Error ? e.message : String(e));
      } finally {
        setBusy(false);
      }
    }, 20);
  };

  const exportCsv = () => {
    if (!result) return;
    const rows: (string | number | null)[][] = [['Section', 'Holding', 'Current', 'Recommended', 'Change', 'Recommended %', 'Override']];
    for (const s of result.sections) for (const r of s.rows) rows.push([s.label, r.name, n(r.value), n(r.recommended), n(r.change), typeof r.recommendedPct === 'number' ? Math.round(r.recommendedPct * 100) / 100 : null, typeof r.override === 'number' ? r.override : null]);
    const url = URL.createObjectURL(new Blob([toCsv(rows)], { type: 'text/csv' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = `recommendations-${date}.csv`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  };

  const customInvalid = profile === 'Custom' && (growth === null || customCash === null || growth < 0 || growth > 100 || customCash < 0 || growth + customCash > 100);

  return (
    <>
      <PageHeader
        title="Recommendations"
        subtitle={<>Runs the formulas of <em>{model?.source ?? 'the recommendations workbook'}</em> for one client: sector targets, model funds, fixed-interest ladder and trade limits, exactly as the workbook calculates them. Recommendations require adviser review.</>}
      />
      <div className="mb-6">
        <Notice tone="warning">
          <strong>Client data handling.</strong> Holdings pasted here stay in this page's memory only — nothing is saved or sent anywhere, and everything is cleared when the page is reloaded. Firm settings (version {es.current.version}, <StatusBadge status={es.current.status === 'approved' ? 'approved' : 'draft'} />) come from the Settings page.
        </Notice>
      </div>
      {loadError && <Notice tone="danger">The workbook model could not be loaded: {loadError}</Notice>}

      <div className="mb-6 grid gap-6 xl:grid-cols-2">
        <Card title="1. Holdings">
          <p className="mb-2 text-xs text-slate-600">Paste the custodian export (Portfolio Id, Asset name, Asset value NZD — tab or comma separated, header optional). Apex names are converted to FS names and term deposits to the workbook's “Bank Term Deposit rate% dd/mm/yyyy” form.</p>
          <textarea aria-label="Holdings paste" className={`${inputCls} font-mono text-xs`} rows={8} value={paste} onChange={(e) => setPaste(e.target.value)} placeholder="Portfolio Id	Asset name	Asset value NZD" />
          <div className="mt-2 flex flex-wrap gap-2">
            <Button variant="primary" disabled={!model || !paste.trim()} onClick={() => doImport(paste)}>Import holdings</Button>
            <Button variant="ghost" disabled={!model} onClick={() => { setPaste(SAMPLE); doImport(SAMPLE); }}>Load synthetic sample</Button>
            {imported && <Button variant="ghost" onClick={() => { setImported(null); setResult(null); setPaste(''); }}>Clear</Button>}
          </div>
          {!model && !loadError && <p className="mt-2 text-sm text-slate-500">Loading workbook model…</p>}
          {imported && (
            <div className="mt-3 space-y-2 text-sm">
              <p>
                <strong>{imported.result.holdings.length}</strong> holdings · {fmt$(imported.result.total)}
                {imported.ids.length > 0 && <span className="text-slate-500"> · portfolio {imported.ids.join(', ')}</span>}
              </p>
              <p className="text-xs text-slate-600">
                {Object.entries(imported.result.bySection).filter(([, v]) => v.length).map(([k, v]) => `${k.toUpperCase()} ${v.length}`).join(' · ')}
                {' · '}
                {imported.result.holdings.filter((h) => h.matchedBy === 'term-deposit').length} term deposits, {imported.result.holdings.filter((h) => h.matchedBy === 'security-code').length} matched by security code
              </p>
              {[...imported.parseErrors, ...imported.result.errors].length > 0 && (
                <Notice tone="danger"><ul className="ml-4 list-disc">{[...imported.parseErrors, ...imported.result.errors].map((e, i) => <li key={i}>{e}</li>)}</ul></Notice>
              )}
            </div>
          )}
        </Card>

        <Card title="2. Client settings">
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Risk profile">
              <select className={inputCls} value={profile} onChange={(e) => setProfile(e.target.value as ClientInputs['profile'])}>
                {RECOMMENDATION_PROFILES.map((p) => <option key={p} value={p}>{p}</option>)}
              </select>
            </Field>
            <Field label="Portfolio date"><input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
            {profile === 'Custom' && (
              <>
                <Field label="Growth assets" hint="Custom profile: equities and property are split like Balanced."><PercentInput ariaLabel="Custom growth" value={growth} onChange={setGrowth} /></Field>
                <Field label="Cash" hint={`Fixed income = ${growth !== null && customCash !== null ? (100 - growth - customCash).toFixed(2) : '—'}%`}><PercentInput ariaLabel="Custom cash" value={customCash} onChange={setCustomCash} /></Field>
              </>
            )}
            <Field label="Annual fee (% of portfolio)"><PercentInput ariaLabel="Annual fee percent" value={feePct} onChange={setFeePct} /></Field>
            <Field label="Other fees per year"><PercentInput suffix="$" ariaLabel="Fixed annual fees" value={feeFixed} onChange={setFeeFixed} /></Field>
            <Field label="Withdrawals per month"><PercentInput suffix="$" ariaLabel="Monthly withdrawals" value={withdrawals} onChange={setWithdrawals} /></Field>
            <Field label="Fees & withdrawals cash must cover" hint="Workbook cell BH2 (annual)."><div className="rounded-md border border-slate-200 bg-slate-50 px-2 py-1.5 text-right text-sm tabular-nums">{fmt$(annualFees)}</div></Field>
            <Field label="Add / remove money"><PercentInput suffix="$" ariaLabel="Add remove" value={addRemove} onChange={setAddRemove} /></Field>
            <Field label="Cash total override"><PercentInput suffix="$" ariaLabel="Cash total override" value={cashOverride} onChange={setCashOverride} /></Field>
          </div>
          <label className="mt-3 flex items-center gap-2 text-sm"><input type="checkbox" checked={updateButton} onChange={(e) => setUpdateButton(e.target.checked)} /> Also press the workbook's “Update Sheet” button</label>
          <div className="mt-4 flex gap-2">
            <Button variant="primary" disabled={!imported || busy || customInvalid} onClick={() => run()}>{busy ? 'Calculating…' : 'Calculate recommendations'}</Button>
            {result && <Button onClick={exportCsv}>Export CSV</Button>}
          </div>
          {customInvalid && <p className="mt-2 text-xs text-red-700">Custom growth and cash must be between 0% and 100% and total at most 100%.</p>}
          {runError && <div className="mt-2"><Notice tone="danger">{runError}</Notice></div>}
        </Card>
      </div>

      {!result ? (
        <EmptyState title="No recommendation yet">Import holdings, choose the client's settings and calculate. Use the synthetic sample to try it.</EmptyState>
      ) : (
        <Results result={result} portfolioDate={date} importedNames={new Set(Object.values(imported?.result.bySection ?? {}).flat().map((h) => h[0]))} overrides={overrides} onOverride={(name, v) => { const ov = { ...overrides, [name]: v }; setOverrides(ov); run(ov); }} />
      )}
    </>
  );
}

function Results({ result, portfolioDate, importedNames, overrides, onOverride }: { result: RecommendationResult; portfolioDate: string; importedNames: Set<string>; overrides: Record<string, number | null>; onOverride: (name: string, v: number | null) => void }) {
  const total = n(result.finalPortfolioTotal) || n(result.portfolioTotal);
  // Term deposits / bonds maturing within the 90-day hold window go to $0 in the workbook: label them.
  const holdUntil = new Date(`${portfolioDate}T00:00:00`);
  holdUntil.setDate(holdUntil.getDate() + 90);
  const matures = (name: string) => {
    const m = /(\d{2})\/(\d{2})\/(\d{4})$/.exec(name.trim());
    return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) <= holdUntil : false;
  };
  return (
    <>
      <div className="mb-6 grid grid-cols-2 gap-3 md:grid-cols-4">
        <Tile label="Current portfolio" value={fmt$(result.portfolioTotal)} />
        <Tile label="Recommended portfolio" value={fmt$(result.recommendedTotal)} />
        <Tile label="Checks" value={result.checks.length === 0 ? 'No errors' : `${result.checks.length} issue(s)`} tone={result.checks.length ? 'bad' : 'ok'} />
        <Tile label="Calculated in" value={`${(result.ms / 1000).toFixed(1)} s`} />
      </div>
      {result.checks.length > 0 && (
        <div className="mb-6"><Notice tone="danger"><strong>Workbook checks (BD2):</strong><ul className="ml-4 list-disc">{result.checks.map((c, i) => <li key={i}>{c}</li>)}</ul></Notice></div>
      )}
      <Card title="Sector summary" className="mb-6">
        <TableWrap>
          <thead><tr><Th>Sector</Th><Th align="right">Current</Th><Th align="right">Current %</Th><Th align="right">Target</Th><Th align="right">Target %</Th><Th align="right">Recommended</Th><Th align="right">Recommended %</Th><Th align="right">Change</Th></tr></thead>
          <tbody>
            {result.sections.map((s) => {
              const change = n(s.recommended) - n(s.current);
              return (
                <tr key={s.key}>
                  <Td className="font-medium">{s.label}</Td>
                  <Td align="right">{fmt$(s.current)}</Td>
                  <Td align="right">{fmtPct(total ? (n(s.current) / n(result.portfolioTotal)) * 100 : null)}</Td>
                  <Td align="right">{fmt$(s.target)}</Td>
                  <Td align="right">{fmtPct(total ? (n(s.target) / total) * 100 : null)}</Td>
                  <Td align="right" className="font-medium">{fmt$(s.recommended)}</Td>
                  <Td align="right">{fmtPct(total ? (n(s.recommended) / total) * 100 : null)}</Td>
                  <Td align="right" className={cx(change > 0 && 'text-blue-700', change < 0 && 'text-red-700')}>{change === 0 ? '—' : `${change > 0 ? '+' : '−'}${fmt$(Math.abs(change))}`}</Td>
                </tr>
              );
            })}
          </tbody>
        </TableWrap>
      </Card>
      {result.sections.filter((s) => s.rows.length).map((s) => (
        <Card key={s.key} title={`${s.label} · target ${fmt$(s.target)}`} className="mb-6">
          <TableWrap>
            <thead><tr><Th>Holding</Th><Th>Type</Th><Th align="right">Current</Th><Th align="right">Recommended</Th><Th align="right">Change</Th><Th align="right">Rec. %</Th><Th align="right">Override ±</Th></tr></thead>
            <tbody>
              {s.rows.map((r) => {
                const ch = n(r.change);
                const added = !importedNames.has(r.name);
                return (
                  <tr key={r.row} className={cx(ch !== 0 && 'bg-slate-50/60')}>
                    <Td className="max-w-md">{r.name}{added && <span className="ml-2 rounded bg-sky-50 px-1 text-[11px] text-sky-800">model</span>}</Td>
                    <Td className="whitespace-nowrap text-xs text-slate-600">{(s.key === 'fi' ? [r.info] : [r.type, r.info]).filter((x) => typeof x === 'string' && x && x !== '-').join(' · ')}</Td>
                    <Td align="right">{fmt$(r.value)}</Td>
                    <Td align="right" className="font-medium">{fmt$(r.recommended)}</Td>
                    <Td align="right" className={cx(ch > 0 && 'text-blue-700', ch < 0 && 'text-red-700')}>{ch === 0 ? '—' : `${ch > 0 ? 'Buy ' : s.key === 'fi' && n(r.recommended) === 0 && matures(r.name) ? 'Matures ' : 'Sell '}${fmt$(Math.abs(ch))}`}</Td>
                    <Td align="right">{fmtPct(r.recommendedPct, 2)}</Td>
                    <Td align="right"><OverrideInput value={overrides[r.name] ?? null} onCommit={(v) => onOverride(r.name, v)} /></Td>
                  </tr>
                );
              })}
              <tr className="bg-slate-50 font-semibold">
                <Td>Total</Td><Td /><Td align="right">{fmt$(s.current)}</Td><Td align="right">{fmt$(s.recommended)}</Td><Td align="right">{fmt$(n(s.recommended) - n(s.current))}</Td><Td /><Td />
              </tr>
            </tbody>
          </TableWrap>
        </Card>
      ))}
      <p className="mb-6 text-xs text-slate-500">
        Overrides work like the workbook's column G: the amount is added to the current value (0 = keep as is). “model” marks holdings the workbook added (model funds, bond and term-deposit ladder rows) that the client does not currently hold.
      </p>
    </>
  );
}

function OverrideInput({ value, onCommit }: { value: number | null; onCommit: (v: number | null) => void }) {
  const [t, setT] = useState(value === null ? '' : String(value));
  return (
    <input
      aria-label="Override"
      className="w-24 rounded border border-slate-300 px-1.5 py-0.5 text-right text-sm"
      value={t}
      placeholder="—"
      onChange={(e) => setT(e.target.value)}
      onBlur={() => {
        const m = parseMoney(t);
        if (!m.ok) return setT(value === null ? '' : String(value));
        if (m.value !== value) onCommit(m.value);
      }}
      onKeyDown={(e) => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur(); }}
    />
  );
}

function Tile({ label, value, tone }: { label: string; value: string; tone?: 'ok' | 'bad' }) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white px-4 py-3 shadow-sm">
      <div className="text-xs font-medium text-slate-500">{label}</div>
      <div className={cx('mt-1 text-xl font-semibold tabular-nums', tone === 'bad' ? 'text-red-700' : 'text-slate-900')}>{value}</div>
    </div>
  );
}
