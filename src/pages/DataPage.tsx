import { Fragment, useMemo, useState } from 'react';
import { Button, Card, inputCls, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, ValidationBadge } from '../components/ui';
import { CSV_COLUMNS, exportAllocationsCsv, parseAllocationCsv, planImport, toCsv, type CsvParseResult, type ImportAction, type ImportPlanItem } from '../lib/csv';
import { formatPct } from '../lib/format';
import { ASSET_CLASSES, emptyBounds, profile } from '../lib/reference';
import { createScenario } from '../lib/scenarios';
import { newId, nowIso, useApp } from '../state/AppContext';
import { evaluate } from '../state/selectors';
import { findConstraints } from '../state/state';

function download(filename: string, text: string) {
  const blob = new Blob([text], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

const stamp = () => new Date().toISOString().slice(0, 10);

const ACTION_LABEL: Record<ImportAction, string> = {
  'create-scenario': 'Import as new scenario (hypothetical)',
  'propose-baseline-revision': 'Apply as new baseline version (draft)',
  skip: 'Skip',
};

export function DataPage() {
  const { state, dispatch } = useApp();
  const [text, setText] = useState('');
  const [fileName, setFileName] = useState<string | null>(null);
  const [parsed, setParsed] = useState<CsvParseResult | null>(null);
  const [choices, setChoices] = useState<Record<number, ImportAction>>({});
  const [result, setResult] = useState<string | null>(null);

  const plan: ImportPlanItem[] = useMemo(() => (parsed ? planImport(parsed.rows, state.baselines) : []), [parsed, state.baselines]);
  const revisions = plan.filter((p) => (choices[p.row.line] ?? p.defaultAction) === 'propose-baseline-revision');
  const [confirmRevisions, setConfirmRevisions] = useState(false);

  const runParse = (t: string) => {
    setResult(null);
    setChoices({});
    setConfirmRevisions(false);
    setParsed(t.trim() ? parseAllocationCsv(t) : null);
  };

  const apply = () => {
    const now = nowIso();
    let scenarios = 0;
    let revised = 0;
    let skipped = 0;
    for (const item of plan) {
      const action = choices[item.row.line] ?? item.defaultAction;
      if (!item.options.includes(action)) continue; // defensive: never apply a disallowed action
      const r = item.row;
      if (action === 'create-scenario') {
        const cons = findConstraints(state, r.profileId);
        dispatch({
          type: 'scenario/add',
          scenario: createScenario({
            id: newId('scn'),
            name: r.name || `Imported ${profile(r.profileId).label} ${r.kind} (line ${r.line})`,
            description: `Imported from ${fileName ?? 'pasted CSV'} line ${r.line}${r.source ? ` · source: ${r.source}` : ''}`,
            profileId: r.profileId,
            now,
            initialWeights: r.weights,
            bounds: cons?.bounds ?? emptyBounds(),
          }),
        });
        scenarios++;
      } else if (action === 'propose-baseline-revision' && (r.kind === 'strategic' || r.kind === 'target')) {
        dispatch({
          type: 'baseline/revise',
          now,
          revision: {
            profileId: r.profileId,
            kind: r.kind,
            weights: r.weights,
            status: r.status === 'approved' || r.status === 'hypothetical' ? 'draft' : r.status,
            source: r.source || `CSV import (${fileName ?? 'pasted'})`,
            notes: `Imported from line ${r.line}`,
            effectiveDate: null,
          },
        });
        revised++;
      } else skipped++;
    }
    dispatch({ type: 'audit', now, action: 'CSV import', detail: `${fileName ?? 'pasted CSV'}: ${scenarios} scenario(s) created, ${revised} baseline revision(s), ${skipped} skipped, ${parsed?.errors.length ?? 0} invalid row(s) not imported` });
    setResult(`Import complete: ${scenarios} scenario(s) created, ${revised} baseline version(s) added, ${skipped} row(s) skipped, ${parsed?.errors.length ?? 0} invalid row(s) not imported.`);
    setParsed(null);
    setText('');
  };

  const template = toCsv([
    CSV_COLUMNS,
    ['scenario', 'balanced_60_40', 'proposed', 'Example proposal', 'hypothetical', 'Example', '', 2, 38, 17, 39, 4, 0],
    ['baseline', 'balanced_50_50', 'target', '', 'draft', 'Example source', '', 3, 47, 15, 31, 4, ''],
  ]);

  return (
    <>
      <PageHeader title="Import / Export" subtitle="CSV exchange of allocations. Imports are validated and previewed first; by default every row becomes a new hypothetical scenario, and approved values can never be overwritten." />

      <Card title="Export" className="mb-6">
        <div className="flex flex-wrap gap-2">
          <Button onClick={() => download(`allocationlab-allocations-${stamp()}.csv`, exportAllocationsCsv(state.baselines, state.scenarios))}>All baselines + scenarios (CSV)</Button>
          <Button onClick={() => download(`allocationlab-baselines-${stamp()}.csv`, exportAllocationsCsv(state.baselines, []))}>Baselines only</Button>
          <Button onClick={() => download(`allocationlab-scenarios-${stamp()}.csv`, exportAllocationsCsv([], state.scenarios))} disabled={state.scenarios.length === 0}>Scenarios only</Button>
          <Button variant="ghost" onClick={() => download('allocationlab-template.csv', template)}>Download template</Button>
        </div>
        <p className="mt-3 text-xs text-slate-500">
          Columns: <code className="rounded bg-slate-100 px-1">{CSV_COLUMNS.join(',')}</code>. Weights are percentages; blank cells mean missing (not zero). Exported scenarios use their saved state. Each row's <code>status</code> column carries its provenance label.
        </p>
      </Card>

      <Card title="Import" className="mb-6">
        {result && <div className="mb-3"><Notice tone="success" onDismiss={() => setResult(null)}>{result}</Notice></div>}
        <div className="grid gap-3 lg:grid-cols-[240px_1fr]">
          <div className="space-y-2">
            <label className="block text-sm font-medium text-slate-700">Choose a CSV file</label>
            <input
              type="file"
              accept=".csv,text/csv"
              className="block w-full text-sm file:mr-2 file:rounded file:border file:border-slate-300 file:bg-white file:px-2 file:py-1"
              onChange={async (e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                if (f.size > 1_000_000) { setParsed({ rows: [], errors: [], fatal: ['File is larger than 1 MB; this prototype only accepts small CSV files.'] }); return; }
                const t = await f.text();
                setFileName(f.name);
                setText(t);
                runParse(t);
                e.target.value = '';
              }}
            />
            <p className="text-xs text-slate-500">…or paste CSV text and click “Validate”.</p>
          </div>
          <div>
            <textarea className={`${inputCls} font-mono text-xs`} rows={6} value={text} placeholder={CSV_COLUMNS.join(',')} onChange={(e) => { setText(e.target.value); setFileName(null); }} />
            <div className="mt-2 flex gap-2">
              <Button onClick={() => runParse(text)} disabled={!text.trim()}>Validate</Button>
              {parsed && <Button variant="ghost" onClick={() => { setParsed(null); setText(''); }}>Clear</Button>}
            </div>
          </div>
        </div>

        {parsed && (
          <div className="mt-5 space-y-4">
            {parsed.fatal.length > 0 && <Notice tone="danger"><strong>File rejected — nothing will be imported.</strong><ul className="ml-4 list-disc">{parsed.fatal.map((f, i) => <li key={i}>{f}</li>)}</ul></Notice>}
            {parsed.errors.length > 0 && (
              <Notice tone="warning">
                <strong>{parsed.errors.length} row(s) are invalid and will not be imported:</strong>
                <ul className="ml-4 list-disc">{parsed.errors.map((e) => <li key={e.line}>Line {e.line}: {e.message}</li>)}</ul>
              </Notice>
            )}
            {plan.length > 0 && (
              <>
                <TableWrap>
                  <thead>
                    <tr>
                      <Th>Line</Th>
                      <Th>Type</Th>
                      <Th>Profile / kind</Th>
                      {ASSET_CLASSES.map((a) => <Th key={a.id} align="right">{a.shortLabel}</Th>)}
                      <Th>Checks</Th>
                      <Th>Existing</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {plan.map((item) => {
                      const r = item.row;
                      const cons = findConstraints(state, r.profileId);
                      const ev = evaluate(state, r.profileId, r.weights, cons?.bounds ?? null, state.multiSector.growthFraction);
                      const action = choices[r.line] ?? item.defaultAction;
                      return (
                        <Fragment key={r.line}>
                        <tr className="align-top">
                          <Td>{r.line}</Td>
                          <Td className="capitalize">{r.recordType}</Td>
                          <Td className="whitespace-nowrap">{profile(r.profileId).label}<div className="text-xs text-slate-500">{r.kind}{r.name && ` · ${r.name}`}</div></Td>
                          {ASSET_CLASSES.map((a) => {
                            const v = r.weights[a.id];
                            const changed = item.existing && item.existing.weights[a.id] !== v;
                            return <Td key={a.id} align="right" className={changed ? 'font-semibold text-blue-800' : v === null ? 'text-orange-700' : ''}>{v === null ? 'missing' : formatPct(v)}</Td>;
                          })}
                          <Td><span title={ev.validation.issues.map((i) => i.message).join('\n')}><ValidationBadge status={ev.validation.status} count={ev.validation.issues.length} /></span></Td>
                          <Td>{item.existing ? <span className="flex items-center gap-1"><StatusBadge status={item.existing.status} /><span className="text-xs">v{item.existing.version}</span></span> : '—'}</Td>
                        </tr>
                        <tr>
                          <td colSpan={11} className="border-b border-slate-200 px-2.5 pb-3">
                            <div className="flex flex-col gap-2 rounded-md bg-slate-50 p-2 md:flex-row md:items-start">
                            <select aria-label={`Import action for line ${r.line}`} className={`${inputCls} md:w-80 md:shrink-0`} value={action} onChange={(e) => { setChoices((c) => ({ ...c, [r.line]: e.target.value as ImportAction })); setConfirmRevisions(false); }}>
                              {item.options.map((o) => <option key={o} value={o}>{ACTION_LABEL[o]}</option>)}
                            </select>
                            <ul className="space-y-0.5 text-xs text-slate-600">{item.notes.map((n, i) => <li key={i}>{n}</li>)}</ul>
                            </div>
                          </td>
                        </tr>
                        </Fragment>
                      );
                    })}
                  </tbody>
                </TableWrap>
                <p className="text-xs text-slate-500">Values in bold differ from the current baseline. Validation issues do not block import, but they will be shown wherever the data is used.</p>

                {revisions.length > 0 && (
                  <Notice tone="warning">
                    <label className="flex items-start gap-2">
                      <input type="checkbox" className="mt-1" checked={confirmRevisions} onChange={(e) => setConfirmRevisions(e.target.checked)} />
                      <span>I confirm that {revisions.length} row(s) will create <strong>new baseline versions</strong> ({revisions.map((r) => `${profile(r.row.profileId).label} ${r.row.kind}`).join(', ')}). Previous versions are kept in history. Approved baselines are not affected.</span>
                    </label>
                  </Notice>
                )}
                <div className="flex gap-2">
                  <Button variant="primary" disabled={revisions.length > 0 && !confirmRevisions} onClick={apply}>Import {plan.filter((p) => (choices[p.row.line] ?? p.defaultAction) !== 'skip').length} row(s)</Button>
                  <Button onClick={() => { setParsed(null); }}>Cancel</Button>
                </div>
              </>
            )}
            {plan.length === 0 && parsed.fatal.length === 0 && <p className="text-sm text-slate-600">No valid rows to import.</p>}
          </div>
        )}
      </Card>
    </>
  );
}
