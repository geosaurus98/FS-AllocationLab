import { useState } from 'react';
import { StackedAllocationChart, SplitBar } from '../components/charts';
import { NewScenarioForm } from '../components/NewScenarioForm';
import { PercentInput } from '../components/PercentInput';
import { Button, Card, EmptyState, Field, HypotheticalBadge, inputCls, IssueList, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, ValidationBadge, cx } from '../components/ui';
import { diffWeights } from '../lib/allocation';
import { formatDate, formatPct, formatPp } from '../lib/format';
import { ASSET_CLASSES, profile } from '../lib/reference';
import { hasUnsavedChanges } from '../lib/scenarios';
import type { BaselineKind, Scenario } from '../lib/types';
import { navigate, href } from '../router';
import { newId, nowIso, useApp } from '../state/AppContext';
import { evaluateScenario } from '../state/selectors';
import { findBaseline, findScenario } from '../state/state';

export function Designer({ scenarioId }: { scenarioId: string | null }) {
  const { state } = useApp();
  const scenario = scenarioId ? findScenario(state, scenarioId) : undefined;

  if (!scenario) {
    return (
      <>
        <PageHeader title="Allocation Designer" subtitle="Edit a hypothetical scenario's weights and constraints. Approved and reference baselines are never edited here." />
        {scenarioId && <div className="mb-4"><Notice tone="warning">Scenario “{scenarioId}” was not found. It may have been deleted.</Notice></div>}
        {state.scenarios.length > 0 && (
          <Card title="Open an existing scenario" className="mb-6">
            <ScenarioPicker current={null} />
          </Card>
        )}
        <Card title="Create a scenario to design">
          <NewScenarioForm />
        </Card>
      </>
    );
  }
  return <DesignerBody scenario={scenario} />;
}

function ScenarioPicker({ current }: { current: string | null }) {
  const { state } = useApp();
  if (state.scenarios.length === 0) return <EmptyState title="No scenarios yet" />;
  return (
    <select className={cx(inputCls, 'max-w-md')} value={current ?? ''} onChange={(e) => navigate({ page: 'designer', scenarioId: e.target.value })} aria-label="Select scenario">
      {current === null && <option value="">Select a scenario…</option>}
      {state.scenarios.map((s) => (
        <option key={s.id} value={s.id}>{s.name} — {profile(s.profileId).label}{hasUnsavedChanges(s) ? ' (unsaved)' : ''}</option>
      ))}
    </select>
  );
}

function DesignerBody({ scenario }: { scenario: Scenario }) {
  const { state, dispatch } = useApp();
  const [compareKind, setCompareKind] = useState<BaselineKind>('strategic');
  const p = profile(scenario.profileId);
  const ev = evaluateScenario(state, scenario);
  const unsaved = hasUnsavedChanges(scenario);
  const ref = findBaseline(state, scenario.profileId, compareKind);
  const refUsable = ref && ref.status !== 'missing';
  const diff = refUsable ? diffWeights(ref.weights, scenario.draft.weights) : null;
  const issuesByAsset = new Set(ev.validation.issues.filter((i) => i.severity === 'error').map((i) => i.assetClass));
  const msOverride = scenario.draft.multiSectorGrowthFraction;
  const totalOk = !ev.validation.issues.some((i) => i.code === 'SUM_NOT_100');
  const remaining = 100 - ev.total.total;

  const id = scenario.id;
  const now = nowIso;

  return (
    <>
      <PageHeader
        title="Allocation Designer"
        subtitle={<>Editing a <strong>hypothetical</strong> scenario. Changes are drafts until saved, and are never written back to baselines. Nothing is rebalanced automatically.</>}
        actions={
          <>
            <Button onClick={() => dispatch({ type: 'scenario/revert', id })} disabled={!unsaved}>Revert to saved</Button>
            <Button
              onClick={() => {
                const nid = newId('scn');
                dispatch({ type: 'scenario/duplicate', id, newId: nid, now: now() });
                navigate({ page: 'designer', scenarioId: nid });
              }}
            >Duplicate</Button>
            <Button variant="primary" onClick={() => dispatch({ type: 'scenario/save', id, now: now() })} disabled={!unsaved}>Save scenario</Button>
          </>
        }
      />

      <div className="mb-4 flex flex-wrap items-center gap-3">
        <ScenarioPicker current={id} />
        <HypotheticalBadge />
        {unsaved ? <span className="rounded bg-amber-100 px-1.5 py-0.5 text-xs font-medium text-amber-900">Unsaved changes</span> : <span className="text-xs text-slate-500">Saved {formatDate(scenario.savedAt)}</span>}
        <a className="text-sm text-slate-700 underline" href={href({ page: 'compare', profileId: scenario.profileId, scenarioId: id })}>Compare with baselines →</a>
      </div>

      <div className="grid gap-6 xl:grid-cols-[minmax(0,1fr)_320px]">
        <div className="space-y-6">
          <Card
            title={<span>Weights &amp; constraints · {p.label}</span>}
            actions={
              <div className="flex flex-wrap items-center gap-2 text-sm">
                <label className="flex items-center gap-1 text-slate-600">
                  Compare to
                  <select className="rounded border border-slate-300 px-1.5 py-1 text-sm" value={compareKind} onChange={(e) => setCompareKind(e.target.value as BaselineKind)}>
                    <option value="strategic">Strategic</option>
                    <option value="target">Target</option>
                  </select>
                </label>
                <Button
                  disabled={!refUsable}
                  title={refUsable ? `Replace draft weights with the ${compareKind} baseline (v${ref?.version})` : 'No baseline data to reset to'}
                  onClick={() => dispatch({ type: 'scenario/resetToBaseline', id, kind: compareKind, includeConstraints: true })}
                >
                  Reset to {compareKind}
                </Button>
              </div>
            }
          >
            <TableWrap>
              <thead>
                <tr>
                  <Th>Asset class</Th>
                  <Th align="right">Weight</Th>
                  <Th align="right">Min</Th>
                  <Th align="right">Max</Th>
                  <Th align="right">{compareKind === 'strategic' ? 'Strategic' : 'Target'}</Th>
                  <Th align="right">Δ vs {compareKind}</Th>
                </tr>
              </thead>
              <tbody>
                {ASSET_CLASSES.map((ac) => {
                  const b = scenario.draft.bounds[ac.id];
                  const d = diff?.[ac.id];
                  return (
                    <tr key={ac.id} className={issuesByAsset.has(ac.id) ? 'bg-red-50/50' : ''}>
                      <Td className="whitespace-nowrap font-medium">{ac.label}<div className="text-[11px] font-normal text-slate-500">{ac.classification === 'lookThrough' ? 'Look-through' : ac.classification === 'growth' ? 'Growth' : 'Defensive'}</div></Td>
                      <Td align="right">
                        <PercentInput className="ml-auto w-24" ariaLabel={`${ac.label} weight`} value={scenario.draft.weights[ac.id]} invalid={issuesByAsset.has(ac.id)} onChange={(v) => dispatch({ type: 'scenario/setWeight', id, assetClass: ac.id, value: v })} />
                      </Td>
                      <Td align="right">
                        <PercentInput className="ml-auto w-20" ariaLabel={`${ac.label} minimum`} placeholder="none" value={b.min} onChange={(v) => dispatch({ type: 'scenario/setBound', id, assetClass: ac.id, bound: { ...b, min: v } })} />
                      </Td>
                      <Td align="right">
                        <PercentInput className="ml-auto w-20" ariaLabel={`${ac.label} maximum`} placeholder="none" value={b.max} onChange={(v) => dispatch({ type: 'scenario/setBound', id, assetClass: ac.id, bound: { ...b, max: v } })} />
                      </Td>
                      <Td align="right" className="text-slate-600">{refUsable ? (ref.weights[ac.id] === null ? <span className="text-orange-700">missing</span> : formatPct(ref.weights[ac.id])) : '—'}</Td>
                      <Td align="right" className={cx(d && d > 0 && 'text-blue-700', d && d < 0 && 'text-red-700')}>{diff ? formatPp(d) : '—'}</Td>
                    </tr>
                  );
                })}
                <tr className="bg-slate-50 font-semibold">
                  <Td>Total</Td>
                  <Td align="right" className={totalOk ? 'text-emerald-700' : 'text-red-700'}>{formatPct(ev.total.total)}</Td>
                  <Td colSpan={2} className="text-xs font-normal text-slate-600">
                    {ev.total.complete ? (totalOk ? 'Sums to 100% ✓' : `${remaining > 0 ? 'Unallocated' : 'Over-allocated'}: ${formatPp(Math.abs(remaining), false)}`) : `${ev.total.missing.length} weight(s) missing`}
                  </Td>
                  <Td align="right" className="text-slate-600">{refUsable ? formatPct(Object.values(ref.weights).reduce<number>((a, v) => a + (v ?? 0), 0)) : '—'}</Td>
                  <Td />
                </tr>
              </tbody>
            </TableWrap>
            {!refUsable && <p className="mt-2 text-xs text-slate-500">The {compareKind} baseline for {p.label} has no data, so differences cannot be shown.</p>}
            <p className="mt-2 text-xs text-slate-500">Blank = missing (not zero). Constraint limits here are scenario-specific copies; baseline constraints are edited under Assumptions. Sum tolerance ±{state.settings.sumTolerance} pp.</p>
          </Card>

          <Card title="Allocation chart">
            <StackedAllocationChart
              rows={[
                ...(refUsable ? [{ key: 'ref', label: `${compareKind === 'strategic' ? 'Strategic' : 'Target'} v${ref.version}`, weights: ref.weights, tag: ref.status }] : []),
                { key: 'saved', label: 'Scenario (saved)', weights: scenario.saved.weights, tag: 'hypothetical' },
                ...(unsaved ? [{ key: 'draft', label: 'Scenario (draft)', weights: scenario.draft.weights, tag: 'hypothetical · unsaved' }] : []),
              ]}
            />
          </Card>
        </div>

        <div className="space-y-6">
          <Card title={<span className="flex items-center gap-2">Validation <ValidationBadge status={ev.validation.status} /></span>}>
            <IssueList issues={ev.validation.issues} emptyText="All checks pass." />
          </Card>

          <Card title="Growth / defensive split">
            <SplitBar growth={ev.split.growth} defensive={ev.split.defensive} unclassified={ev.split.unclassified} />
            <dl className="mt-3 grid grid-cols-2 gap-y-1 text-sm">
              <dt className="text-slate-600">Growth</dt><dd className="text-right tabular-nums">{formatPct(ev.split.growth)}</dd>
              <dt className="text-slate-600">Defensive</dt><dd className="text-right tabular-nums">{formatPct(ev.split.defensive)}</dd>
              <dt className="text-slate-600">Unclassified</dt><dd className={cx('text-right tabular-nums', ev.split.unclassified > 0 && 'text-amber-800')}>{formatPct(ev.split.unclassified)}</dd>
              {p.nominalGrowth !== null && (<><dt className="text-slate-600">Profile label implies</dt><dd className="text-right tabular-nums">{p.nominalGrowth}/{100 - p.nominalGrowth}</dd></>)}
            </dl>
            <div className="mt-4 border-t border-slate-100 pt-3">
              <Field
                label="Multi-Sector look-through"
                hint={
                  msOverride === 'inherit'
                    ? `Using global assumption: ${state.multiSector.growthFraction === null ? 'not set (unclassified)' : `${(state.multiSector.growthFraction * 100).toFixed(1)}% growth`} · status ${state.multiSector.status}`
                    : 'Scenario-specific override.'
                }
              >
                <div className="flex gap-2">
                  <select
                    className={inputCls}
                    value={msOverride === 'inherit' ? 'inherit' : msOverride === null ? 'unclassified' : 'override'}
                    onChange={(e) => {
                      const v = e.target.value;
                      dispatch({ type: 'scenario/setMsFraction', id, value: v === 'inherit' ? 'inherit' : v === 'unclassified' ? null : (state.multiSector.growthFraction ?? 0.5) });
                    }}
                  >
                    <option value="inherit">Use global assumption</option>
                    <option value="override">Override for this scenario</option>
                    <option value="unclassified">Treat as unclassified</option>
                  </select>
                  {typeof msOverride === 'number' && (
                    <PercentInput
                      className="w-28 shrink-0"
                      ariaLabel="Multi-Sector growth percentage"
                      value={Math.round(msOverride * 1e6) / 1e4}
                      onChange={(v) => dispatch({ type: 'scenario/setMsFraction', id, value: v === null ? null : v / 100 })}
                    />
                  )}
                </div>
              </Field>
              {typeof msOverride === 'number' && (msOverride < 0 || msOverride > 1) && <p className="mt-1 text-xs text-red-700">Growth fraction must be between 0% and 100%; it is being treated as unclassified.</p>}
            </div>
          </Card>

          <Card title="Scenario details">
            <ScenarioMeta scenario={scenario} />
            <dl className="mt-3 space-y-1 text-xs text-slate-600">
              <div>Profile: <strong>{p.label}</strong></div>
              <div>Created: {formatDate(scenario.createdAt)}</div>
              <div>
                Origin:{' '}
                {scenario.origin.baselineId
                  ? `${scenario.origin.baselineKind} baseline v${scenario.origin.baselineVersion}`
                  : 'blank / import'}
                {scenario.origin.duplicatedFromScenarioId && ' · duplicated from another scenario'}
              </div>
              {ref && <div className="flex items-center gap-1">Current {compareKind} baseline: <StatusBadge status={ref.status} /> v{ref.version}</div>}
            </dl>
          </Card>
        </div>
      </div>
    </>
  );
}

function ScenarioMeta({ scenario }: { scenario: Scenario }) {
  const { dispatch } = useApp();
  const [name, setName] = useState(scenario.name);
  const [desc, setDesc] = useState(scenario.description);
  const [lastId, setLastId] = useState(scenario.id);
  if (lastId !== scenario.id) {
    setLastId(scenario.id);
    setName(scenario.name);
    setDesc(scenario.description);
  }
  const dirty = name !== scenario.name || desc !== scenario.description;
  return (
    <div className="space-y-2">
      <Field label="Name"><input className={inputCls} value={name} onChange={(e) => setName(e.target.value)} /></Field>
      <Field label="Notes"><textarea className={inputCls} rows={2} value={desc} onChange={(e) => setDesc(e.target.value)} /></Field>
      <Button disabled={!dirty || !name.trim()} onClick={() => dispatch({ type: 'scenario/meta', id: scenario.id, name, description: desc })}>Update details</Button>
    </div>
  );
}
