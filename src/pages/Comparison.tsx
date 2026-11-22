import { DiffChart, StackedAllocationChart } from '../components/charts';
import { Card, EmptyState, HypotheticalBadge, inputCls, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, ValidationBadge, cx, Field } from '../components/ui';
import { diffWeights, type WeightDiff } from '../lib/allocation';
import { formatPct, formatPp } from '../lib/format';
import { ASSET_CLASSES, PROFILES, profile } from '../lib/reference';
import type { ProfileId, Weights } from '../lib/types';
import { PROFILE_IDS } from '../lib/types';
import { href, navigate } from '../router';
import { useApp } from '../state/AppContext';
import { evaluate, evaluateBaseline, evaluateScenario, type Evaluation } from '../state/selectors';
import { findBaseline, findScenario } from '../state/state';

function isProfileId(v: string | null): v is ProfileId {
  return v !== null && (PROFILE_IDS as readonly string[]).includes(v);
}

export function Comparison({ profileId: rawProfile, scenarioId }: { profileId: string | null; scenarioId: string | null }) {
  const { state } = useApp();
  const profileId: ProfileId = isProfileId(rawProfile) ? rawProfile : 'balanced6040';
  const p = profile(profileId);
  const strategic = findBaseline(state, profileId, 'strategic');
  const target = findBaseline(state, profileId, 'target');
  const profileScenarios = state.scenarios.filter((s) => s.profileId === profileId);
  const scenario = scenarioId ? findScenario(state, scenarioId) : undefined;
  const scenarioValid = scenario && scenario.profileId === profileId ? scenario : undefined;

  const has = (w?: { status: string }) => Boolean(w && w.status !== 'missing');
  const sW = has(strategic) ? strategic!.weights : null;
  const tW = has(target) ? target!.weights : null;
  const pW = scenarioValid ? scenarioValid.saved.weights : null;

  const sEv = strategic && sW ? evaluateBaseline(state, strategic) : null;
  const tEv = target && tW ? evaluateBaseline(state, target) : null;
  const pEv = scenarioValid ? evaluate(state, profileId, scenarioValid.saved.weights, scenarioValid.saved.bounds, evaluateScenario(state, scenarioValid, 'saved').fraction) : null;

  const d = (a: Weights | null, b: Weights | null): WeightDiff | null => (a && b ? diffWeights(a, b) : null);
  const dTS = d(sW, tW);
  const dPS = d(sW, pW);
  const dPT = d(tW, pW);

  const go = (pid: ProfileId, sid: string | null) => navigate({ page: 'compare', profileId: pid, scenarioId: sid });

  return (
    <>
      <PageHeader title="Profile Comparison" subtitle="Strategic vs target vs proposed allocations for one profile, with percentage-point (pp) differences." />

      <Card className="mb-6">
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <Field label="Profile">
            <select className={inputCls} value={profileId} onChange={(e) => go(e.target.value as ProfileId, null)}>
              {PROFILES.map((x) => <option key={x.id} value={x.id}>{x.label}</option>)}
            </select>
          </Field>
          <Field label="Proposed (scenario, saved state)" hint={profileScenarios.length === 0 ? <>No scenarios for {p.label}. <a className="underline" href={href({ page: 'scenarios' })}>Create one</a>.</> : undefined}>
            <select className={inputCls} value={scenarioValid?.id ?? ''} onChange={(e) => go(profileId, e.target.value || null)} disabled={profileScenarios.length === 0}>
              <option value="">— none —</option>
              {profileScenarios.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
            </select>
          </Field>
        </div>
        {scenarioId && !scenarioValid && <div className="mt-3"><Notice tone="warning">The selected scenario does not belong to {p.label} or no longer exists.</Notice></div>}
      </Card>

      <Card title={`${p.label}: allocation comparison`} className="mb-6">
        <TableWrap>
          <thead>
            <tr>
              <Th>Asset class</Th>
              <Th align="right"><ColHead label="Strategic" status={strategic?.status} version={strategic?.version} /></Th>
              <Th align="right"><ColHead label="Target" status={target?.status} version={target?.version} /></Th>
              <Th align="right"><span className="flex flex-col items-end gap-1">Proposed {scenarioValid && <HypotheticalBadge />}</span></Th>
              <Th align="right">Target − Strategic</Th>
              <Th align="right">Proposed − Strategic</Th>
              <Th align="right">Proposed − Target</Th>
            </tr>
          </thead>
          <tbody>
            {ASSET_CLASSES.map((ac) => (
              <tr key={ac.id}>
                <Td className="whitespace-nowrap font-medium">{ac.label}</Td>
                <WCell w={sW} id={ac.id} />
                <WCell w={tW} id={ac.id} />
                <WCell w={pW} id={ac.id} />
                <DCell d={dTS} id={ac.id} />
                <DCell d={dPS} id={ac.id} />
                <DCell d={dPT} id={ac.id} />
              </tr>
            ))}
            <SummaryRow label="Total" s={sEv} t={tEv} pr={pEv} pick={(e) => e.total.total} />
            <SummaryRow label="Growth" s={sEv} t={tEv} pr={pEv} pick={(e) => e.split.growth} />
            <SummaryRow label="Defensive" s={sEv} t={tEv} pr={pEv} pick={(e) => e.split.defensive} />
            <SummaryRow label="Unclassified (Multi-Sector)" s={sEv} t={tEv} pr={pEv} pick={(e) => e.split.unclassified} />
            <tr>
              <Td className="font-medium">Validation</Td>
              {[sEv, tEv, pEv].map((e, i) => (
                <Td key={i} align="right">{e ? <span title={e.validation.issues.map((x) => x.message).join('\n')}><ValidationBadge status={e.validation.status} count={e.validation.issues.length} /></span> : <span className="text-slate-400">—</span>}</Td>
              ))}
              <Td colSpan={3} />
            </tr>
          </tbody>
        </TableWrap>
        <p className="mt-2 text-xs text-slate-500">
          Differences are shown only where both sides have data. “No data” means the allocation has not been supplied — it is not treated as zero.
          {p.nominalGrowth !== null && ` ${p.label} implies ${p.nominalGrowth}% growth / ${100 - p.nominalGrowth}% defensive.`}
        </p>
      </Card>

      <div className="grid gap-6 lg:grid-cols-2">
        <Card title="Stacked allocations">
          {sW || tW || pW ? (
            <StackedAllocationChart
              rows={[
                ...(sW ? [{ key: 's', label: 'Strategic', weights: sW, tag: strategic?.status }] : []),
                ...(tW ? [{ key: 't', label: 'Target', weights: tW, tag: target?.status }] : []),
                ...(pW ? [{ key: 'p', label: 'Proposed', weights: pW, tag: 'hypothetical' }] : []),
              ]}
            />
          ) : (
            <EmptyState title="Nothing to chart">No strategic, target or proposed data exists for {p.label}.</EmptyState>
          )}
        </Card>
        <Card title="Percentage-point differences">
          {dPS ? (
            <DiffChart diff={dPS} label="Proposed − Strategic" />
          ) : dTS ? (
            <DiffChart diff={dTS} label="Target − Strategic" />
          ) : (
            <EmptyState title="No comparison available">Select a scenario, or supply both strategic and target data for {p.label}, to see differences.</EmptyState>
          )}
        </Card>
      </div>
    </>
  );
}

function ColHead({ label, status, version }: { label: string; status?: Parameters<typeof StatusBadge>[0]['status']; version?: number }) {
  return (
    <span className="flex flex-col items-end gap-1">
      {label}
      {status && <span className="flex items-center gap-1 normal-case"><StatusBadge status={status} />{status !== 'missing' && <span className="text-[10px] text-slate-500">v{version}</span>}</span>}
    </span>
  );
}

function WCell({ w, id }: { w: Weights | null; id: keyof Weights }) {
  if (!w) return <Td align="right" className="text-slate-400">No data</Td>;
  const v = w[id];
  return <Td align="right" className={v === null ? 'text-orange-700' : ''}>{v === null ? 'missing' : formatPct(v)}</Td>;
}

function DCell({ d, id }: { d: WeightDiff | null; id: keyof Weights }) {
  const v = d?.[id] ?? null;
  return <Td align="right" className={cx(!d && 'text-slate-400', v !== null && v > 0 && 'text-blue-700', v !== null && v < 0 && 'text-red-700')}>{d ? formatPp(v) : '—'}</Td>;
}

function SummaryRow({ label, s, t, pr, pick }: { label: string; s: Evaluation | null; t: Evaluation | null; pr: Evaluation | null; pick: (e: Evaluation) => number }) {
  const vals = [s, t, pr].map((e) => (e ? pick(e) : null));
  const diff = (a: number | null, b: number | null) => (a === null || b === null ? null : b - a);
  const cells = [diff(vals[0]!, vals[1]!), diff(vals[0]!, vals[2]!), diff(vals[1]!, vals[2]!)];
  return (
    <tr className="bg-slate-50">
      <Td className="font-medium">{label}</Td>
      {vals.map((v, i) => <Td key={i} align="right" className="font-medium">{v === null ? <span className="text-slate-400">—</span> : formatPct(v)}</Td>)}
      {cells.map((c, i) => <Td key={i} align="right" className={cx(c !== null && c > 0 && 'text-blue-700', c !== null && c < 0 && 'text-red-700')}>{formatPp(c)}</Td>)}
    </tr>
  );
}
