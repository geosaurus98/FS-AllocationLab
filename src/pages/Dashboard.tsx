import { StackedAllocationChart, SplitBar } from '../components/charts';
import { Card, PageHeader, StatusBadge, Td, Th, TableWrap, ValidationBadge, Notice } from '../components/ui';
import { formatPct } from '../lib/format';
import { PROFILES } from '../lib/reference';
import { hasUnsavedChanges } from '../lib/scenarios';
import { href } from '../router';
import { useApp } from '../state/AppContext';
import { evaluateBaseline, evaluateScenario } from '../state/selectors';
import { findBaseline, findConstraints } from '../state/state';

export function Dashboard() {
  const { state } = useApp();

  const rows = PROFILES.map((p) => {
    const strategic = findBaseline(state, p.id, 'strategic');
    const target = findBaseline(state, p.id, 'target');
    const cons = findConstraints(state, p.id);
    return {
      p,
      strategic,
      target,
      cons,
      sEval: strategic ? evaluateBaseline(state, strategic) : null,
      tEval: target ? evaluateBaseline(state, target) : null,
    };
  });

  const counts = {
    approved: state.baselines.filter((b) => b.status === 'approved').length,
    missing: state.baselines.filter((b) => b.status === 'missing').length,
    invalid: rows.filter((r) => r.sEval && ['invalid', 'incomplete'].includes(r.sEval.validation.status)).length,
    scenarios: state.scenarios.length,
    unsaved: state.scenarios.filter(hasUnsavedChanges).length,
  };
  const scenarioIssues = state.scenarios.filter((s) => ['invalid', 'incomplete'].includes(evaluateScenario(state, s).validation.status)).length;

  return (
    <>
      <PageHeader
        title="Dashboard"
        subtitle="Overview of the standard profiles: baseline allocations, their provenance status and validation results."
      />

      {counts.approved === 0 && (
        <div className="mb-4">
          <Notice tone="info">No allocations in this workspace are marked <strong>Approved</strong>. All baseline figures shown are synthetic or illustrative, and are not company-approved allocations or recommendations.</Notice>
        </div>
      )}

      <div className="mb-6 grid grid-cols-2 gap-3 md:grid-cols-4">
        <Stat label="Profiles" value={PROFILES.length} />
        <Stat label="Baselines missing data" value={counts.missing} sub={`of ${state.baselines.length} strategic/target records`} />
        <Stat label="Strategic allocations failing checks" value={counts.invalid} tone={counts.invalid ? 'bad' : 'ok'} />
        <Stat label="Scenarios" value={counts.scenarios} sub={`${counts.unsaved} with unsaved edits · ${scenarioIssues} failing checks`} />
      </div>

      <Card title="Profiles — strategic allocation summary" className="mb-6">
        <TableWrap>
          <thead>
            <tr>
              <Th>Profile</Th>
              <Th>Strategic</Th>
              <Th align="right">Cash</Th>
              <Th align="right">Fixed Int.</Th>
              <Th align="right">Aus. Eq.</Th>
              <Th align="right">Global Eq.</Th>
              <Th align="right">Property</Th>
              <Th align="right">Multi-Sector</Th>
              <Th align="right">Total</Th>
              <Th>Growth / defensive</Th>
              <Th>Validation</Th>
              <Th>Target</Th>
              <Th>Constraints</Th>
            </tr>
          </thead>
          <tbody>
            {rows.map(({ p, strategic, target, cons, sEval, tEval }) => (
              <tr key={p.id} className="hover:bg-slate-50">
                <Td className="font-medium">
                  <a className="text-slate-900 underline-offset-2 hover:underline" href={href({ page: 'compare', profileId: p.id, scenarioId: null })}>{p.label}</a>
                </Td>
                <Td>{strategic ? <span className="flex items-center gap-1 whitespace-nowrap"><StatusBadge status={strategic.status} /><span className="text-xs text-slate-500">v{strategic.version}</span></span> : '—'}</Td>
                {strategic && (['cash', 'fixedInterest', 'ausEquities', 'globalEquities', 'property', 'multiSector'] as const).map((id) => (
                  <Td key={id} align="right" className={strategic.weights[id] === null ? 'text-orange-700' : ''}>{strategic.weights[id] === null ? 'missing' : formatPct(strategic.weights[id], 1)}</Td>
                ))}
                <Td align="right" className="font-medium">{sEval ? formatPct(sEval.total.total, 1) : '—'}</Td>
                <Td>
                  {sEval && (
                    <div className="w-28">
                      <SplitBar compact growth={sEval.split.growth} defensive={sEval.split.defensive} unclassified={sEval.split.unclassified} />
                      <div className="mt-0.5 text-xs tabular-nums text-slate-600">
                        {formatPct(sEval.split.growth, 1)} / {formatPct(sEval.split.defensive, 1)}
                        {sEval.split.unclassified > 0 && <span className="text-amber-800"> · {formatPct(sEval.split.unclassified, 1)} unclass.</span>}
                      </div>
                    </div>
                  )}
                </Td>
                <Td>{sEval && <span title={sEval.validation.issues.map((i) => i.message).join('\n')}><ValidationBadge status={sEval.validation.status} count={sEval.validation.issues.length} /></span>}</Td>
                <Td>{target && <span className="flex items-center gap-1 whitespace-nowrap"><StatusBadge status={target.status} />{tEval && target.status !== 'missing' && <ValidationBadge status={tEval.validation.status} />}</span>}</Td>
                <Td>{cons && <StatusBadge status={cons.status} />}</Td>
              </tr>
            ))}
          </tbody>
        </TableWrap>
        <p className="mt-3 text-xs text-slate-500">Hover a validation badge for details. Profile names link to the comparison view.</p>
      </Card>

      <Card title="Strategic allocations by profile">
        <StackedAllocationChart
          rows={rows.filter((r) => r.strategic && r.strategic.status !== 'missing').map((r) => ({ key: r.p.id, label: r.p.label, weights: r.strategic!.weights, tag: r.strategic!.status }))}
        />
      </Card>
    </>
  );
}

function Stat({ label, value, sub, tone }: { label: string; value: number; sub?: string; tone?: 'ok' | 'bad' }) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white px-4 py-3 shadow-sm">
      <div className="text-xs font-medium text-slate-500">{label}</div>
      <div className={`mt-1 text-2xl font-semibold tabular-nums ${tone === 'bad' ? 'text-red-700' : 'text-slate-900'}`}>{value}</div>
      {sub && <div className="mt-0.5 text-xs text-slate-500">{sub}</div>}
    </div>
  );
}
