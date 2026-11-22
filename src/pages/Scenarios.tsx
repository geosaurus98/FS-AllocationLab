import { useState } from 'react';
import { SplitBar, StackedAllocationChart } from '../components/charts';
import { NewScenarioForm } from '../components/NewScenarioForm';
import { Button, Card, EmptyState, HypotheticalBadge, PageHeader, Td, Th, TableWrap, ValidationBadge, cx } from '../components/ui';
import { diffWeights } from '../lib/allocation';
import { formatDate, formatPct, formatPp } from '../lib/format';
import { ASSET_CLASSES, profile } from '../lib/reference';
import { hasUnsavedChanges } from '../lib/scenarios';
import { href, navigate } from '../router';
import { newId, nowIso, useApp } from '../state/AppContext';
import { evaluateScenario } from '../state/selectors';

const MAX_COMPARE = 6;

export function Scenarios() {
  const { state, dispatch } = useApp();
  const [selected, setSelected] = useState<string[]>([]);
  const [confirmDelete, setConfirmDelete] = useState<string | null>(null);

  const sel = selected.map((id) => state.scenarios.find((s) => s.id === id)).filter((s): s is NonNullable<typeof s> => Boolean(s));
  const toggle = (id: string) => setSelected((cur) => (cur.includes(id) ? cur.filter((x) => x !== id) : cur.length >= MAX_COMPARE ? cur : [...cur, id]));
  const first = sel[0];

  return (
    <>
      <PageHeader
        title="Scenarios"
        subtitle="Hypothetical proposals. Each scenario holds its own copy of weights and constraints, so creating, editing or deleting scenarios never changes the baselines."
      />

      <Card title="New scenario" className="mb-6">
        <NewScenarioForm />
      </Card>

      <Card title={`Scenario workspace (${state.scenarios.length})`} className="mb-6" actions={selected.length > 0 ? <Button variant="ghost" onClick={() => setSelected([])}>Clear selection</Button> : undefined}>
        {state.scenarios.length === 0 ? (
          <EmptyState title="No scenarios yet">Create a scenario above — for example, start from the illustrative Balanced 60/40 strategic allocation — then edit it in the Allocation Designer.</EmptyState>
        ) : (
          <TableWrap>
            <thead>
              <tr>
                <Th><span className="sr-only">Compare</span></Th>
                <Th>Name</Th>
                <Th>Profile</Th>
                <Th>Origin</Th>
                <Th align="right">Total</Th>
                <Th>Growth / defensive</Th>
                <Th>Validation</Th>
                <Th>Saved</Th>
                <Th align="right">Actions</Th>
              </tr>
            </thead>
            <tbody>
              {state.scenarios.map((s) => {
                const ev = evaluateScenario(state, s, 'saved');
                const unsaved = hasUnsavedChanges(s);
                return (
                  <tr key={s.id} className={cx('hover:bg-slate-50', selected.includes(s.id) && 'bg-sky-50/60')}>
                    <Td>
                      <input type="checkbox" aria-label={`Select ${s.name} for comparison`} checked={selected.includes(s.id)} onChange={() => toggle(s.id)} disabled={!selected.includes(s.id) && selected.length >= MAX_COMPARE} />
                    </Td>
                    <Td className="font-medium">
                      <a className="hover:underline" href={href({ page: 'designer', scenarioId: s.id })}>{s.name}</a>
                      {unsaved && <span className="ml-2 rounded bg-amber-100 px-1 py-0.5 text-[11px] font-medium text-amber-900">unsaved edits</span>}
                      {s.description && <div className="text-xs font-normal text-slate-500">{s.description}</div>}
                    </Td>
                    <Td className="whitespace-nowrap">{profile(s.profileId).label}</Td>
                    <Td className="text-xs text-slate-600">{s.origin.baselineKind ? `${s.origin.baselineKind} v${s.origin.baselineVersion}` : 'blank / import'}{s.origin.duplicatedFromScenarioId ? ' · copy' : ''}</Td>
                    <Td align="right">{formatPct(ev.total.total)}</Td>
                    <Td><div className="w-32"><SplitBar compact growth={ev.split.growth} defensive={ev.split.defensive} unclassified={ev.split.unclassified} /><div className="mt-0.5 text-xs tabular-nums text-slate-600">{formatPct(ev.split.growth, 1)} / {formatPct(ev.split.defensive, 1)}</div></div></Td>
                    <Td><span title={ev.validation.issues.map((i) => i.message).join('\n')}><ValidationBadge status={ev.validation.status} count={ev.validation.issues.length} /></span></Td>
                    <Td className="whitespace-nowrap text-xs text-slate-600">{formatDate(s.savedAt)}</Td>
                    <Td align="right">
                      <div className="flex justify-end gap-1">
                        <Button variant="ghost" onClick={() => navigate({ page: 'designer', scenarioId: s.id })}>Edit</Button>
                        <Button variant="ghost" onClick={() => dispatch({ type: 'scenario/duplicate', id: s.id, newId: newId('scn'), now: nowIso() })}>Duplicate</Button>
                        {confirmDelete === s.id ? (
                          <>
                            <Button variant="danger" onClick={() => { dispatch({ type: 'scenario/delete', id: s.id }); setSelected((c) => c.filter((x) => x !== s.id)); setConfirmDelete(null); }}>Confirm delete</Button>
                            <Button variant="ghost" onClick={() => setConfirmDelete(null)}>Cancel</Button>
                          </>
                        ) : (
                          <Button variant="ghost" className="text-red-700" onClick={() => setConfirmDelete(s.id)}>Delete</Button>
                        )}
                      </div>
                    </Td>
                  </tr>
                );
              })}
            </tbody>
          </TableWrap>
        )}
        {state.scenarios.length > 0 && <p className="mt-2 text-xs text-slate-500">Tick up to {MAX_COMPARE} scenarios to compare them below. Comparison uses each scenario's saved state.</p>}
      </Card>

      {sel.length > 0 && first && (
        <Card title={<span className="flex items-center gap-2">Scenario comparison <HypotheticalBadge /></span>}>
          <StackedAllocationChart rows={sel.map((s) => ({ key: s.id, label: s.name.length > 22 ? `${s.name.slice(0, 21)}…` : s.name, weights: s.saved.weights, tag: profile(s.profileId).label }))} />
          <div className="mt-6">
            <TableWrap>
              <thead>
                <tr>
                  <Th>Asset class</Th>
                  {sel.map((s, i) => <Th key={s.id} align="right">{s.name}{i > 0 && <div className="font-normal normal-case text-slate-500">Δ vs first</div>}</Th>)}
                </tr>
              </thead>
              <tbody>
                {ASSET_CLASSES.map((ac) => (
                  <tr key={ac.id}>
                    <Td className="font-medium">{ac.label}</Td>
                    {sel.map((s, i) => {
                      const v = s.saved.weights[ac.id];
                      const dv = i > 0 ? diffWeights(first.saved.weights, s.saved.weights)[ac.id] : null;
                      return (
                        <Td key={s.id} align="right">
                          {v === null ? <span className="text-orange-700">missing</span> : formatPct(v)}
                          {i > 0 && <div className={cx('text-xs', dv !== null && dv > 0 && 'text-blue-700', dv !== null && dv < 0 && 'text-red-700')}>{formatPp(dv)}</div>}
                        </Td>
                      );
                    })}
                  </tr>
                ))}
                {(['growth', 'defensive', 'unclassified'] as const).map((k) => (
                  <tr key={k} className="bg-slate-50">
                    <Td className="font-medium capitalize">{k}</Td>
                    {sel.map((s) => <Td key={s.id} align="right">{formatPct(evaluateScenario(state, s, 'saved').split[k])}</Td>)}
                  </tr>
                ))}
              </tbody>
            </TableWrap>
          </div>
        </Card>
      )}
    </>
  );
}
