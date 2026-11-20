import { useState } from 'react';
import { emptyBounds, PROFILES } from '../lib/reference';
import { createScenario } from '../lib/scenarios';
import type { ProfileId } from '../lib/types';
import { navigate } from '../router';
import { newId, nowIso, useApp } from '../state/AppContext';
import { findBaseline, findConstraints } from '../state/state';
import { Button, Field, inputCls } from './ui';

type Start = 'strategic' | 'target' | 'blank';

export function NewScenarioForm({ defaultProfile = 'balanced6040', onCreated }: { defaultProfile?: ProfileId; onCreated?: (id: string) => void }) {
  const { state, dispatch } = useApp();
  const [profileId, setProfileId] = useState<ProfileId>(defaultProfile);
  const [start, setStart] = useState<Start>('strategic');
  const [name, setName] = useState('');

  const base = start === 'blank' ? null : findBaseline(state, profileId, start);
  const baseMissing = start !== 'blank' && (!base || base.status === 'missing');

  const create = () => {
    const id = newId('scn');
    const cons = findConstraints(state, profileId);
    const scenario = createScenario({
      id,
      name,
      profileId,
      now: nowIso(),
      baseline: baseMissing ? null : base,
      bounds: cons?.bounds ?? emptyBounds(),
    });
    dispatch({ type: 'scenario/add', scenario });
    setName('');
    if (onCreated) onCreated(id);
    else navigate({ page: 'designer', scenarioId: id });
  };

  return (
    <div className="grid gap-3 sm:grid-cols-[1fr_1fr_1fr_auto] sm:items-end">
      <Field label="Profile">
        <select className={inputCls} value={profileId} onChange={(e) => setProfileId(e.target.value as ProfileId)}>
          {PROFILES.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
        </select>
      </Field>
      <Field label="Start from" hint={baseMissing ? 'No data for this baseline — the scenario will start with blank (missing) weights.' : undefined}>
        <select className={inputCls} value={start} onChange={(e) => setStart(e.target.value as Start)}>
          <option value="strategic">Strategic baseline (copy)</option>
          <option value="target">Target baseline (copy)</option>
          <option value="blank">Blank — all weights missing</option>
        </select>
      </Field>
      <Field label="Scenario name">
        <input className={inputCls} value={name} placeholder="e.g. Lower Fixed Interest test" onChange={(e) => setName(e.target.value)} />
      </Field>
      <Button variant="primary" onClick={create}>Create scenario</Button>
    </div>
  );
}
