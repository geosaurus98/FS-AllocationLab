import { useState, type ReactNode } from 'react';
import { PercentInput } from '../components/PercentInput';
import { Button, Card, Field, inputCls, IssueList, Notice, PageHeader, StatusBadge, Td, Th, TableWrap, ValidationBadge, cx } from '../components/ui';
import { sumWeights } from '../lib/allocation';
import { formatDate, formatPct } from '../lib/format';
import { ASSET_CLASSES, PROFILES, profile } from '../lib/reference';
import { cloneBounds } from '../lib/scenarios';
import type { BaselineAllocation, Bounds, ConstraintSet, DataStatus, ProfileId, Settings, Weights } from '../lib/types';
import { validateBounds } from '../lib/validation';
import { nowIso, useApp } from '../state/AppContext';
import { loadFixtureState } from '../state/fixtures';
import { evaluate, evaluateBaseline } from '../state/selectors';
import { baselineLockReason, findConstraints } from '../state/state';

type EditableStatus = Exclude<DataStatus, 'approved'>;
const EDITABLE_STATUSES: EditableStatus[] = ['draft', 'illustrative', 'synthetic', 'missing'];

export function Assumptions() {
  const { state } = useApp();
  const [editing, setEditing] = useState<string | null>(null);
  const [approving, setApproving] = useState<string | null>(null);
  const [editingCons, setEditingCons] = useState<ProfileId | null>(null);

  return (
    <>
      <PageHeader
        title="Assumptions"
        subtitle="Reference data with explicit status, source and version. Every change creates a new version and keeps the previous one in history. Approved records are locked."
      />

      <Card title="Baseline allocations (strategic & target)" className="mb-6">
        <TableWrap>
          <thead>
            <tr>
              <Th>Profile</Th>
              <Th>Kind</Th>
              <Th>Status</Th>
              <Th align="right">Ver.</Th>
              <Th>Source</Th>
              <Th align="right">Total</Th>
              <Th>Checks</Th>
              <Th>Updated</Th>
              <Th align="right">Actions</Th>
            </tr>
          </thead>
          <tbody>
            {state.baselines.map((b) => {
              const ev = evaluateBaseline(state, b);
              const lock = baselineLockReason(b);
              return (
                <BaselineRows key={b.id} b={b} editing={editing === b.id} approving={approving === b.id} onClose={() => { setEditing(null); setApproving(null); }}>
                  <Td className="whitespace-nowrap font-medium">{profile(b.profileId).label}</Td>
                  <Td className="capitalize">{b.kind}</Td>
                  <Td><StatusBadge status={b.status} /></Td>
                  <Td align="right">v{b.version}</Td>
                  <Td className="max-w-xs text-xs text-slate-600">{b.source}{b.notes && <div className="italic">{b.notes}</div>}</Td>
                  <Td align="right">{b.status === 'missing' ? '—' : formatPct(ev.total.total)}</Td>
                  <Td>{b.status === 'missing' ? <span className="text-xs text-slate-400">no data</span> : <span title={ev.validation.issues.map((i) => i.message).join('\n')}><ValidationBadge status={ev.validation.status} count={ev.validation.issues.length} /></span>}</Td>
                  <Td className="whitespace-nowrap text-xs text-slate-600">{formatDate(b.updatedAt)}</Td>
                  <Td align="right">
                    <div className="flex justify-end gap-1">
                      <Button variant="ghost" disabled={Boolean(lock)} title={lock ?? 'Create a new version of this baseline'} onClick={() => { setEditing(b.id); setApproving(null); }}>{lock ? 'Locked' : 'Revise'}</Button>
                      {b.status !== 'approved' && b.status !== 'missing' && <Button variant="ghost" onClick={() => { setApproving(b.id); setEditing(null); }}>Mark approved…</Button>}
                    </div>
                  </Td>
                </BaselineRows>
              );
            })}
          </tbody>
        </TableWrap>
      </Card>

      <Card title="Constraint sets (min / max per asset class)" className="mb-6">
        <TableWrap>
          <thead>
            <tr>
              <Th>Profile</Th>
              <Th>Status</Th>
              <Th align="right">Ver.</Th>
              {ASSET_CLASSES.map((a) => <Th key={a.id} align="right">{a.shortLabel}</Th>)}
              <Th>Source</Th>
              <Th align="right">Actions</Th>
            </tr>
          </thead>
          <tbody>
            {state.constraints.map((c) => (
              <ConstraintRows key={c.id} c={c} editing={editingCons === c.profileId} onEdit={() => setEditingCons(c.profileId)} onClose={() => setEditingCons(null)} />
            ))}
          </tbody>
        </TableWrap>
        <p className="mt-2 text-xs text-slate-500">“not set” means no limit is configured. The minimum Cash weight is the Cash minimum. Scenarios copy these limits when created; later edits here do not alter existing scenarios.</p>
      </Card>

      <div className="mb-6 grid gap-6 lg:grid-cols-2">
        <MultiSectorCard />
        <SettingsCard settings={state.settings} />
      </div>

      <HistoryCard />
    </>
  );
}

function BaselineRows({ b, editing, approving, onClose, children }: { b: BaselineAllocation; editing: boolean; approving: boolean; onClose: () => void; children: ReactNode }) {
  return (
    <>
      <tr className={cx('hover:bg-slate-50', (editing || approving) && 'bg-sky-50/50')}>{children}</tr>
      {editing && (
        <tr>
          <td colSpan={9} className="border-b border-slate-200 bg-slate-50 p-4"><BaselineEditor b={b} onClose={onClose} /></td>
        </tr>
      )}
      {approving && (
        <tr>
          <td colSpan={9} className="border-b border-slate-200 bg-slate-50 p-4"><ApproveForm b={b} onClose={onClose} /></td>
        </tr>
      )}
    </>
  );
}

function BaselineEditor({ b, onClose }: { b: BaselineAllocation; onClose: () => void }) {
  const { state, dispatch } = useApp();
  const [weights, setWeights] = useState<Weights>({ ...b.weights });
  const [status, setStatus] = useState<EditableStatus>(b.status === 'approved' ? 'draft' : b.status === 'missing' ? 'draft' : b.status);
  const [source, setSource] = useState(b.status === 'missing' ? '' : b.source);
  const [notes, setNotes] = useState('');
  const cons = findConstraints(state, b.profileId);
  const ev = evaluate(state, b.profileId, weights, cons?.bounds ?? null, state.multiSector.growthFraction);
  const total = sumWeights(weights).total;

  return (
    <div>
      <Notice tone="info">
        You are revising the <strong>{b.kind}</strong> baseline for <strong>{profile(b.profileId).label}</strong> (currently v{b.version}, {b.status}). Saving creates v{b.version + 1}; v{b.version} is kept in history. Revisions cannot be saved as “approved”.
      </Notice>
      <div className="mt-3 grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
        {ASSET_CLASSES.map((ac) => (
          <Field key={ac.id} label={ac.label}>
            <PercentInput ariaLabel={`${ac.label} baseline weight`} value={weights[ac.id]} onChange={(v) => setWeights((w) => ({ ...w, [ac.id]: v }))} />
          </Field>
        ))}
      </div>
      <div className="mt-2 text-sm">Total: <strong className="tabular-nums">{formatPct(total)}</strong></div>
      <div className="mt-3 grid gap-3 sm:grid-cols-3">
        <Field label="Status">
          <select className={inputCls} value={status} onChange={(e) => setStatus(e.target.value as EditableStatus)}>
            {EDITABLE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Source (required)"><input className={inputCls} value={source} onChange={(e) => setSource(e.target.value)} placeholder="e.g. Investment committee paper, draft" /></Field>
        <Field label="Change note"><input className={inputCls} value={notes} onChange={(e) => setNotes(e.target.value)} /></Field>
      </div>
      <div className="mt-3"><IssueList issues={ev.validation.issues} /></div>
      <p className="mt-2 text-xs text-slate-500">Validation issues do not block saving a draft, but they are shown wherever this baseline is used.</p>
      <div className="mt-3 flex gap-2">
        <Button
          variant="primary"
          disabled={!source.trim()}
          onClick={() => {
            dispatch({ type: 'baseline/revise', now: nowIso(), revision: { profileId: b.profileId, kind: b.kind, weights, status, source: source.trim(), notes: notes.trim(), effectiveDate: null } });
            onClose();
          }}
        >
          Save as v{b.version + 1}
        </Button>
        <Button onClick={onClose}>Cancel</Button>
      </div>
    </div>
  );
}

function ApproveForm({ b, onClose }: { b: BaselineAllocation; onClose: () => void }) {
  const { state, dispatch } = useApp();
  const [note, setNote] = useState('');
  const [typed, setTyped] = useState('');
  const ev = evaluateBaseline(state, b);
  const blocking = ev.validation.status === 'invalid' || ev.validation.status === 'incomplete';
  return (
    <div className="space-y-3">
      <Notice tone="warning">
        Marking a baseline as <strong>approved</strong> records that it has been approved through the firm's governance process outside this tool. This prototype has no approval workflow or authentication. Once approved, the record is locked: it cannot be revised or overwritten by import.
      </Notice>
      {blocking && <Notice tone="danger">This allocation currently fails validation ({ev.validation.status}). Resolve the issues before approving.</Notice>}
      <div className="grid gap-3 sm:grid-cols-2">
        <Field label="Approval reference (required)"><input className={inputCls} value={note} onChange={(e) => setNote(e.target.value)} placeholder="e.g. IC minutes ref / date" /></Field>
        <Field label='Type "APPROVE" to confirm'><input className={inputCls} value={typed} onChange={(e) => setTyped(e.target.value)} /></Field>
      </div>
      <div className="flex gap-2">
        <Button variant="primary" disabled={blocking || !note.trim() || typed !== 'APPROVE'} onClick={() => { dispatch({ type: 'baseline/approve', profileId: b.profileId, kind: b.kind, note: note.trim(), now: nowIso() }); onClose(); }}>Mark approved</Button>
        <Button onClick={onClose}>Cancel</Button>
      </div>
    </div>
  );
}

function ConstraintRows({ c, editing, onEdit, onClose }: { c: ConstraintSet; editing: boolean; onEdit: () => void; onClose: () => void }) {
  const fmtBound = (min: number | null, max: number | null) => {
    if (min === null && max === null) return <span className="text-slate-400">not set</span>;
    if (max === null) return `≥ ${formatPct(min, 1)}`;
    if (min === null) return `≤ ${formatPct(max, 1)}`;
    return `${formatPct(min, 1)} – ${formatPct(max, 1)}`;
  };
  return (
    <>
      <tr className={cx('hover:bg-slate-50', editing && 'bg-sky-50/50')}>
        <Td className="whitespace-nowrap font-medium">{profile(c.profileId).label}</Td>
        <Td><StatusBadge status={c.status} /></Td>
        <Td align="right">v{c.version}</Td>
        {ASSET_CLASSES.map((a) => <Td key={a.id} align="right" className="whitespace-nowrap text-xs">{fmtBound(c.bounds[a.id].min, c.bounds[a.id].max)}</Td>)}
        <Td className="max-w-[14rem] text-xs text-slate-600">{c.source}</Td>
        <Td align="right"><Button variant="ghost" disabled={c.status === 'approved'} onClick={onEdit}>{c.status === 'approved' ? 'Locked' : 'Revise'}</Button></Td>
      </tr>
      {editing && (
        <tr>
          <td colSpan={11} className="border-b border-slate-200 bg-slate-50 p-4"><ConstraintEditor c={c} onClose={onClose} /></td>
        </tr>
      )}
    </>
  );
}

function ConstraintEditor({ c, onClose }: { c: ConstraintSet; onClose: () => void }) {
  const { dispatch } = useApp();
  const [bounds, setBounds] = useState<Bounds>(cloneBounds(c.bounds));
  const [status, setStatus] = useState<EditableStatus>(c.status === 'approved' || c.status === 'missing' ? 'draft' : c.status);
  const [source, setSource] = useState(c.status === 'missing' ? '' : c.source);
  const issues = validateBounds(bounds);
  return (
    <div>
      <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
        {ASSET_CLASSES.map((ac) => (
          <div key={ac.id} className="text-sm">
            <div className="mb-1 font-medium text-slate-700">{ac.label}{ac.id === 'cash' && <span className="font-normal text-slate-500"> (min Cash)</span>}</div>
            <div className="flex gap-1">
              <PercentInput ariaLabel={`${ac.label} minimum`} placeholder="min" value={bounds[ac.id].min} onChange={(v) => setBounds((b) => ({ ...b, [ac.id]: { ...b[ac.id], min: v } }))} />
              <PercentInput ariaLabel={`${ac.label} maximum`} placeholder="max" value={bounds[ac.id].max} onChange={(v) => setBounds((b) => ({ ...b, [ac.id]: { ...b[ac.id], max: v } }))} />
            </div>
          </div>
        ))}
      </div>
      <div className="mt-3 grid gap-3 sm:grid-cols-2">
        <Field label="Status">
          <select className={inputCls} value={status} onChange={(e) => setStatus(e.target.value as EditableStatus)}>
            {EDITABLE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Source (required)"><input className={inputCls} value={source} onChange={(e) => setSource(e.target.value)} /></Field>
      </div>
      {issues.length > 0 && <div className="mt-3"><IssueList issues={issues} /></div>}
      <div className="mt-3 flex gap-2">
        <Button variant="primary" disabled={issues.length > 0 || !source.trim()} onClick={() => { dispatch({ type: 'constraints/revise', profileId: c.profileId, bounds, status, source: source.trim(), now: nowIso() }); onClose(); }}>Save as v{c.version + 1}</Button>
        <Button onClick={onClose}>Cancel</Button>
      </div>
    </div>
  );
}

function MultiSectorCard() {
  const { state, dispatch } = useApp();
  const ms = state.multiSector;
  const [pct, setPct] = useState<number | null>(ms.growthFraction === null ? null : Math.round(ms.growthFraction * 1e6) / 1e4);
  const [status, setStatus] = useState<EditableStatus>(ms.status === 'approved' ? 'draft' : ms.status);
  const [source, setSource] = useState(ms.source);
  const invalid = pct !== null && (pct < 0 || pct > 100);
  const changed = (pct === null ? null : pct / 100) !== ms.growthFraction || status !== ms.status || source !== ms.source;
  return (
    <Card title={<span className="flex items-center gap-2">Multi-Sector look-through <StatusBadge status={ms.status} /> <span className="text-xs font-normal text-slate-500">v{ms.version}</span></span>}>
      <p className="mb-3 text-sm text-slate-600">
        Share of Multi-Sector exposure treated as <em>growth</em>; the remainder is treated as defensive. Leave blank if unknown — Multi-Sector weight is then reported as <strong>unclassified</strong> and excluded from the growth/defensive split.
      </p>
      <div className="grid gap-3 sm:grid-cols-3">
        <Field label="Growth fraction"><PercentInput ariaLabel="Multi-Sector growth fraction" placeholder="not set" value={pct} invalid={invalid} onChange={setPct} /></Field>
        <Field label="Status">
          <select className={inputCls} value={status} onChange={(e) => setStatus(e.target.value as EditableStatus)}>
            {EDITABLE_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Source"><input className={inputCls} value={source} onChange={(e) => setSource(e.target.value)} /></Field>
      </div>
      {invalid && <p className="mt-2 text-sm text-red-700">Must be between 0% and 100%.</p>}
      <div className="mt-3">
        <Button variant="primary" disabled={!changed || invalid || !source.trim()} onClick={() => dispatch({ type: 'multiSector/revise', growthFraction: pct === null ? null : pct / 100, status, source: source.trim(), now: nowIso() })}>Save as v{ms.version + 1}</Button>
      </div>
    </Card>
  );
}

function SettingsCard({ settings }: { settings: Settings }) {
  const { dispatch } = useApp();
  const [tol, setTol] = useState<number | null>(settings.sumTolerance);
  const [ntol, setNtol] = useState<number | null>(settings.nominalSplitTolerance);
  const valid = tol !== null && tol >= 0 && tol <= 5 && ntol !== null && ntol >= 0 && ntol <= 20;
  const changed = tol !== settings.sumTolerance || ntol !== settings.nominalSplitTolerance;
  return (
    <Card title="Validation settings">
      <div className="grid gap-3 sm:grid-cols-2">
        <Field label="Sum-to-100 tolerance" hint="Percentage points, 0–5. Default 0.01."><PercentInput ariaLabel="Sum tolerance" suffix="pp" value={tol} onChange={setTol} /></Field>
        <Field label="Nominal split tolerance" hint="pp difference allowed between actual growth and a profile's named split (e.g. 60/40)."><PercentInput ariaLabel="Nominal split tolerance" suffix="pp" value={ntol} onChange={setNtol} /></Field>
      </div>
      <div className="mt-3">
        <Button variant="primary" disabled={!valid || !changed} onClick={() => valid && dispatch({ type: 'settings/update', settings: { sumTolerance: tol, nominalSplitTolerance: ntol }, now: nowIso() })}>Save settings</Button>
      </div>
    </Card>
  );
}

function HistoryCard() {
  const { state, dispatch } = useApp();
  const [confirmReset, setConfirmReset] = useState(false);
  return (
    <Card
      title="Version history & audit log"
      actions={
        confirmReset ? (
          <>
            <span className="self-center text-xs text-red-700">Discard all scenarios and revisions?</span>
            <Button variant="danger" onClick={() => { dispatch({ type: 'state/replace', state: loadFixtureState() }); setConfirmReset(false); }}>Reset everything</Button>
            <Button onClick={() => setConfirmReset(false)}>Cancel</Button>
          </>
        ) : (
          <Button variant="danger" onClick={() => setConfirmReset(true)}>Reset demo data…</Button>
        )
      }
    >
      {state.auditLog.length === 0 ? (
        <p className="text-sm text-slate-500">No changes to reference data yet.</p>
      ) : (
        <ul className="max-h-72 space-y-1 overflow-y-auto text-sm">
          {[...state.auditLog].reverse().map((e, i) => (
            <li key={i} className="flex gap-3 border-b border-slate-100 py-1">
              <span className="w-40 shrink-0 text-xs tabular-nums text-slate-500">{new Date(e.at).toLocaleString('en-NZ')}</span>
              <span className="font-medium">{e.action}</span>
              <span className="text-slate-600">{e.detail}</span>
            </li>
          ))}
        </ul>
      )}
      <p className="mt-3 text-xs text-slate-500">
        Superseded versions retained: {state.baselineHistory.length} baseline, {state.constraintHistory.length} constraint, {state.multiSectorHistory.length} Multi-Sector. Profiles: {PROFILES.length}.
      </p>
    </Card>
  );
}
