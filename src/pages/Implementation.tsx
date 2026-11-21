import { useMemo, useState } from 'react';
import { Button, Card, Field, inputCls, Notice, Td, Th, TableWrap, cx } from '../components/ui';
import { MAX_ROUNDS, type ImplementationRow, type RoundAmounts, type SectionKey } from '../engine/workbook/run';
import { computerCheck, parseApexTransactions, type CheckResult } from '../engine/workbook/transactions';
import { BLANK, isErr, type Scalar } from '../engine/workbook/values';
import { readXlsxSheet } from '../lib/xlsx';
import { detectDelimiter, parseCsv, toCsv } from '../lib/csv';
import { parseMoney } from '../lib/holdings';

/** Implementation advice: trade rounds, checks, trade instructions and the computer check. */

export interface RoundsState {
  amounts: RoundAmounts;
  dates: (string | null)[];
  /** Three sign-off initials per holding per round: checks[name][round] = [c1, c2, c3]. */
  checks: Record<string, string[][]>;
  count: number;
  current: number;
}

export const emptyRounds = (): RoundsState => ({ amounts: {}, dates: [new Date().toISOString().slice(0, 10)], checks: {}, count: 1, current: 0 });

const SECTION_LABEL: Record<SectionKey, string> = { cash: 'Cash', fi: 'Fixed Interest', ae: 'Australasian Equities', ge: 'Global Equities', prop: 'Property', ms: 'Multi Sector' };
const n = (v: Scalar) => (typeof v === 'number' ? v : 0);
const fmt$ = (v: Scalar | number | null) =>
  typeof v === 'number' ? v.toLocaleString('en-NZ', { style: 'currency', currency: 'NZD', maximumFractionDigits: 0 }) : v === null || v === BLANK || v === '' ? '—' : isErr(v) ? v.code : String(v);
const txt = (v: Scalar) => (typeof v === 'string' && v !== '-' ? v : typeof v === 'number' ? String(v) : '');
/** Brokerage and spreads are stored as fractions (0.0025 = 0.25%); larger numbers are dollar amounts. */
const rate = (v: Scalar) => {
  if (typeof v !== 'number') return txt(v);
  if (Math.abs(v) >= 1) return fmt$(v);
  return `${Number((v * 100).toFixed(3))}%`;
};

function download(name: string, text: string, type = 'text/csv') {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function Implementation(props: {
  rows: ImplementationRow[];
  checks: string[];
  exposure: Record<string, Scalar>;
  custodian: 'Apex' | 'OA';
  portfolioId: string;
  rounds: RoundsState;
  onRounds: (r: RoundsState) => void;
  dataNames: { name: string; assetClass: string }[];
  onAddHolding: (section: SectionKey, name: string) => void;
  /** Holdings added during implementation — always shown. */
  addedNames: string[];
  busy: boolean;
}) {
  const { rows, rounds, onRounds, custodian } = props;
  const [showAll, setShowAll] = useState(false);
  const [openRationale, setOpenRationale] = useState<string | null>(null);
  const [addSection, setAddSection] = useState<SectionKey>('fi');
  const [addName, setAddName] = useState('');
  const [check, setCheck] = useState<CheckResult | null>(null);
  const [checkError, setCheckError] = useState<string | null>(null);
  const [checkFile, setCheckFile] = useState<string | null>(null);
  const cur = rounds.current;

  const visible = rows.filter((r) => showAll || r.section === 'cash' || n(r.change) !== 0 || props.addedNames.includes(r.name) || (rounds.amounts[r.name] ?? []).some((x) => x));
  const setAmount = (name: string, i: number, v: number | null) => {
    const a = [...(rounds.amounts[name] ?? [])];
    while (a.length < MAX_ROUNDS) a.push(null);
    a[i] = v;
    onRounds({ ...rounds, amounts: { ...rounds.amounts, [name]: a } });
  };
  const setCheckInitials = (name: string, k: number, v: string) => {
    const c = (rounds.checks[name] ?? []).map((x) => [...x]);
    while (c.length <= cur) c.push(['', '', '']);
    c[cur]![k] = v;
    onRounds({ ...rounds, checks: { ...rounds.checks, [name]: c } });
  };
  const roundTotal = (i: number, section?: SectionKey) =>
    rows.filter((r) => !section || r.section === section).reduce((a, r) => a + (rounds.amounts[r.name]?.[i] ?? 0), 0);

  // trade instructions for the current round (sells first, then buys)
  const instructions = useMemo(() => {
    const list = rows
      .filter((r) => r.section !== 'cash')
      .map((r) => ({ r, amt: rounds.amounts[r.name]?.[cur] ?? 0 }))
      .filter((x) => x.amt !== 0);
    return [...list.filter((x) => x.amt < 0), ...list.filter((x) => x.amt > 0)];
  }, [rows, rounds, cur]);
  const broker = (r: ImplementationRow) => (custodian === 'Apex' ? [txt(r.apexBroker), rate(r.apexBrokerage)] : [txt(r.oaBroker), rate(r.oaBrokerage)]);
  const cname = (r: ImplementationRow) => (custodian === 'Apex' ? r.apexName : r.oaName) || r.name;

  const exportInstructions = () => {
    const out: (string | number | null)[][] = [['Round', 'Date', 'Action', 'Holding (FS name)', `${custodian} name`, 'Amount', 'Broker', 'Brokerage']];
    for (const { r, amt } of instructions) out.push([cur + 1, rounds.dates[cur] ?? '', amt < 0 ? 'Sell' : 'Buy', r.name, cname(r), Math.abs(amt), ...broker(r)]);
    download(`trade-instructions-round-${cur + 1}-${rounds.dates[cur] ?? ''}.csv`, toCsv(out));
  };

  const runCheck = async (f: File) => {
    setCheckError(null);
    setCheck(null);
    setCheckFile(f.name);
    try {
      if (custodian !== 'Apex') throw new Error('The computer check is available for Apex reports. OneAnswer reports are not supported yet.');
      let grid: unknown[][];
      if (/\.(xlsx|xlsm)$/i.test(f.name)) grid = (await readXlsxSheet(await f.arrayBuffer())).rows.map((r) => r.map((c) => (c !== '' && /^-?\d+(\.\d+)?$/.test(c) ? Number(c) : c)));
      else {
        const t = await f.text();
        grid = parseCsv(t, detectDelimiter(t));
      }
      const report = parseApexTransactions(grid, props.portfolioId);
      setCheck(computerCheck({ holdings: rows.map((r) => ({ name: r.name, apexName: r.apexName, oaName: r.oaName, type: txt(r.type) })), rounds: rounds.amounts, roundDates: rounds.dates, currentRound: cur, report, custodian }));
    } catch (e) {
      setCheckError(e instanceof Error ? e.message : String(e));
    }
  };

  const namesForSection = props.dataNames.filter((d) => {
    const map: Record<SectionKey, string[]> = { cash: ['cash'], fi: ['fixed interest'], ae: ['australasian equities'], ge: ['global equities'], prop: ['property'], ms: ['diversified', 'multi sector'] };
    return map[addSection].includes(d.assetClass.toLowerCase());
  });

  return (
    <>
      <Card title="Rounds" className="mb-6" actions={props.busy ? <span className="text-xs text-slate-500">Updating…</span> : undefined}>
        <div className="flex flex-wrap items-end gap-3">
          {Array.from({ length: rounds.count }, (_, i) => (
            <div key={i} className={cx('rounded-md border px-2 py-1.5', i === cur ? 'border-fuchsia-400 bg-fuchsia-50' : 'border-slate-200')}>
              <button type="button" className="text-sm font-semibold" onClick={() => onRounds({ ...rounds, current: i })}>Round {i + 1}</button>
              <input
                type="date"
                aria-label={`Round ${i + 1} date`}
                className="ml-2 rounded border border-slate-300 px-1 text-xs"
                value={rounds.dates[i] ?? ''}
                onChange={(e) => {
                  const d = [...rounds.dates];
                  d[i] = e.target.value || null;
                  onRounds({ ...rounds, dates: d });
                }}
              />
              <div className="text-[11px] text-slate-500">Net {fmt$(roundTotal(i))}</div>
            </div>
          ))}
          <Button
            disabled={rounds.count >= MAX_ROUNDS}
            title={rounds.count >= MAX_ROUNDS ? 'The workbook has six round columns' : undefined}
            onClick={() => onRounds({ ...rounds, count: rounds.count + 1, current: rounds.count, dates: [...rounds.dates.slice(0, rounds.count), new Date().toISOString().slice(0, 10)] })}
          >
            Next round
          </Button>
          <Button
            variant="ghost"
            onClick={() => {
              const amounts = { ...rounds.amounts };
              for (const r of rows) {
                if (r.section === 'cash') continue;
                const rem = n(r.remaining);
                if (rem !== 0 && !(amounts[r.name]?.[cur])) {
                  const a = [...(amounts[r.name] ?? [])];
                  while (a.length < MAX_ROUNDS) a.push(null);
                  a[cur] = Math.round(rem * 100) / 100;
                  amounts[r.name] = a;
                }
              }
              onRounds({ ...rounds, amounts });
            }}
          >
            Fill round {cur + 1} with remaining
          </Button>
        </div>
        <p className="mt-2 text-xs text-slate-500">
          Enter the amount traded in each round (buys positive, sells negative), as in the workbook's Round columns. “Remaining” is the workbook's own calculation. Each round has three sign-off initials. Six rounds at most, as in the workbook template.
        </p>
      </Card>

      {(props.checks.length > 0 || Object.values(props.exposure).some((v) => n(v) >= 0.1)) && (
        <div className="mb-6">
          <Notice tone="danger">
            <strong>Workbook checks:</strong>
            <ul className="ml-4 list-disc">{props.checks.map((c, i) => <li key={i}>{c}</li>)}</ul>
          </Notice>
        </div>
      )}

      <Card
        title="Implementation advice"
        className="mb-6"
        actions={<label className="flex items-center gap-2 text-sm"><input type="checkbox" checked={showAll} onChange={(e) => setShowAll(e.target.checked)} /> Show unchanged holdings</label>}
      >
        <TableWrap>
          <thead>
            <tr>
              <Th>Holding</Th>
              <Th>Broker</Th>
              <Th align="right">Spread buy/sell</Th>
              <Th align="right">Change</Th>
              {Array.from({ length: rounds.count }, (_, i) => <Th key={i} align="right" className={cx(i === cur && 'bg-fuchsia-50')}>Round {i + 1}</Th>)}
              <Th>Checks (R{cur + 1})</Th>
              <Th align="right">Remaining</Th>
            </tr>
          </thead>
          <tbody>
            {(['cash', 'fi', 'ae', 'ge', 'prop', 'ms'] as SectionKey[]).map((sec) => {
              const secRows = visible.filter((r) => r.section === sec);
              if (secRows.length === 0) return null;
              return [
                <tr key={`${sec}-h`} className="bg-slate-100"><Td colSpan={6 + rounds.count} className="text-xs font-semibold uppercase tracking-wide text-slate-600">{SECTION_LABEL[sec]}</Td></tr>,
                ...secRows.map((r) => {
                  const ch = n(r.change);
                  const [b, bk] = broker(r);
                  const rem = n(r.remaining);
                  const rationale = txt(r.rationale);
                  return [
                    <tr key={r.row} className="align-top">
                      <Td className="max-w-xs">
                        <div className="font-medium">{r.name}{props.addedNames.includes(r.name) && <span className="ml-2 rounded bg-sky-50 px-1 text-[11px] font-normal text-sky-800">added</span>}</div>
                        {cname(r) !== r.name && <div className="text-[11px] text-slate-500">{custodian}: {cname(r)}</div>}
                        {rationale && <button type="button" className="text-[11px] text-sky-700 underline" onClick={() => setOpenRationale(openRationale === r.name ? null : r.name)}>{openRationale === r.name ? 'Hide rationale' : 'Rationale'}</button>}
                      </Td>
                      <Td className="whitespace-nowrap text-xs">{b}{bk && <div className="text-slate-500">{bk}</div>}</Td>
                      <Td align="right" className="whitespace-nowrap text-xs">{rate(r.buySpread) || '—'} / {rate(r.sellSpread) || '—'}</Td>
                      <Td align="right" className={cx('whitespace-nowrap', ch > 0 && 'text-blue-700', ch < 0 && 'text-red-700')}>{ch === 0 ? '—' : fmt$(ch)}</Td>
                      {Array.from({ length: rounds.count }, (_, i) => (
                        <Td key={i} align="right" className={cx(i === cur && 'bg-fuchsia-50/60')}>
                          <AmountInput value={rounds.amounts[r.name]?.[i] ?? null} disabled={sec === 'cash' && false} onCommit={(v) => setAmount(r.name, i, v)} />
                        </Td>
                      ))}
                      <Td className="whitespace-nowrap">
                        {[0, 1, 2].map((k) => (
                          <input key={k} aria-label={`Check ${k + 1} ${r.name}`} maxLength={4} className="mr-1 w-10 rounded border border-slate-300 px-1 text-xs uppercase" value={rounds.checks[r.name]?.[cur]?.[k] ?? ''} onChange={(e) => setCheckInitials(r.name, k, e.target.value.toUpperCase())} />
                        ))}
                      </Td>
                      <Td align="right" className={cx('font-medium', Math.abs(rem) >= 1 ? 'text-amber-800' : 'text-emerald-700')}>{Math.abs(rem) < 1 ? '✓ done' : fmt$(r.remaining)}</Td>
                    </tr>,
                    openRationale === r.name ? <tr key={`${r.row}-r`}><Td colSpan={6 + rounds.count} className="bg-slate-50 text-xs text-slate-700">{rationale}</Td></tr> : null,
                  ];
                }),
                <tr key={`${sec}-t`} className="bg-slate-50 text-xs font-semibold">
                  <Td>Total</Td><Td /><Td /><Td align="right">{fmt$(secRows.reduce((a, r) => a + n(r.change), 0))}</Td>
                  {Array.from({ length: rounds.count }, (_, i) => <Td key={i} align="right">{fmt$(roundTotal(i, sec))}</Td>)}
                  <Td /><Td align="right">{fmt$(secRows.reduce((a, r) => a + n(r.remaining), 0))}</Td>
                </tr>,
              ];
            })}
          </tbody>
        </TableWrap>
        <div className="mt-4 grid gap-3 sm:grid-cols-[180px_1fr_auto] sm:items-end">
          <Field label="Add a holding to">
            <select className={inputCls} value={addSection} onChange={(e) => setAddSection(e.target.value as SectionKey)}>
              {(Object.keys(SECTION_LABEL) as SectionKey[]).map((k) => <option key={k} value={k}>{SECTION_LABEL[k]}</option>)}
            </select>
          </Field>
          <Field label="Investment (FS name)" hint="e.g. the actual bond bought for a “2027 Bonds (BBB+…)” row — its rounds count against that row's Remaining.">
            <input className={inputCls} list="fs-names" value={addName} onChange={(e) => setAddName(e.target.value)} placeholder="Start typing an investment name" />
            <datalist id="fs-names">{namesForSection.map((d) => <option key={d.name} value={d.name} />)}</datalist>
          </Field>
          <Button disabled={!addName.trim() || props.busy} onClick={() => { props.onAddHolding(addSection, addName.trim()); setAddName(''); }}>Add holding</Button>
        </div>
      </Card>

      <div className="mb-6 grid gap-6 xl:grid-cols-2">
        <Card title={`Trade instructions — round ${cur + 1}${rounds.dates[cur] ? ` (${rounds.dates[cur]})` : ''}`} actions={instructions.length > 0 ? <Button onClick={exportInstructions}>Export CSV</Button> : undefined}>
          {instructions.length === 0 ? (
            <p className="text-sm text-slate-500">No trades entered for this round.</p>
          ) : (
            <TableWrap>
              <thead><tr><Th>Action</Th><Th>{custodian} name</Th><Th align="right">Amount</Th><Th>Broker</Th></tr></thead>
              <tbody>
                {instructions.map(({ r, amt }) => (
                  <tr key={r.row}>
                    <Td className={cx('font-semibold', amt < 0 ? 'text-red-700' : 'text-blue-700')}>{amt < 0 ? 'Sell' : 'Buy'}</Td>
                    <Td className="max-w-sm">{cname(r)}{cname(r) !== r.name && <div className="text-[11px] text-slate-500">{r.name}</div>}</Td>
                    <Td align="right">{fmt$(Math.abs(amt))}</Td>
                    <Td className="text-xs">{broker(r).filter(Boolean).join(' · ') || '—'}</Td>
                  </tr>
                ))}
                <tr className="bg-slate-50 text-sm font-semibold">
                  <Td>Net</Td><Td /><Td align="right">{fmt$(instructions.reduce((a, x) => a + x.amt, 0))}</Td><Td />
                </tr>
              </tbody>
            </TableWrap>
          )}
        </Card>

        <Card title={`Computer check — round ${cur + 1}`}>
          <p className="mb-2 text-xs text-slate-600">Load the {custodian} transaction report (.xlsx or .csv). Amounts in this round, plus earlier rounds dated within 5 days of the report, are compared with the transactions for portfolio <strong>{props.portfolioId || '—'}</strong>; within $100 is a match.</p>
          <input type="file" accept=".xlsx,.xlsm,.csv" aria-label="Transaction report" disabled={!props.portfolioId} className="text-sm file:mr-2 file:rounded file:border file:border-slate-300 file:bg-white file:px-2 file:py-1" onChange={(e) => { const f = e.target.files?.[0]; e.target.value = ''; if (f) void runCheck(f); }} />
          {!props.portfolioId && <p className="mt-1 text-xs text-amber-800">Enter the portfolio number in client settings first.</p>}
          {checkError && <div className="mt-2"><Notice tone="danger">{checkError}</Notice></div>}
          {check && (
            <div className="mt-3 space-y-3">
              <p className="text-xs text-slate-600">{checkFile} · rounds included: {check.roundsIncluded.map((i) => i + 1).join(', ')}</p>
              <TableWrap>
                <thead><tr><Th>Holding</Th><Th align="right">Rounds</Th><Th align="right">Transactions</Th><Th>Result</Th></tr></thead>
                <tbody>
                  {check.lines.map((l) => (
                    <tr key={l.holding} className={l.status === 'match' ? 'bg-emerald-50' : 'bg-orange-50'}>
                      <Td className="max-w-xs">{l.holding}{l.custodianName !== l.holding && <div className="text-[11px] text-slate-500">{l.custodianName}</div>}</Td>
                      <Td align="right">{fmt$(l.roundsTotal)}</Td>
                      <Td align="right">{l.transactionAmount === null ? '—' : fmt$(l.transactionAmount)}</Td>
                      <Td className="whitespace-nowrap text-xs font-medium">{l.status === 'match' ? '✓ Match' : l.status === 'mismatch' ? '✕ Amount differs' : '? Not in transactions'}</Td>
                    </tr>
                  ))}
                </tbody>
              </TableWrap>
              {check.unmatchedTransactions.length > 0 && (
                <Notice tone="warning">
                  <strong>Transactions not in any round:</strong>
                  <ul className="ml-4 list-disc">{check.unmatchedTransactions.map((t, i) => <li key={i}>{t.assetName} — {t.type} {fmt$(Math.abs(t.amount))} ({t.eventDate})</li>)}</ul>
                </Notice>
              )}
            </div>
          )}
        </Card>
      </div>
    </>
  );
}

function AmountInput({ value, onCommit, disabled }: { value: number | null; onCommit: (v: number | null) => void; disabled?: boolean }) {
  const [t, setT] = useState(value === null ? '' : String(value));
  const [last, setLast] = useState(value);
  if (last !== value) {
    setLast(value);
    setT(value === null ? '' : String(value));
  }
  return (
    <input
      aria-label="Round amount"
      disabled={disabled}
      className={cx('w-24 rounded border px-1.5 py-0.5 text-right text-sm tabular-nums', value !== null && value < 0 ? 'border-red-200 text-red-700' : value ? 'border-blue-200 text-blue-700' : 'border-slate-300')}
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
