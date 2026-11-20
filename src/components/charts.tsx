import { useState } from 'react';
import { Bar, BarChart, CartesianGrid, Cell, ReferenceLine, ResponsiveContainer, Tooltip, XAxis, YAxis } from 'recharts';
import { formatPct, formatPp } from '../lib/format';
import { ASSET_CLASSES, ASSET_COLORS } from '../lib/reference';
import type { AssetClassId, Weights } from '../lib/types';
import { cx } from './ui';

const INK = { primary: '#0b0b0b', secondary: '#52514e', muted: '#898781', grid: '#e1e0d9', axis: '#c3c2b7' };

export interface StackRow {
  key: string;
  label: string;
  weights: Weights;
  /** Short tag shown in tooltip, e.g. "Synthetic", "Hypothetical". */
  tag?: string;
}

interface TooltipPayload {
  payload?: { label: string; tag?: string; missing: string[]; total: number } & Partial<Record<AssetClassId, number>>;
}

function StackTooltip({ active, payload, focus }: { active?: boolean; payload?: TooltipPayload[]; focus: AssetClassId | null }) {
  const p = payload?.[0]?.payload;
  if (!active || !p) return null;
  return (
    <div className="rounded-md border border-slate-200 bg-white px-3 py-2 text-xs shadow-lg">
      <div className="mb-1 font-semibold text-slate-900">
        {p.label} {p.tag && <span className="font-normal text-slate-500">· {p.tag}</span>}
      </div>
      <table>
        <tbody>
          {ASSET_CLASSES.map((ac) => (
            <tr key={ac.id} className={cx(focus === ac.id && 'font-semibold')}>
              <td className="pr-2"><span className="mr-1.5 inline-block h-2.5 w-2.5 rounded-sm align-middle" style={{ background: ASSET_COLORS[ac.id] }} />{ac.label}</td>
              <td className="text-right tabular-nums">{p.missing.includes(ac.id) ? <span className="text-orange-700">missing</span> : formatPct(p[ac.id] ?? 0)}</td>
            </tr>
          ))}
          <tr className="border-t border-slate-200"><td className="pr-2 pt-1">Total</td><td className="pt-1 text-right tabular-nums">{formatPct(p.total)}</td></tr>
        </tbody>
      </table>
    </div>
  );
}

export function AssetLegend({ focus, onFocus }: { focus: AssetClassId | null; onFocus: (id: AssetClassId | null) => void }) {
  return (
    <ul className="flex flex-wrap gap-x-4 gap-y-1 text-xs text-slate-700" aria-label="Asset class legend">
      {ASSET_CLASSES.map((ac) => (
        <li key={ac.id}>
          <button
            type="button"
            className={cx('flex items-center gap-1.5 rounded px-1 py-0.5 hover:bg-slate-100', focus && focus !== ac.id && 'opacity-40')}
            onMouseEnter={() => onFocus(ac.id)}
            onMouseLeave={() => onFocus(null)}
            onFocus={() => onFocus(ac.id)}
            onBlur={() => onFocus(null)}
          >
            <span className="inline-block h-2.5 w-2.5 rounded-sm" style={{ background: ASSET_COLORS[ac.id] }} />
            {ac.label}
          </button>
        </li>
      ))}
    </ul>
  );
}

/**
 * Horizontal stacked allocation bars (one bar per row). Missing weights are drawn as nothing
 * and called out in the tooltip and a note beneath — they are never filled in.
 */
export function StackedAllocationChart({ rows, height }: { rows: StackRow[]; height?: number }) {
  const [focus, setFocus] = useState<AssetClassId | null>(null);
  const data = rows.map((r) => {
    const missing = ASSET_CLASSES.filter((a) => r.weights[a.id] === null).map((a) => a.id);
    const entry: Record<string, unknown> = { key: r.key, label: r.label, tag: r.tag, missing, total: 0 };
    let total = 0;
    for (const a of ASSET_CLASSES) {
      const v = r.weights[a.id];
      entry[a.id] = v ?? 0;
      total += v ?? 0;
    }
    entry.total = total;
    return entry;
  });
  const maxTotal = Math.max(100, ...data.map((d) => d.total as number));
  const h = height ?? Math.max(120, rows.length * 44 + 40);
  const anyMissing = data.some((d) => (d.missing as string[]).length > 0);

  return (
    <div>
      <div className="mb-2"><AssetLegend focus={focus} onFocus={setFocus} /></div>
      <div style={{ height: h }} className="w-full">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={data} layout="vertical" margin={{ top: 4, right: 16, bottom: 4, left: 4 }} barCategoryGap={10}>
            <CartesianGrid horizontal={false} stroke={INK.grid} />
            <XAxis type="number" domain={[0, Math.ceil(maxTotal / 10) * 10]} tickFormatter={(v: number) => `${v}%`} tick={{ fontSize: 11, fill: INK.muted }} axisLine={{ stroke: INK.axis }} tickLine={false} />
            <YAxis type="category" dataKey="label" width={140} tick={{ fontSize: 12, fill: INK.secondary }} axisLine={{ stroke: INK.axis }} tickLine={false} />
            <ReferenceLine x={100} stroke={INK.secondary} strokeDasharray="3 3" />
            <Tooltip content={<StackTooltip focus={focus} />} cursor={{ fill: 'rgba(11,11,11,0.04)' }} />
            {ASSET_CLASSES.map((ac, i) => (
              <Bar
                key={ac.id}
                dataKey={ac.id}
                stackId="a"
                fill={ASSET_COLORS[ac.id]}
                stroke="#ffffff"
                strokeWidth={1}
                fillOpacity={focus && focus !== ac.id ? 0.25 : 1}
                radius={i === ASSET_CLASSES.length - 1 ? [0, 4, 4, 0] : 0}
                isAnimationActive={false}
                onMouseEnter={() => setFocus(ac.id)}
                onMouseLeave={() => setFocus(null)}
              />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </div>
      {anyMissing && <p className="mt-1 text-xs text-orange-700">Some bars are incomplete: missing weights are not drawn and are not assumed to be zero.</p>}
    </div>
  );
}

const DIV_POS = '#2a78d6';
const DIV_NEG = '#e34948';

/** Percentage-point differences per asset class (diverging bars around zero). */
export function DiffChart({ diff, label }: { diff: Record<AssetClassId, number | null>; label: string }) {
  const data = ASSET_CLASSES.map((ac) => ({ name: ac.shortLabel, full: ac.label, value: diff[ac.id] ?? 0, missing: diff[ac.id] === null }));
  const maxAbs = Math.max(1, ...data.map((d) => Math.abs(d.value)));
  const lim = Math.ceil(maxAbs);
  return (
    <div>
      <p className="mb-2 text-xs text-slate-600">{label} <span className="ml-2 inline-flex items-center gap-1"><span className="inline-block h-2.5 w-2.5 rounded-sm" style={{ background: DIV_POS }} />increase</span> <span className="ml-2 inline-flex items-center gap-1"><span className="inline-block h-2.5 w-2.5 rounded-sm" style={{ background: DIV_NEG }} />decrease</span></p>
      <div style={{ height: 220 }}>
        <ResponsiveContainer width="100%" height="100%">
          <BarChart data={data} margin={{ top: 8, right: 8, bottom: 4, left: 0 }}>
            <CartesianGrid vertical={false} stroke={INK.grid} />
            <XAxis dataKey="name" tick={{ fontSize: 11, fill: INK.secondary }} axisLine={{ stroke: INK.axis }} tickLine={false} interval={0} />
            <YAxis domain={[-lim, lim]} tickFormatter={(v: number) => `${v}`} tick={{ fontSize: 11, fill: INK.muted }} axisLine={false} tickLine={false} width={36} unit=" pp" />
            <ReferenceLine y={0} stroke={INK.secondary} />
            <Tooltip
              cursor={{ fill: 'rgba(11,11,11,0.04)' }}
              content={({ active, payload }) => {
                const p = payload?.[0]?.payload as (typeof data)[number] | undefined;
                if (!active || !p) return null;
                return (
                  <div className="rounded-md border border-slate-200 bg-white px-3 py-2 text-xs shadow-lg">
                    <div className="font-semibold">{p.full}</div>
                    <div className="tabular-nums">{p.missing ? 'n/a — missing data' : formatPp(p.value)}</div>
                  </div>
                );
              }}
            />
            <Bar dataKey="value" isAnimationActive={false} maxBarSize={36}>
              {data.map((d) => (
                <Cell key={d.name} fill={d.missing ? '#e1e0d9' : d.value >= 0 ? DIV_POS : DIV_NEG} />
              ))}
            </Bar>
          </BarChart>
        </ResponsiveContainer>
      </div>
    </div>
  );
}

/** Compact growth / defensive / unclassified bar. */
export function SplitBar({ growth, defensive, unclassified, compact }: { growth: number; defensive: number; unclassified: number; compact?: boolean }) {
  const total = Math.max(growth + defensive + unclassified, 0.0001);
  const seg = (v: number) => `${(v / Math.max(total, 100)) * 100}%`;
  return (
    <div className={cx('w-full', compact ? 'min-w-[120px]' : '')}>
      <div className="flex h-2.5 w-full overflow-hidden rounded-full bg-slate-100" aria-hidden>
        <div style={{ width: seg(growth), background: '#0d366b' }} />
        <div style={{ width: seg(defensive), background: '#86b6ef' }} className="border-l border-white" />
        {unclassified > 0 && <div style={{ width: seg(unclassified), backgroundImage: 'repeating-linear-gradient(45deg,#fab219 0 3px,#fff 3px 6px)' }} className="border-l border-white" />}
      </div>
      {!compact && (
        <div className="mt-1.5 flex flex-wrap gap-x-4 text-xs text-slate-700">
          <span><span className="mr-1 inline-block h-2 w-2 rounded-sm bg-[#0d366b]" />Growth {formatPct(growth)}</span>
          <span><span className="mr-1 inline-block h-2 w-2 rounded-sm bg-[#86b6ef]" />Defensive {formatPct(defensive)}</span>
          {unclassified > 0 && <span className="text-amber-800"><span className="mr-1 inline-block h-2 w-2 rounded-sm bg-[#fab219]" />Unclassified {formatPct(unclassified)}</span>}
        </div>
      )}
    </div>
  );
}
