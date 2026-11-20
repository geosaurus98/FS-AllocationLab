import { useEffect, useState } from 'react';
import { parsePercent } from '../lib/format';
import { cx } from './ui';

interface Props {
  value: number | null;
  onChange: (v: number | null) => void;
  ariaLabel: string;
  placeholder?: string;
  invalid?: boolean;
  disabled?: boolean;
  className?: string;
  suffix?: string;
}

/**
 * Text-based numeric input. Keeps the user's raw text while typing; commits only
 * parseable values. Blank commits `null` (missing) — never 0. Values are not clamped or rounded.
 */
export function PercentInput({ value, onChange, ariaLabel, placeholder = '—', invalid, disabled, className, suffix = '%' }: Props) {
  const [text, setText] = useState(value === null ? '' : String(value));
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const parsed = parsePercent(text);
    if (!(parsed.ok && parsed.value === value)) {
      setText(value === null ? '' : String(value));
      setError(null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  return (
    <div className={cx('relative', className)}>
      <input
        type="text"
        inputMode="decimal"
        aria-label={ariaLabel}
        aria-invalid={Boolean(error || invalid)}
        title={error ?? undefined}
        disabled={disabled}
        value={text}
        placeholder={placeholder}
        onChange={(e) => {
          const t = e.target.value;
          setText(t);
          const p = parsePercent(t);
          if (p.ok) {
            setError(null);
            onChange(p.value);
          } else setError(p.error);
        }}
        className={cx(
          'w-full rounded-md border bg-white py-1 pl-2 pr-6 text-right text-sm tabular-nums shadow-sm focus:outline-none focus:ring-1 disabled:bg-slate-50 disabled:text-slate-500',
          error ? 'border-red-500 focus:ring-red-500' : invalid ? 'border-red-400 bg-red-50 focus:ring-red-500' : 'border-slate-300 focus:border-slate-500 focus:ring-slate-500',
        )}
      />
      <span className="pointer-events-none absolute right-2 top-1/2 -translate-y-1/2 text-xs text-slate-400">{suffix}</span>
      {error && <span className="absolute left-0 top-full z-10 mt-0.5 whitespace-nowrap rounded bg-red-700 px-1.5 py-0.5 text-[11px] text-white">{error}</span>}
    </div>
  );
}
