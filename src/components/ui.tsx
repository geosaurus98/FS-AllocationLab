import type { ReactNode } from 'react';
import type { DataStatus, Severity, ValidationIssue, ValidationStatus } from '../lib/types';

export function cx(...c: (string | number | boolean | null | undefined)[]): string {
  return c.filter((x): x is string => typeof x === 'string' && x.length > 0).join(' ');
}

export function PageHeader({ title, subtitle, actions }: { title: string; subtitle?: ReactNode; actions?: ReactNode }) {
  return (
    <div className="mb-6 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
      <div>
        <h1 className="text-xl font-semibold text-slate-900">{title}</h1>
        {subtitle && <p className="mt-1 max-w-3xl text-sm text-slate-600">{subtitle}</p>}
      </div>
      {actions && <div className="flex flex-wrap gap-2">{actions}</div>}
    </div>
  );
}

export function Card({ title, children, actions, className }: { title?: ReactNode; children: ReactNode; actions?: ReactNode; className?: string }) {
  return (
    <section className={cx('rounded-lg border border-slate-200 bg-white shadow-sm', className)}>
      {(title || actions) && (
        <header className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 px-4 py-3">
          {title && <h2 className="text-sm font-semibold text-slate-800">{title}</h2>}
          {actions && <div className="flex flex-wrap gap-2">{actions}</div>}
        </header>
      )}
      <div className="p-4">{children}</div>
    </section>
  );
}

type BtnVariant = 'primary' | 'secondary' | 'danger' | 'ghost';
export function Button({
  variant = 'secondary',
  className,
  ...props
}: React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: BtnVariant }) {
  const styles: Record<BtnVariant, string> = {
    primary: 'bg-slate-900 text-white hover:bg-slate-700 disabled:bg-slate-300',
    secondary: 'border border-slate-300 bg-white text-slate-800 hover:bg-slate-50 disabled:text-slate-400',
    danger: 'border border-red-300 bg-white text-red-700 hover:bg-red-50 disabled:text-slate-400',
    ghost: 'text-slate-700 hover:bg-slate-100 disabled:text-slate-400',
  };
  return (
    <button
      type="button"
      {...props}
      className={cx('inline-flex items-center gap-1 whitespace-nowrap rounded-md px-3 py-1.5 text-sm font-medium transition-colors disabled:cursor-not-allowed', styles[variant], className)}
    />
  );
}

const STATUS_STYLE: Record<DataStatus, { label: string; cls: string; title: string }> = {
  approved: { label: 'Approved', cls: 'bg-emerald-50 text-emerald-800 ring-emerald-600/30', title: 'Marked approved in this workspace. Locked against edits and import overwrites.' },
  illustrative: { label: 'Illustrative', cls: 'bg-sky-50 text-sky-800 ring-sky-600/30', title: 'Illustrative example — not an approved allocation.' },
  synthetic: { label: 'Synthetic', cls: 'bg-violet-50 text-violet-800 ring-violet-600/30', title: 'Synthetic demo data — not real or approved.' },
  draft: { label: 'Draft', cls: 'bg-amber-50 text-amber-800 ring-amber-600/30', title: 'Draft — not approved.' },
  missing: { label: 'Missing', cls: 'bg-slate-100 text-slate-600 ring-slate-400/40', title: 'No data supplied.' },
};

export function StatusBadge({ status }: { status: DataStatus }) {
  const s = STATUS_STYLE[status];
  return (
    <span title={s.title} className={cx('inline-flex items-center rounded px-1.5 py-0.5 text-xs font-medium ring-1 ring-inset', s.cls)}>
      {s.label}
    </span>
  );
}

export function HypotheticalBadge() {
  return (
    <span title="Hypothetical proposal — not an approved allocation." className="inline-flex items-center rounded border border-dashed border-orange-400 bg-orange-50 px-1.5 py-0.5 text-xs font-medium text-orange-800">
      Hypothetical
    </span>
  );
}

const VSTATUS: Record<ValidationStatus, { label: string; icon: string; cls: string }> = {
  valid: { label: 'Valid', icon: '✓', cls: 'text-emerald-700 bg-emerald-50 ring-emerald-600/30' },
  warnings: { label: 'Warnings', icon: '!', cls: 'text-amber-800 bg-amber-50 ring-amber-600/30' },
  invalid: { label: 'Invalid', icon: '✕', cls: 'text-red-700 bg-red-50 ring-red-600/30' },
  incomplete: { label: 'Incomplete', icon: '?', cls: 'text-orange-800 bg-orange-50 ring-orange-600/30' },
  missing: { label: 'No data', icon: '–', cls: 'text-slate-600 bg-slate-100 ring-slate-400/40' },
};

export function ValidationBadge({ status, count }: { status: ValidationStatus; count?: number }) {
  const v = VSTATUS[status];
  return (
    <span className={cx('inline-flex items-center gap-1 rounded px-1.5 py-0.5 text-xs font-medium ring-1 ring-inset', v.cls)}>
      <span aria-hidden className="font-bold">{v.icon}</span>
      {v.label}
      {count !== undefined && count > 0 && <span className="opacity-70">({count})</span>}
    </span>
  );
}

const SEV: Record<Severity, { icon: string; cls: string; label: string }> = {
  error: { icon: '✕', cls: 'border-red-200 bg-red-50 text-red-800', label: 'Error' },
  warning: { icon: '!', cls: 'border-amber-200 bg-amber-50 text-amber-900', label: 'Warning' },
  info: { icon: 'i', cls: 'border-slate-200 bg-slate-50 text-slate-700', label: 'Info' },
};

export function IssueList({ issues, emptyText = 'No validation issues.' }: { issues: ValidationIssue[]; emptyText?: string }) {
  if (issues.length === 0) {
    return <p className="flex items-center gap-2 text-sm text-emerald-700"><span aria-hidden className="font-bold">✓</span>{emptyText}</p>;
  }
  return (
    <ul className="space-y-1.5">
      {issues.map((i, idx) => {
        const s = SEV[i.severity];
        return (
          <li key={idx} className={cx('flex gap-2 rounded border px-2.5 py-1.5 text-sm', s.cls)}>
            <span aria-hidden className="mt-px w-3 shrink-0 text-center font-bold">{s.icon}</span>
            <span><span className="sr-only">{s.label}: </span>{i.message}</span>
          </li>
        );
      })}
    </ul>
  );
}

export function EmptyState({ title, children, action }: { title: string; children?: ReactNode; action?: ReactNode }) {
  return (
    <div className="flex flex-col items-center justify-center rounded-lg border border-dashed border-slate-300 bg-white px-6 py-10 text-center">
      <p className="text-sm font-semibold text-slate-800">{title}</p>
      {children && <div className="mt-1 max-w-md text-sm text-slate-600">{children}</div>}
      {action && <div className="mt-4">{action}</div>}
    </div>
  );
}

export function Notice({ tone = 'info', children, onDismiss }: { tone?: 'info' | 'warning' | 'danger' | 'success'; children: ReactNode; onDismiss?: () => void }) {
  const cls = {
    info: 'border-sky-200 bg-sky-50 text-sky-900',
    warning: 'border-amber-200 bg-amber-50 text-amber-900',
    danger: 'border-red-200 bg-red-50 text-red-900',
    success: 'border-emerald-200 bg-emerald-50 text-emerald-900',
  }[tone];
  return (
    <div className={cx('flex items-start justify-between gap-3 rounded-md border px-3 py-2 text-sm', cls)} role={tone === 'danger' ? 'alert' : 'status'}>
      <div>{children}</div>
      {onDismiss && <button type="button" onClick={onDismiss} className="text-xs underline">Dismiss</button>}
    </div>
  );
}

export function Field({ label, children, hint }: { label: string; children: ReactNode; hint?: ReactNode }) {
  return (
    <label className="block text-sm">
      <span className="mb-1 block font-medium text-slate-700">{label}</span>
      {children}
      {hint && <span className="mt-1 block text-xs text-slate-500">{hint}</span>}
    </label>
  );
}

export const inputCls =
  'w-full rounded-md border border-slate-300 bg-white px-2 py-1.5 text-sm text-slate-900 shadow-sm focus:border-slate-500 focus:outline-none focus:ring-1 focus:ring-slate-500';

export function Th({ children, className, align = 'left' }: { children?: ReactNode; className?: string; align?: 'left' | 'right' | 'center' }) {
  return (
    <th scope="col" className={cx('whitespace-nowrap border-b border-slate-200 bg-slate-50 px-2.5 py-2 text-xs font-semibold uppercase tracking-wide text-slate-600', align === 'right' ? 'text-right' : align === 'center' ? 'text-center' : 'text-left', className)}>
      {children}
    </th>
  );
}

export function Td({ children, className, align = 'left', colSpan }: { children?: ReactNode; className?: string; align?: 'left' | 'right' | 'center'; colSpan?: number }) {
  return (
    <td colSpan={colSpan} className={cx('border-b border-slate-100 px-2.5 py-2 text-sm text-slate-800', align === 'right' && 'text-right tabular-nums', align === 'center' && 'text-center', className)}>
      {children}
    </td>
  );
}

export function TableWrap({ children }: { children: ReactNode }) {
  return <div className="-mx-4 overflow-x-auto px-4"><table className="min-w-full border-collapse">{children}</table></div>;
}
