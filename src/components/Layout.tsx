import { useState, type ReactNode } from 'react';
import type { Route } from '../router';
import { href } from '../router';
import { useApp } from '../state/AppContext';
import { cx, Notice } from './ui';

const NAV: { page: Route['page']; label: string; icon: string }[] = [
  { page: 'dashboard', label: 'Dashboard', icon: '▦' },
  { page: 'designer', label: 'Allocation Designer', icon: '✎' },
  { page: 'compare', label: 'Profile Comparison', icon: '⇄' },
  { page: 'scenarios', label: 'Scenarios', icon: '◫' },
  { page: 'assumptions', label: 'Assumptions', icon: '§' },
  { page: 'data', label: 'Import / Export', icon: '⇅' },
];

function routeFor(page: Route['page']): Route {
  switch (page) {
    case 'designer':
      return { page, scenarioId: null };
    case 'compare':
      return { page, profileId: null, scenarioId: null };
    default:
      return { page } as Route;
  }
}

export function Layout({ route, children }: { route: Route; children: ReactNode }) {
  const [open, setOpen] = useState(false);
  const { notice, dismissNotice, persistError } = useApp();
  return (
    <div className="min-h-screen bg-slate-50 text-slate-900 lg:flex">
      <aside className={cx('fixed inset-y-0 left-0 z-30 w-60 shrink-0 border-r border-slate-800 bg-slate-900 text-slate-200 transition-transform lg:static lg:translate-x-0', open ? 'translate-x-0' : '-translate-x-full')}>
        <div className="border-b border-slate-800 px-5 py-4">
          <div className="text-base font-semibold text-white">AllocationLab</div>
          <div className="text-xs text-slate-400">Financial Strategies · internal</div>
        </div>
        <nav className="p-2" aria-label="Main">
          {NAV.map((n) => (
            <a
              key={n.page}
              href={href(routeFor(n.page))}
              onClick={() => setOpen(false)}
              aria-current={route.page === n.page ? 'page' : undefined}
              className={cx('mb-0.5 flex items-center gap-3 rounded-md px-3 py-2 text-sm', route.page === n.page ? 'bg-slate-700 text-white' : 'text-slate-300 hover:bg-slate-800 hover:text-white')}
            >
              <span aria-hidden className="w-4 text-center text-slate-400">{n.icon}</span>
              {n.label}
            </a>
          ))}
        </nav>
        <div className="absolute bottom-0 w-full border-t border-slate-800 p-4 text-[11px] leading-relaxed text-slate-400">
          Prototype. Synthetic &amp; illustrative data only. Not investment advice and not a recommendation.
        </div>
      </aside>
      {open && <div className="fixed inset-0 z-20 bg-black/30 lg:hidden" onClick={() => setOpen(false)} />}
      <div className="min-w-0 flex-1">
        <div className="sticky top-0 z-10 flex items-center gap-3 border-b border-amber-200 bg-amber-50 px-4 py-1.5 text-xs text-amber-900">
          <button type="button" className="rounded border border-amber-300 px-2 py-0.5 lg:hidden" onClick={() => setOpen(true)} aria-label="Open navigation">☰</button>
          <span><strong>Prototype workspace.</strong> Values are synthetic or illustrative unless marked <em>Approved</em>. Scenarios are hypothetical and never change baselines.</span>
        </div>
        <main className="mx-auto max-w-7xl px-4 py-6 sm:px-6">
          {notice && <div className="mb-4"><Notice tone="warning" onDismiss={dismissNotice}>{notice}</Notice></div>}
          {persistError && <div className="mb-4"><Notice tone="danger">Changes could not be saved to browser storage. They will be lost when the page is closed — export a CSV to keep them.</Notice></div>}
          {children}
        </main>
      </div>
    </div>
  );
}
