import { useEffect, useState } from 'react';

export type Route =
  | { page: 'dashboard' }
  | { page: 'designer'; scenarioId: string | null }
  | { page: 'compare'; profileId: string | null; scenarioId: string | null }
  | { page: 'scenarios' }
  | { page: 'assumptions' }
  | { page: 'data' }
  | { page: 'clients' }
  | { page: 'settings' };

export function parseHash(hash: string): Route {
  const [path = '', query = ''] = hash.replace(/^#\/?/, '').split('?');
  const parts = path.split('/').filter(Boolean).map(decodeURIComponent);
  const q = new URLSearchParams(query);
  switch (parts[0]) {
    case 'designer':
      return { page: 'designer', scenarioId: parts[1] ?? null };
    case 'compare':
      return { page: 'compare', profileId: parts[1] ?? null, scenarioId: q.get('scenario') };
    case 'scenarios':
      return { page: 'scenarios' };
    case 'assumptions':
      return { page: 'assumptions' };
    case 'data':
      return { page: 'data' };
    case 'clients':
      return { page: 'clients' };
    case 'settings':
      return { page: 'settings' };
    default:
      return { page: 'dashboard' };
  }
}

export function href(route: Route): string {
  switch (route.page) {
    case 'designer':
      return route.scenarioId ? `#/designer/${encodeURIComponent(route.scenarioId)}` : '#/designer';
    case 'compare': {
      let h = route.profileId ? `#/compare/${encodeURIComponent(route.profileId)}` : '#/compare';
      if (route.scenarioId) h += `?scenario=${encodeURIComponent(route.scenarioId)}`;
      return h;
    }
    default:
      return `#/${route.page}`;
  }
}

export function navigate(route: Route): void {
  window.location.hash = href(route);
}

export function useRoute(): Route {
  const [route, setRoute] = useState(() => parseHash(window.location.hash));
  useEffect(() => {
    const on = () => setRoute(parseHash(window.location.hash));
    window.addEventListener('hashchange', on);
    return () => window.removeEventListener('hashchange', on);
  }, []);
  return route;
}
