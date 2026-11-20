import { Layout } from './components/Layout';
import { Assumptions } from './pages/Assumptions';
import { Comparison } from './pages/Comparison';
import { Dashboard } from './pages/Dashboard';
import { DataPage } from './pages/DataPage';
import { Designer } from './pages/Designer';
import { ClientHoldings } from './pages/ClientHoldings';
import { SettingsPage } from './pages/SettingsPage';
import { Recommendations } from './pages/Recommendations';
import { Scenarios } from './pages/Scenarios';
import { useRoute } from './router';

export function App() {
  const route = useRoute();
  let page;
  switch (route.page) {
    case 'designer':
      page = <Designer scenarioId={route.scenarioId} />;
      break;
    case 'compare':
      page = <Comparison profileId={route.profileId} scenarioId={route.scenarioId} />;
      break;
    case 'scenarios':
      page = <Scenarios />;
      break;
    case 'assumptions':
      page = <Assumptions />;
      break;
    case 'data':
      page = <DataPage />;
      break;
    case 'clients':
      page = <ClientHoldings />;
      break;
    case 'settings':
      page = <SettingsPage />;
      break;
    case 'recommendations':
      page = <Recommendations />;
      break;
    default:
      page = <Dashboard />;
  }
  return <Layout route={route}>{page}</Layout>;
}
