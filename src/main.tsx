import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import { AppProvider } from './state/AppContext';
import { ClientDataProvider } from './state/ClientDataContext';
import { EngineSettingsProvider } from './state/EngineSettingsContext';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <AppProvider>
      <EngineSettingsProvider>
        <ClientDataProvider>
          <App />
        </ClientDataProvider>
      </EngineSettingsProvider>
    </AppProvider>
  </StrictMode>,
);
