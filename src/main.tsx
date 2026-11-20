import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { App } from './App';
import { AppProvider } from './state/AppContext';
import { ClientDataProvider } from './state/ClientDataContext';
import './index.css';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <AppProvider>
      <ClientDataProvider>
        <App />
      </ClientDataProvider>
    </AppProvider>
  </StrictMode>,
);
