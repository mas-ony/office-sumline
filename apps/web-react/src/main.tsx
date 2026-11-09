import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import '@sumline/react/style.css';
import '@sumline/web-shared/app.css';
import { App } from './App';
import { registerServiceWorker } from '@sumline/web-shared/shell';

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
);

if (import.meta.env.PROD) registerServiceWorker();
