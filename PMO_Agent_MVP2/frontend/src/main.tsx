import { createRoot } from 'react-dom/client';
import './index.css';
import App from './App.tsx';

const rootElement = document.getElementById('root');

if (rootElement) {
  window.addEventListener('error', (event) => {
    rootElement.innerHTML = `
      <div style="padding: 2rem; background: #fee2e2; border: 1px solid #ef4444; color: #991b1b; border-radius: 8px; font-family: sans-serif; margin: 2rem;">
        <h1 style="font-size: 1.25rem; margin-bottom: 0.5rem;">Runtime Error Captured</h1>
        <p style="font-family: monospace; font-size: 0.875rem; background: rgba(255,255,255,0.5); padding: 0.5rem; border-radius: 4px;">${event.message}</p>
        <p style="font-size: 0.75rem; color: #b91c1c; margin-top: 0.5rem;">File: ${event.filename} | Line: ${event.lineno}:${event.colno}</p>
        <button onclick="window.location.reload()" style="margin-top: 1rem; padding: 0.5rem 1rem; background: #ef4444; color: white; border: none; border-radius: 4px; cursor: pointer;">Reload Page</button>
      </div>
    `;
  });

  createRoot(rootElement).render(<App />);
}
