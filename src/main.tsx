import { createRoot } from 'react-dom/client';
import { registerSW } from 'virtual:pwa-register';
import App from './App.tsx';
import './index.css';

// Register offline Service Worker only in top-level http/https windows (skip in iframes and file://)
try {
  if (
    typeof window !== 'undefined' &&
    window.location.protocol !== 'file:' &&
    window.self === window.top
  ) {
    registerSW({ immediate: true });
  }
} catch {
  // Ignore SW registration errors when cookies/storage are restricted
}

createRoot(document.getElementById('root')!).render(<App />);
