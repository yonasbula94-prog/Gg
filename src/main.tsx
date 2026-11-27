import { createRoot } from 'react-dom/client';
import App from './App.tsx';
import { AppErrorBoundary } from './components/AppErrorBoundary.tsx';
import '@fontsource/plus-jakarta-sans/500.css';
import '@fontsource/plus-jakarta-sans/600.css';
import '@fontsource/plus-jakarta-sans/700.css';
import '@fontsource/plus-jakarta-sans/800.css';
import '@fontsource/dm-sans/400.css';
import '@fontsource/dm-sans/500.css';
import '@fontsource/dm-sans/600.css';
import '@fontsource/dm-sans/700.css';
import '@fontsource/dm-sans/400-italic.css';
import '@fontsource/jetbrains-mono/400.css';
import '@fontsource/jetbrains-mono/500.css';
import '@fontsource/jetbrains-mono/600.css';
import './index.css';

// Unregister any legacy PWA service workers so deployed updates and /api/* calls are never served stale
try {
  if (typeof window !== 'undefined' && 'serviceWorker' in navigator) {
    navigator.serviceWorker
      .getRegistrations()
      .then((regs) => {
        regs.forEach((reg) => {
          reg.unregister().catch(() => {});
        });
      })
      .catch(() => {});
  }
  if (typeof window !== 'undefined' && 'caches' in window) {
    caches
      .keys()
      .then((keys) => {
        keys
          .filter((k) => k.includes('workbox') || k.includes('precache'))
          .forEach((k) => {
            caches.delete(k).catch(() => {});
          });
      })
      .catch(() => {});
  }
} catch {
  // Ignore when storage/SW APIs are restricted
}

createRoot(document.getElementById('root')!).render(
  <AppErrorBoundary>
    <App />
  </AppErrorBoundary>
);
