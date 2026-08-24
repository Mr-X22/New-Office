import { initShell } from './core/app-shell.js';

initShell();

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js').catch(() => {
      // Si falla (p. ej. sirviendo desde file://), la app sigue funcionando sin caché offline.
    });
  });
}
