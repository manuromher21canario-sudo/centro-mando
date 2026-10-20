// Ayudante mínimo para que el móvil deje instalar la web como app.
// No guarda nada: todo se pide siempre a internet (así nunca se ven datos viejos).
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
