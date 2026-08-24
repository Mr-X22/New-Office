const CACHE_NAME = 'carpeta-shell-v1';

// Solo el esqueleto (shell) se precachea. Los módulos y sus librerías se
// cachean "on demand" la primera vez que se usan (runtime caching abajo),
// para no gastar espacio en disco con cosas que tal vez nunca se abran.
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.json',
  './css/style.css',
  './js/app.js',
  './js/core/app-shell.js',
  './js/core/storage.js',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;

  event.respondWith(
    caches.match(request).then((cached) => {
      const network = fetch(request)
        .then((response) => {
          if (response.ok) {
            const clone = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, clone));
          }
          return response;
        })
        .catch(() => cached);
      // Estrategia stale-while-revalidate: responde rápido con caché si existe,
      // y de paso actualiza en segundo plano.
      return cached || network;
    })
  );
});
