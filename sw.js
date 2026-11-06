// Funciona sin señal: guarda la app en el teléfono (los datos viajan aparte).
const V = 'domus-v1';
const ARCHIVOS = ['./', './index.html', './icon-192.png', './icon-512.png'];
self.addEventListener('install', e => { e.waitUntil(caches.open(V).then(c => c.addAll(ARCHIVOS))); self.skipWaiting(); });
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== V).map(k => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener('fetch', e => {
  const u = new URL(e.request.url);
  if (e.request.method !== 'GET' || u.origin !== location.origin) return;
  e.respondWith(fetch(e.request).then(r => {
    const c = r.clone(); caches.open(V).then(x => x.put(e.request, c)); return r;
  }).catch(() => caches.match(e.request, { ignoreSearch: true })));
});
