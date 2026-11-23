// Service Worker para Bodega Central (Guzmán y Larraín Depto. Eléctrico)
// Versión 37: PWA estable sin bloqueos de red ni bucles de recarga

const CACHE_NAME = 'bodega-central-v37';
const PRECACHE_ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon.png',
  './apple-touch-icon.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(PRECACHE_ASSETS).catch(err => {
        console.warn('[SW v37] Error precacheando:', err);
      });
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    })
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // Ignorar peticiones API
  if (req.method !== 'GET' || url.pathname.includes('/api/')) {
    return;
  }

  // Navegación (red primero, si no hay internet usar caché)
  if (req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname.endsWith('/')) {
    event.respondWith(
      fetch(req).catch(() => {
        return caches.match('./index.html').then(cached => cached || caches.match('./'));
      })
    );
    return;
  }

  // Recursos estáticos (imágenes, iconos, fuentes)
  event.respondWith(
    caches.match(req).then(cached => {
      return cached || fetch(req).then(resp => {
        if (resp && resp.status === 200) {
          const clone = resp.clone();
          caches.open(CACHE_NAME).then(c => c.put(req, clone));
        }
        return resp;
      }).catch(() => null);
    })
  );
});
