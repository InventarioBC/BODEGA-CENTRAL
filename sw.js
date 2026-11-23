// Service Worker para Bodega Central (Guzmán y Larraín Depto. Eléctrico)
// Versión 35: Carga instantánea como App Móvil (PWA) con soporte offline y auto-actualización.

const CACHE_NAME = 'bodega-central-v35';
const PRECACHE_ASSETS = [
  '/',
  '/Buscador_Bodega_Central.html',
  '/manifest.json',
  '/icon.png',
  '/apple-touch-icon.png'
];

// 1. INSTALACIÓN: Pre-cachear el cascarón de la app para que siempre cargue al abrir desde la pantalla de inicio
self.addEventListener('install', event => {
  console.log('[SW v35] Instalando y precacheando cascarón PWA...');
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(async cache => {
      for (const item of PRECACHE_ASSETS) {
        try {
          await cache.add(item);
        } catch (err) {
          console.warn('[SW v31] No se pudo precachear:', item, err);
        }
      }
      console.log('[SW v31] Cascarón PWA precacheado exitosamente.');
    })
  );
});

// 2. ACTIVACIÓN: Eliminar cachés obsoletas y tomar control inmediato
self.addEventListener('activate', event => {
  console.log('[SW v35] Activando y limpiando versiones anteriores...');
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(key => {
          if (key !== CACHE_NAME) {
            console.log('[SW v35] Eliminando caché vieja:', key);
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// 3. FETCH: Manejo de navegación para pantalla de inicio y recursos
self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // Ignorar peticiones API (siempre van en vivo al servidor)
  if (req.method !== 'GET' || url.pathname.startsWith('/api/')) {
    return;
  }

  // Para navegación principal (al abrir desde icono de inicio o recargar)
  if (req.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname === '/') {
    event.respondWith(
      // Intentar red primero con timeout de 2.5s para no bloquear si hay lentitud o túnel reconectando
      Promise.race([
        fetch(req, { cache: 'no-cache' }).then(networkResp => {
          if (networkResp && networkResp.status === 200) {
            const respClone = networkResp.clone();
            caches.open(CACHE_NAME).then(cache => {
              cache.put(req, respClone);
              cache.put('/Buscador_Bodega_Central.html', networkResp.clone());
              cache.put('/', networkResp.clone());
            });
            return networkResp;
          }
          // Si el servidor respondió con 502/503/504 (túnel caído), lanzar error para caer en caché
          throw new Error('Respuesta de red inválida: ' + (networkResp ? networkResp.status : 'desconocida'));
        }),
        new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout de red 2500ms')), 2500))
      ]).catch(err => {
        console.warn('[SW v30] Red falló o demoró, cargando desde caché PWA local:', err.message);
        return caches.match('/Buscador_Bodega_Central.html')
          .then(cached => cached || caches.match('/'))
          .then(cached => {
            if (cached) return cached;
            // Si no hay nada en caché, devolver mensaje amigable
            return new Response(
              '<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Bodega Central</title></head><body style="font-family:sans-serif;text-align:center;padding:40px 20px;background:#EAEFF4;color:#2D3748;"><h2>Bodega Central</h2><p>Conectando con el servidor...</p><button onclick="window.location.reload()" style="background:#4D6E87;color:white;border:none;padding:12px 24px;border-radius:12px;font-size:16px;cursor:pointer;">Reintentar</button></body></html>',
              { headers: { 'Content-Type': 'text/html; charset=utf-8' } }
            );
          });
      })
    );
    return;
  }

  // Para iconos, imágenes y archivos estáticos: Caché primero, con actualización en segundo plano
  event.respondWith(
    caches.match(req).then(cachedResp => {
      const fetchPromise = fetch(req).then(networkResp => {
        if (networkResp && networkResp.status === 200) {
          const respClone = networkResp.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(req, respClone));
        }
        return networkResp;
      }).catch(() => null);

      return cachedResp || fetchPromise;
    })
  );
});
