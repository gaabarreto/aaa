/* Service Worker de DataPrice Vzla.
   Cachea el "cascarón" de la app (HTML/JS/CSS/maestros JSON/íconos) para que
   funcione sin conexión. Las llamadas a Firebase/CDNs externos NO se tocan:
   van directo a la red, tal cual las maneja app.js.

   Estrategia "red primero": siempre intenta traer la versión más nueva del
   servidor (así, al subir cambios a Netlify, los asesores la ven al primer
   intento). Solo si no hay señal, o la red tarda más de TIEMPO_ESPERA_RED,
   usa la copia guardada. */
const CACHE_NAME = 'dataprice-vzla-v2';
const TIEMPO_ESPERA_RED = 4000;
const ARCHIVOS_CASCARON = [
  './',
  './index.html',
  './app.js',
  './style.css',
  './manifest.json',
  './data/productos.json',
  './data/clientes.json',
  './data/vendedores.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(ARCHIVOS_CASCARON))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((nombres) =>
      Promise.all(nombres.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);

  // Solo maneja archivos propios del mismo origen (no Firebase, no CDNs).
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    // cache: 'no-cache' evita que el navegador devuelva su propia copia vieja.
    const redFetch = fetch(event.request, { cache: 'no-cache' }).then((respuesta) => {
      if (respuesta && respuesta.ok) cache.put(event.request, respuesta.clone());
      return respuesta;
    });

    const cacheado = await cache.match(event.request);
    if (!cacheado) return redFetch; // nunca se guardó: no hay alternativa a la red

    // Red con límite de tiempo: si no responde a tiempo o falla, usa la copia guardada
    // (la descarga sigue por detrás y deja actualizada la copia para la próxima).
    const tiempoAgotado = new Promise((resolve) => setTimeout(() => resolve(null), TIEMPO_ESPERA_RED));
    const respuesta = await Promise.race([redFetch.catch(() => null), tiempoAgotado]);
    return respuesta || cacheado;
  })());
});
