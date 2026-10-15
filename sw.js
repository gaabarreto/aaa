/* Service Worker de DataPrice Vzla.
   Cachea el "cascarón" de la app (HTML/JS/CSS/maestros JSON/íconos) para que
   abra instantáneo y funcione sin conexión, igual que ya lo hacía el sitio
   en el navegador. Las llamadas a Firebase/CDNs externos NO se tocan: van
   directo a la red, tal cual las maneja app.js. */
const CACHE_NAME = 'dataprice-vzla-v1';
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

  event.respondWith(
    caches.match(event.request).then((cacheado) => {
      const redFetch = fetch(event.request)
        .then((respuesta) => {
          if (respuesta && respuesta.ok) {
            const clon = respuesta.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clon));
          }
          return respuesta;
        })
        .catch(() => cacheado); // sin conexión: usa lo cacheado

      return cacheado || redFetch;
    })
  );
});
