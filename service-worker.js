// Simple Service Worker - Cache static files only
const CACHE_NAME = 'his-pro-cache-v1';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(cacheNames => {
      return Promise.all(
        cacheNames.map(cacheName => {
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const { request } = event;
  const url = new URL(request.url);

  // Skip non-GET requests
  if (request.method !== 'GET') {
    return;
  }

  // Skip external requests
  if (url.origin !== location.origin) {
    return;
  }

  // Only cache static files (css, js, fonts, images)
  const isStaticAsset = /\.(js|css|png|jpg|jpeg|gif|svg|woff|woff2|ttf|eot)$/.test(url.pathname);

  if (!isStaticAsset) {
    return; // Don't cache dynamic content
  }

  event.respondWith(
    caches.open(CACHE_NAME).then(cache => {
      return cache.match(request).then(response => {
        if (response) {
          return response;
        }
        return fetch(request).then(response => {
          if (!response || !response.ok) {
            return response;
          }
          // Cache successful static responses
          const clone = response.clone();
          cache.put(request, clone);
          return response;
        }).catch(() => {
          // Return cached version on network error
          return cache.match(request);
        });
      });
    })
  );
});
