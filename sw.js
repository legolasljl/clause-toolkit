const CACHE_NAME = 'clause-toolkit-v1';

const ASSETS_TO_CACHE = [
  './',
  './clause_toolkit_offline.html',
  './encrypted_data.js',
  './manifest.json',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './lib/xlsx.full.min.js',
  './lib/mammoth.browser.min.js',
  './lib/pdf.min.js',
  './lib/docx.umd.js',
  './lib/jszip.min.js',
  './lib/exceljs.min.js'
];

// Install: 预缓存所有静态资源
self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => {
      return cache.addAll(ASSETS_TO_CACHE);
    }).then(() => {
      return self.skipWaiting();
    })
  );
});

// Activate: 清理旧版本缓存
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames
          .filter((name) => name.startsWith('clause-toolkit-') && name !== CACHE_NAME)
          .map((name) => caches.delete(name))
      );
    }).then(() => {
      return self.clients.claim();
    })
  );
});

// Fetch: Cache First 策略
self.addEventListener('fetch', (event) => {
  event.respondWith(
    caches.match(event.request).then((cachedResponse) => {
      if (cachedResponse) {
        return cachedResponse;
      }
      return fetch(event.request).then((networkResponse) => {
        if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
          const responseToCache = networkResponse.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(event.request, responseToCache);
          });
        }
        return networkResponse;
      });
    })
  );
});
