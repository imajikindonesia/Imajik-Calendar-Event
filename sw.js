// ================= IMAJIK Command Center — Service Worker =================
const CACHE_VERSION = 'imajik-cc-v1';
const APP_SHELL_CACHE = CACHE_VERSION + '-shell';
const CDN_CACHE = CACHE_VERSION + '-cdn';

// App shell — file lokal yang di-cache saat install
const APP_SHELL_FILES = [
  './',
  './index.html',
  './style.css',
  './manifest.json',
  './logo-imajik.png',
  './foto-ceo.jpg'
];

// CDN resources — di-cache saat pertama kali di-request
const CDN_ORIGINS = [
  'fonts.googleapis.com',
  'fonts.gstatic.com',
  'cdnjs.cloudflare.com'
];

// API endpoint — selalu network-first
const API_ORIGIN = 'script.google.com';

// ================= INSTALL =================
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(APP_SHELL_CACHE)
      .then(cache => cache.addAll(APP_SHELL_FILES))
      .then(() => self.skipWaiting())
      .catch(err => {
        console.warn('[SW] Install: gagal cache beberapa file, lanjut saja.', err);
        return self.skipWaiting();
      })
  );
});

// ================= ACTIVATE =================
// Hapus cache lama saat versi baru diaktifkan
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys
          .filter(key => !key.startsWith(CACHE_VERSION))
          .map(key => {
            console.log('[SW] Hapus cache lama:', key);
            return caches.delete(key);
          })
      );
    }).then(() => self.clients.claim())
  );
});

// ================= FETCH =================
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // 1. API calls (Apps Script) — selalu network-first, tidak di-cache
  if (url.hostname === API_ORIGIN) {
    event.respondWith(
      fetch(event.request).catch(() => {
        return new Response(JSON.stringify({
          success: false,
          error: 'Offline — tidak bisa menghubungi server.'
        }), {
          headers: { 'Content-Type': 'application/json' }
        });
      })
    );
    return;
  }

  // 2. CDN resources (fonts, icons) — cache-first, fetch kalau belum ada
  if (CDN_ORIGINS.some(origin => url.hostname.includes(origin))) {
    event.respondWith(
      caches.match(event.request).then(cached => {
        if (cached) return cached;
        return fetch(event.request).then(response => {
          if (response && response.status === 200) {
            const clone = response.clone();
            caches.open(CDN_CACHE).then(cache => cache.put(event.request, clone));
          }
          return response;
        }).catch(() => cached); // Fallback ke cache kalau fetch gagal
      })
    );
    return;
  }

  // 3. App shell — cache-first, fallback ke network
  event.respondWith(
    caches.match(event.request).then(cached => {
      // Return cache segera, tapi juga update cache di background (stale-while-revalidate)
      const fetchPromise = fetch(event.request).then(response => {
        if (response && response.status === 200 && event.request.method === 'GET') {
          const clone = response.clone();
          caches.open(APP_SHELL_CACHE).then(cache => cache.put(event.request, clone));
        }
        return response;
      }).catch(() => null);

      return cached || fetchPromise;
    })
  );
});

// ================= MESSAGE HANDLER =================
// Untuk komunikasi dari main thread (misal: force update cache)
self.addEventListener('message', event => {
  if (event.data === 'skipWaiting') {
    self.skipWaiting();
  }
});
