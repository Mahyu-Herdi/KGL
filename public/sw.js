const CACHE_NAME = 'eskavator-sukses-v1';

// Daftar file. Hati-hati, jangan pakai .html untuk index karena Vercel benci itu.
const URLS_TO_CACHE = [
  '/',
  '/dashboard_admin.html',
  '/manifest.json'
];

self.addEventListener('install', (event) => {
  self.skipWaiting();
  
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      // Kita pakai perulangan manual. Kalau 1 file gagal, yang lain TETAP JALAN!
      for (let url of URLS_TO_CACHE) {
        try {
          const response = await fetch(url);
          if (response.ok) {
            await cache.put(url, response);
          }
        } catch (error) {
          console.log('Abaikan error dari Vercel untuk:', url);
        }
      }
    })
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(keys.map((key) => {
        if (key !== CACHE_NAME) return caches.delete(key);
      }));
    })
  );
  return self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  event.respondWith(
    (async () => {
      try {
        // Coba pakai internet
        const networkResponse = await fetch(event.request);
        if (networkResponse && networkResponse.status === 200) {
          const cache = await caches.open(CACHE_NAME);
          cache.put(event.request, networkResponse.clone());
        }
        return networkResponse;
      } catch (error) {
        // MATI INTERNET: Ambil dari Cache
        const cache = await caches.open(CACHE_NAME);
        const cachedResponse = await cache.match(event.request, { ignoreSearch: true });
        if (cachedResponse) return cachedResponse;

        // Kalau halaman gak ketemu pas offline, arahkan ke Root (/)
        if (event.request.mode === 'navigate') {
          return cache.match('/');
        }
      }
    })()
  );
});
