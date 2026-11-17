// The service worker of the built app (written as dist/sw.js by vite.config.js, which fills in its
// version and the list of its files). It keeps the app's own files, and nothing else, so Intact opens and
// works offline once it has been opened online: on a plane, behind a network that blocks it, or
// installed as an app. Documents are never fetched, cached or stored: they are opened from the
// device's own files and live only in the page.
const CACHE = 'intact-__VERSION__';
const FILES = __FILES__;

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(CACHE).then((cache) => cache.addAll(FILES)).then(() => self.skipWaiting()));
});

// A new version replaces the old one's files.
self.addEventListener('activate', (event) => {
  event.waitUntil(caches.keys()
    .then((keys) => Promise.all(keys.filter((k) => k.startsWith('intact-') && k !== CACHE).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET' || new URL(request.url).origin !== self.location.origin) return;
  if (request.mode === 'navigate') {
    // The page: the network's when online (so a new version shows), this copy when offline.
    event.respondWith(fetch(request).catch(() => caches.match(new URL('./', self.registration.scope).href, { ignoreVary: true })));
    return;
  }
  // The app's files are named by their content: a copy is always the right one. Servers often send
  // Vary: Origin, and scripts are requested with an Origin header the cached request didn't have.
  event.respondWith(caches.match(request, { ignoreVary: true }).then((hit) => hit || fetch(request)));
});
