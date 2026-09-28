// AGA Sightings offline worker.
// When you change any app file, raise the version number below so phones pick it up.
const CACHE = 'aga-sightings-v1';
const SHELL = ['./', './index.html', './config.js', './manifest.json', './icon-192.png', './icon-512.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))));
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                       // saving sightings always goes to the network
  const url = new URL(req.url);
  if (/script\.google|googleusercontent/.test(url.hostname)) return;

  // App files: open instantly from the phone, refresh the copy in the background when online
  if (url.origin === self.location.origin) {
    e.respondWith(caches.open(CACHE).then(cache =>
      cache.match(req, { ignoreSearch: true }).then(hit => {
        const net = fetch(req).then(res => { if (res.ok) cache.put(req, res.clone()); return res; }).catch(() => hit || cache.match('./'));
        return hit || net;
      })
    ));
    return;
  }

  // Fonts: keep a copy so the app looks the same offline
  if (/fonts\.(googleapis|gstatic)\.com/.test(url.hostname)) {
    e.respondWith(caches.open(CACHE).then(cache =>
      cache.match(req).then(hit => hit || fetch(req).then(res => { cache.put(req, res.clone()); return res; }))
    ));
  }
});
