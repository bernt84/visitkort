/* Visitkort – service worker (offline + installation)
   Ret VERSION, når du uploader nye filer, så telefonerne henter dem. */
const VERSION = 'visitkort-v5';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon-192.png', './icon-512.png', './icon-512-maskable.png', './apple-touch-icon.png',
               './qrcode.min.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION)
    .then(c => Promise.all(SHELL.map(u => c.add(new Request(u, { mode: 'same-origin' })).catch(() => {}))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  // Data fra Google hentes altid live
  if (/(^|\.)google\.com$|googleusercontent\.com$/.test(url.hostname)) return;

  // Siden selv: hent ny version, brug gemt version uden net
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => {
      if (r.ok) { const cp = r.clone(); caches.open(VERSION).then(c => c.put('./index.html', cp)); }
      return r;
    }).catch(() => caches.match('./index.html')));
    return;
  }

  // Ikoner, QR-bibliotek, skrifttyper: gemt først
  e.respondWith(caches.match(req).then(hit => hit || fetch(req).then(r => {
    if (r.ok || r.type === 'opaque') { const cp = r.clone(); caches.open(VERSION).then(c => c.put(req, cp)); }
    return r;
  })));
});
