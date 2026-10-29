/* AgriLedger offline helper: keeps the app files on the phone so it opens with no signal.
   Data is never stored here (it lives in the phone's database); Google Sheet traffic is never cached.
   Build 20261029-174916 */
const APP_CACHE = 'agriledger-app-20261029-174916';
const EXT_CACHE = 'agriledger-ext-v1';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icons/icon-192.png', './icons/icon-512.png', './icons/maskable-512.png', './icons/favicon-64.png'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(APP_CACHE)
    .then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' }))))
    .then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  e.waitUntil(caches.keys()
    .then(keys => Promise.all(keys.filter(k => k !== APP_CACHE && k !== EXT_CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (/(^|\.)script\.google(usercontent)?\.com$|googleusercontent\.com$/.test(url.hostname)) return;

  if (url.origin === self.location.origin) {
    // open instantly from the phone; check the server for a newer copy in the background
    e.respondWith(caches.open(APP_CACHE).then(async cache => {
      const key = req.mode === 'navigate' ? './index.html' : req;
      const hit = await cache.match(key, { ignoreSearch: true });
      const fresh = fetch(req, { cache: 'no-cache' }).then(res => { if (res.ok) cache.put(key, res.clone()); return res; }).catch(() => hit);
      return hit || fresh;
    }));
    return;
  }

  if (/fonts\.googleapis\.com|fonts\.gstatic\.com|cdnjs\.cloudflare\.com/.test(url.hostname)) {
    e.respondWith(caches.open(EXT_CACHE).then(async cache => {
      const hit = await cache.match(req);
      if (hit) return hit;
      const res = await fetch(req);
      if (res.ok || res.type === 'opaque') cache.put(req, res.clone());
      return res;
    }));
  }
});
