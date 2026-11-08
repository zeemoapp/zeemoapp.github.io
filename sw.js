/**
 * ZeeMoApp service worker: caches the app shell so the app opens offline.
 * API calls go to Apps Script on another origin and are never cached.
 * Bump VERSION whenever you deploy changed front-end files.
 */

const VERSION = 'zeemo-shell-v5';
const SHELL = [
  './', 'index.html', 'styles.css', 'app.js', 'zeemo-client.js', 'config.js',
  'manifest.webmanifest', 'icons/icon-192.png', 'icons/icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.method !== 'GET' || url.origin !== self.location.origin) return;

  // Network first for the page and config, so deploys and URL changes apply promptly.
  if (req.mode === 'navigate' || url.pathname.endsWith('config.js')) {
    event.respondWith(
      fetch(req)
        .then(res => { const copy = res.clone(); caches.open(VERSION).then(c => c.put(req, copy)); return res; })
        .catch(() => caches.match(req).then(r => r || caches.match('index.html')))
    );
    return;
  }

  // Cache first for the rest of the shell.
  event.respondWith(caches.match(req).then(hit => hit || fetch(req)));
});
