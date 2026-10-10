const CACHE_NAME = 'minex-shell-v27.25';
const APP_VERSION = '27.25';
const APP_SHELL = [
  `./?v=${APP_VERSION}`,
  `./index.html?v=${APP_VERSION}`,
  `./styles.css?v=${APP_VERSION}`,
  `./app.js?v=${APP_VERSION}`,
  './jszip.min.js',
  './case_template_base64.js',
  './data_bootstrap_min.js',
  './manifest.webmanifest',
  './icon.svg'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  if (event.request.method !== 'GET') return;
  const requestUrl = new URL(event.request.url);
  if (requestUrl.origin !== self.location.origin) return;
  const isApplicationFile = /\/(?:app|index|sw|styles)\.(?:js|html|css)$/.test(requestUrl.pathname);

  event.respondWith(
    fetch(isApplicationFile ? new Request(event.request, {cache: 'no-store'}) : event.request)
      .then(response => {
        if (response.ok) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, copy));
        }
        return response;
      })
      .catch(() => caches.match(event.request).then(response => response || caches.match('./index.html')))
  );
});
