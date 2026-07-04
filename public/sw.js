// SPL DPR service worker.
// Purpose: satisfy PWA installability so "Add to Home Screen" / install works.
// Deliberately NETWORK-ONLY (no caching) so the app can never serve stale content —
// this is why an earlier caching SW was removed. Do not add caches here.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', (event) => {
  // Pass every request straight to the network. A fetch handler is required for
  // installability; keeping it pass-through avoids all stale-cache problems.
  event.respondWith(fetch(event.request));
});
