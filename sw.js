self.addEventListener('install', function(e) {
    self.skipWaiting();
});

self.addEventListener('activate', function(e) {
    e.waitUntil(self.clients.claim());
});

self.addEventListener('fetch', function(e) {
    if (e.request.url.indexOf('script.google.com') !== -1 || e.request.method === 'POST') {
        return; // Bypass Service Worker for API & POST requests
    }
    e.respondWith(
        fetch(e.request).catch(function() {
            return new Response('Offline Mode - Mutabaah Digital');
        })
    );
});
