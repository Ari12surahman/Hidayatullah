const CACHE_NAME = 'mutabaah-admin-cache-v3';

self.addEventListener('install', function(e) {
    self.skipWaiting();
});

self.addEventListener('activate', function(e) {
    e.waitUntil(
        caches.keys().then(function(cacheNames) {
            return Promise.all(
                cacheNames.map(function(cacheName) {
                    if (cacheName !== CACHE_NAME) {
                        return caches.delete(cacheName);
                    }
                })
            );
        }).then(function() {
            return self.clients.claim();
        })
    );
});

self.addEventListener('fetch', function(e) {
    if (e.request.method !== 'GET' || e.request.url.indexOf('script.google.com') !== -1 || e.request.url.indexOf('googleusercontent.com') !== -1 || e.request.url.indexOf('chrome-extension') !== -1) {
        return; 
    }

    e.respondWith(
        caches.open(CACHE_NAME).then(function(cache) {
            return cache.match(e.request).then(function(response) {
                const fetchPromise = fetch(e.request).then(function(networkResponse) {
                    if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'basic') {
                        cache.put(e.request, networkResponse.clone());
                    } else if (networkResponse && networkResponse.status === 200 && networkResponse.type === 'cors') {
                        cache.put(e.request, networkResponse.clone());
                    }
                    return networkResponse;
                }).catch(function(err) {
                    console.log('Network request failed', err);
                });
                
                return response || fetchPromise;
            });
        })
    );
});
