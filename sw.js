const CACHE_NAME = 'pack-tracker-v2.1';

// List all the core files your app needs to load the UI.
// Include any local icons or assets here if you add them later.
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './manifest.json',
    './firebase-config.js'
];

// 1. Install Event: Cache the essential files (App Shell)
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('[Service Worker] Pre-caching offline assets');
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
    // Force the waiting service worker to become the active one immediately
    self.skipWaiting(); 
});

// 2. Activate Event: Clean up old caches when you update the version number
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cacheName) => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('[Service Worker] Clearing old cache:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
    // Ensure the service worker takes control of the page immediately
    self.clients.claim(); 
});

// 3. Fetch Event: Intercept network requests
self.addEventListener('fetch', (event) => {
    // CRITICAL: Do not intercept Firebase API calls. 
    // Let Firebase's native SDK handle its own offline database queuing.
    if (event.request.url.includes('firestore.googleapis.com') || event.request.url.includes('identitytoolkit')) {
        return; 
    }

    // Cache-First Strategy for app assets
    event.respondWith(
        caches.match(event.request).then((cachedResponse) => {
            // Return the cached file if we have it
            if (cachedResponse) {
                return cachedResponse;
            }
            
            // Otherwise, try to fetch it from the network
            return fetch(event.request).catch(() => {
                console.log('[Service Worker] Fetch failed, user is likely offline.');
                // You could optionally return a fallback offline page here if needed
            });
        })
    );
});
