const CACHE_NAME = 'pack-tracker-v2.4';

// List all the core files your app needs to load the UI.
// Include any local icons or assets here if you add them later.
// NOTE: every ES module index.html imports has to be in here. A module import that misses the
// cache is a hard failure — the app won't boot at all, which is precisely the moment (a tablet
// cold-started in the field with no signal) that it has to.
const ASSETS_TO_CACHE = [
    './',
    './index.html',
    './manifest.json',
    './firebase-config.js',
    './sheet-sync-config.js',
    './icon-192.png',
    './icon-512.png'
];

// The Firebase SDK is imported from Google's CDN, so without pre-caching it the app can only
// cold-start offline if the browser's HTTP cache happens to still hold it. These URLs are
// version-pinned, so caching them is safe — bump them if firebase-config.js moves version.
const VENDOR_ASSETS = [
    'https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js',
    'https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js'
];

// 1. Install Event: Cache the essential files (App Shell)
self.addEventListener('install', (event) => {
    event.waitUntil((async () => {
        const cache = await caches.open(CACHE_NAME);
        console.log('[Service Worker] Pre-caching offline assets');
        await cache.addAll(ASSETS_TO_CACHE);
        // Best-effort and individually caught: addAll is all-or-nothing, so one CDN hiccup
        // would otherwise abort the whole install and leave the tablet with no worker at all.
        await Promise.all(VENDOR_ASSETS.map(url =>
            cache.add(url).catch(err => console.warn('[Service Worker] Could not pre-cache', url, err))
        ));
    })());
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
