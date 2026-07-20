// Bump APP_VERSION whenever you want to force a clean cache wipe (it names the cache and drives
// old-cache cleanup on activate). NOTE: app freshness no longer depends on remembering to bump it.
// The app shell is served stale-while-revalidate, so a new index.html deployed to GitHub Pages is
// picked up automatically on the next launch — the stale cache that used to pin old builds is gone.
const APP_VERSION = 'v0.8.2';
const CACHE_NAME = `pack-tracker-${APP_VERSION}`;

// Core files the UI needs to boot. Every ES module index.html imports must be here — a module
// import that misses the cache is a hard failure at exactly the wrong moment (a tablet cold-started
// in the field with no signal), so keep this list in sync with the app's local imports.
const CORE_ASSETS = [
    './',
    './index.html',
    './manifest.json',
    './firebase-config.js',
    './sheet-sync-config.js',
    './icon-192.png',
    './icon-512.png'
];

// Firebase SDK is imported from Google's CDN. These URLs are version-pinned (immutable for a given
// version), so we pre-cache them and serve them cache-first. Bump them if firebase-config.js moves.
const VENDOR_ASSETS = [
    'https://www.gstatic.com/firebasejs/10.7.1/firebase-app.js',
    'https://www.gstatic.com/firebasejs/10.7.1/firebase-firestore.js'
];

// 1. Install — pre-cache the app shell + vendor SDK.
self.addEventListener('install', (event) => {
    event.waitUntil((async () => {
        const cache = await caches.open(CACHE_NAME);
        await cache.addAll(CORE_ASSETS);
        // Best-effort + individually caught: addAll is all-or-nothing, so one CDN hiccup would
        // otherwise abort the whole install and leave the device with no worker at all.
        await Promise.all(VENDOR_ASSETS.map(url =>
            cache.add(url).catch(err => console.warn('[SW] Could not pre-cache', url, err))
        ));
    })());
    // Deliberately NO skipWaiting() here. A freshly installed worker waits instead of taking over,
    // so an update can never reload a technician mid-service. The page detects the waiting worker,
    // shows an "Update now" banner, and posts SKIP_WAITING only when the user chooses to apply it.
});

// 2. Activate — drop any caches from previous versions, then take control.
self.addEventListener('activate', (event) => {
    event.waitUntil((async () => {
        const names = await caches.keys();
        await Promise.all(names.map(n => (n !== CACHE_NAME ? caches.delete(n) : null)));
        await self.clients.claim();
    })());
});

// The page's update banner posts this when the user taps "Update now".
self.addEventListener('message', (event) => {
    if (event.data && event.data.type === 'SKIP_WAITING') self.skipWaiting();
});

function isVendorAsset(url) {
    return url.startsWith('https://www.gstatic.com/firebasejs/');
}

async function putInCache(request, response) {
    if (response && response.ok) {
        const cache = await caches.open(CACHE_NAME);
        cache.put(request, response.clone());
    }
    return response;
}

// Serve the cached copy immediately (instant cold start, works fully offline) while fetching a
// fresh copy in the background to update the cache for next launch. This is what makes GitHub Pages
// pushes propagate without any manual cache-version bump.
async function staleWhileRevalidate(request) {
    const cache = await caches.open(CACHE_NAME);
    const cached = await cache.match(request)
        || (request.mode === 'navigate' ? await cache.match('./index.html') : undefined);

    const networkFetch = fetch(request)
        .then(response => { if (response && response.ok) cache.put(request, response.clone()); return response; })
        .catch(() => undefined);

    // Cached first if we have it; otherwise wait on the network; fall back to the shell for navigations.
    return cached || (await networkFetch) || (await cache.match('./index.html'));
}

// 3. Fetch — route by request type.
self.addEventListener('fetch', (event) => {
    const request = event.request;
    const url = request.url;

    // Never intercept Firebase's live traffic — the SDK handles its own offline queueing.
    if (url.includes('firestore.googleapis.com') || url.includes('identitytoolkit') || url.includes('firebaseio.com')) return;

    // Only GET requests are cacheable; let everything else hit the network untouched.
    if (request.method !== 'GET') return;

    // Version-pinned SDK is immutable per URL — cache-first is safe and fastest.
    if (isVendorAsset(url)) {
        event.respondWith(
            caches.match(request).then(cached => cached || fetch(request).then(r => putInCache(request, r)).catch(() => cached))
        );
        return;
    }

    // App shell, same-origin assets, and navigations: stale-while-revalidate.
    event.respondWith(staleWhileRevalidate(request));
});
