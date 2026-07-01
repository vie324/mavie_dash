/* vie Dashboard Service Worker
 * - 静的アセット: stale-while-revalidate（オフラインでも前回表示を即座に出す）
 * - GAS API: network-first + キャッシュフォールバック（オフライン時は前回データ）
 * - CDN: cache-first
 */
const VERSION = 'v5';
const STATIC_CACHE = `vie-static-${VERSION}`;
const API_CACHE = `vie-api-${VERSION}`;
const CDN_CACHE = `vie-cdn-${VERSION}`;

const PRECACHE = [
    './',
    './index.html',
    './manifest.webmanifest',
    './assets/css/tailwind.css',
    './assets/css/dashboard.css',
    './assets/js/backend.js',
    './assets/js/dashboard.js',
    './assets/js/enhancements.js',
    './assets/js/reviews.js',
    './assets/logo.png',
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(STATIC_CACHE)
            .then(cache => Promise.allSettled(PRECACHE.map(u => cache.add(u))))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys => Promise.all(
            keys.filter(k => ![STATIC_CACHE, API_CACHE, CDN_CACHE].includes(k)).map(k => caches.delete(k))
        )).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (event) => {
    const req = event.request;
    if (req.method !== 'GET') return; // POST(保存系)は素通し

    const url = new URL(req.url);

    // GAS API: network-first（オフライン時のみキャッシュ）
    if (url.hostname.endsWith('script.google.com') || url.hostname.endsWith('googleusercontent.com')) {
        event.respondWith(
            fetch(req)
                .then(res => {
                    if (res && res.ok) {
                        const clone = res.clone();
                        caches.open(API_CACHE).then(c => c.put(req, clone));
                    }
                    return res;
                })
                .catch(() => caches.match(req))
        );
        return;
    }

    // CDN (fonts / chart.js / lucide): cache-first
    if (url.origin !== location.origin) {
        event.respondWith(
            caches.match(req).then(hit => hit || fetch(req).then(res => {
                const clone = res.clone();
                caches.open(CDN_CACHE).then(c => c.put(req, clone));
                return res;
            }).catch(() => hit))
        );
        return;
    }

    // 同一オリジン静的アセット: stale-while-revalidate
    event.respondWith(
        caches.match(req).then(hit => {
            const fetched = fetch(req).then(res => {
                if (res && res.ok) {
                    const clone = res.clone();
                    caches.open(STATIC_CACHE).then(c => c.put(req, clone));
                }
                return res;
            }).catch(() => hit);
            return hit || fetched;
        })
    );
});
