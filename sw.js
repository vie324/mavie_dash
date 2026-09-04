/* vie Dashboard Service Worker (v4 - SalonOne版・日報入力刷新)
 * - /api/ は常にネットワーク（キャッシュしない: 認証クッキー付きの動的データのため）
 * - 同一オリジンの静的アセットは stale-while-revalidate
 * - CDN（Chart.js等のバージョン固定URL）は cache-first
 */
const VERSION = 'v4';
const STATIC_CACHE = `vie-static-${VERSION}`;
const CDN_CACHE = `vie-cdn-${VERSION}`;

const PRECACHE = [
    './',
    './index.html',
    './manifest.webmanifest',
    './assets/css/tailwind.css?v=4',
    './assets/css/dashboard.css?v=4',
    './assets/vendor/chart.umd.min.js',
    './assets/vendor/lucide.min.js',
    './assets/logo.png',
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(STATIC_CACHE).then((cache) => cache.addAll(PRECACHE)).then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) =>
            Promise.all(keys.filter((k) => ![STATIC_CACHE, CDN_CACHE].includes(k)).map((k) => caches.delete(k)))
        ).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (event) => {
    const req = event.request;
    if (req.method !== 'GET') return;
    const url = new URL(req.url);

    // APIは常にネットワーク直行（オフライン時はエラーを返す）
    if (url.origin === location.origin && url.pathname.startsWith('/api/')) {
        return;
    }

    // 同一オリジン静的アセット: stale-while-revalidate
    if (url.origin === location.origin) {
        event.respondWith(
            caches.open(STATIC_CACHE).then(async (cache) => {
                const cached = await cache.match(req);
                const fetched = fetch(req).then((res) => {
                    if (res.ok) cache.put(req, res.clone());
                    return res;
                }).catch(() => cached);
                return cached || fetched;
            })
        );
        return;
    }

    // CDN（バージョン固定）: cache-first
    event.respondWith(
        caches.open(CDN_CACHE).then(async (cache) => {
            const cached = await cache.match(req);
            if (cached) return cached;
            const res = await fetch(req);
            if (res.ok) cache.put(req, res.clone());
            return res;
        })
    );
});
