// SalonOne 分析API クライアント（サーバー側専用）
// - APIキーは環境変数 SALONONE_API_KEY からのみ読む（ブラウザには絶対に出さない）
// - インメモリTTLキャッシュ + シングルフライトでレート制限（60req/分）を節約
// - キー未設定時はデモデータにフォールバック

'use strict';

const { demoFetch } = require('./demo');

const BASE_URL = (process.env.SALONONE_BASE_URL || 'https://salonone.net/api/analytics/v1').replace(/\/+$/, '');

// エンドポイント別キャッシュTTL（秒）
const TTL = {
    'meta': 600,
    'sales/summary': 300,
    'marketing/by-channel': 300,
    'marketing/by-staff': 300,
    'marketing/retention': 600,
    'shops': 3600,
    'staffs': 3600,
    'menus': 3600,
    'menu-categories': 3600,
    'visit-sources': 3600,
    'customers': 900,
};

const cache = new Map();     // url -> { expires, value }
const inflight = new Map();  // url -> Promise

class UpstreamError extends Error {
    constructor(status, code, retryAfter) {
        super(`SalonOne API error: ${status} ${code || ''}`);
        this.status = status;
        this.code = code;
        this.retryAfter = retryAfter;
    }
}

function isDemo() {
    return !process.env.SALONONE_API_KEY;
}

function buildUrl(path, params) {
    const url = new URL(`${BASE_URL}/${path}`);
    for (const [k, v] of Object.entries(params || {})) {
        if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, String(v));
    }
    return url.toString();
}

async function rawFetch(path, params) {
    if (isDemo()) return demoFetch(path, params || {});
    const url = buildUrl(path, params);
    const res = await fetch(url, {
        headers: { 'X-SalonOne-Api-Key': process.env.SALONONE_API_KEY, 'Accept': 'application/json' },
    });
    if (res.status === 429) {
        const retryAfter = parseInt(res.headers.get('Retry-After') || '30', 10);
        throw new UpstreamError(429, 'rate_limited', retryAfter);
    }
    if (!res.ok) {
        let code = null;
        try { code = (await res.json()).error?.code; } catch (_) { /* JSONでないエラー応答 */ }
        throw new UpstreamError(res.status, code);
    }
    return res.json();
}

// キャッシュ付き取得。エラー時は期限切れキャッシュがあればそれを返す（stale-if-error）。
async function fetchSalonOne(path, params) {
    const url = buildUrl(path, params);
    const ttl = (TTL[path] || 300) * 1000;
    const hit = cache.get(url);
    const now = Date.now();
    if (hit && hit.expires > now) return hit.value;
    if (inflight.has(url)) return inflight.get(url);

    const p = rawFetch(path, params)
        .then(value => {
            cache.set(url, { expires: Date.now() + ttl, value });
            if (cache.size > 500) {
                for (const [k, v] of cache) { if (v.expires < Date.now()) cache.delete(k); }
            }
            return value;
        })
        .catch(err => {
            if (hit) return hit.value; // 期限切れでも直近値があれば返す
            throw err;
        })
        .finally(() => inflight.delete(url));
    inflight.set(url, p);
    return p;
}

// 一覧系エンドポイントを has_more=false までページング取得（安全上限つき）
// 上限で打ち切った場合は truncated=true を返す（呼び出し側で「一部のみ」表示に使う）
async function fetchAllPages(path, params, maxPages = 20) {
    const all = [];
    let cursor = null;
    let pii = false;
    let truncated = false;
    for (let i = 0; i < maxPages; i++) {
        const page = await fetchSalonOne(path, { ...params, limit: 1000, cursor: cursor || undefined });
        const data = Array.isArray(page) ? page : (page.data || []);
        all.push(...data);
        pii = pii || !!page?.meta?.pii_included;
        if (!page?.meta?.has_more || !page?.meta?.next_cursor) break;
        cursor = page.meta.next_cursor;
        if (i === maxPages - 1) truncated = true;
    }
    return { data: all, pii_included: pii, truncated };
}

// ---- 個人情報の防御的除去 ----
// キーの発行設定に関わらず、ダッシュボードが使わないPIIはプロキシで必ず落とす。
const CUSTOMER_ALLOWED_FIELDS = new Set([
    'id', 'brand_id', 'shop_id', 'staff_id', 'visit_source_id',
    'birth_year', 'age_bracket', 'gender',
    'first_visit_at', 'last_visit_at', 'visit_count',
    'created_at', 'updated_at', 'deleted_at',
]);

function stripCustomerPii(row) {
    const out = {};
    for (const k of Object.keys(row || {})) {
        if (CUSTOMER_ALLOWED_FIELDS.has(k)) out[k] = row[k];
    }
    return out;
}

module.exports = { fetchSalonOne, fetchAllPages, stripCustomerPii, isDemo, UpstreamError };
