// 領収書の写真などのファイル保存（Supabase Storage）。
// vie_kv と同じ SUPABASE_URL + SUPABASE_SERVICE_ROLE_KEY を使い、非公開バケットに保存する。
// 閲覧は API が発行する短時間の署名付きURL経由のみ（ブラウザ用キーからは一切見えない）。

'use strict';

const FETCH_TIMEOUT_MS = 15000;
const BUCKET = process.env.SUPABASE_RECEIPT_BUCKET || 'vie-receipts';
const MAX_BYTES = 4 * 1024 * 1024;
const ALLOWED = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' };
// 保存先パス: receipts/<店舗ID>/<YYYY-MM>/r_<16桁hex>.<拡張子>
const RECEIPT_PATH_RE = /^receipts\/(\d+)\/\d{4}-\d{2}\/r_[0-9a-f]{16}\.(jpg|png|webp)$/;
const MAX_RECEIPTS = 5;

function config() {
    const url = process.env.SUPABASE_URL;
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY;
    if (!url || !key) return null;
    return { url: url.replace(/\/+$/, ''), key };
}

function storageAvailable() {
    return config() !== null;
}

function headers(cfg, extra) {
    return { apikey: cfg.key, Authorization: `Bearer ${cfg.key}`, ...extra };
}

let bucketReady = false;
async function ensureBucket(cfg) {
    if (bucketReady) return;
    const base = `${cfg.url}/storage/v1/bucket`;
    const got = await fetch(`${base}/${BUCKET}`, { headers: headers(cfg), signal: AbortSignal.timeout(FETCH_TIMEOUT_MS) });
    if (got.ok) { bucketReady = true; return; }
    const res = await fetch(base, {
        method: 'POST',
        headers: headers(cfg, { 'Content-Type': 'application/json' }),
        body: JSON.stringify({ id: BUCKET, name: BUCKET, public: false, file_size_limit: MAX_BYTES, allowed_mime_types: Object.keys(ALLOWED) }),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (res.ok || res.status === 409) { bucketReady = true; return; }
    const text = await res.text();
    if (/already exists/i.test(text)) { bucketReady = true; return; }
    throw new Error(`storage bucket create failed: ${res.status} ${text.slice(0, 200)}`);
}

// 保存。戻り値は保存先パス
async function putObject(path, buffer, contentType) {
    const cfg = config();
    if (!cfg) throw new Error('storage_unconfigured');
    await ensureBucket(cfg);
    const res = await fetch(`${cfg.url}/storage/v1/object/${BUCKET}/${path}`, {
        method: 'POST',
        headers: headers(cfg, { 'Content-Type': contentType, 'x-upsert': 'false', 'Cache-Control': 'private, max-age=31536000' }),
        body: buffer,
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`storage upload failed: ${res.status} ${(await res.text()).slice(0, 200)}`);
    return path;
}

// 短時間だけ有効な閲覧用URL
async function signedUrl(path, expiresInSec = 600) {
    const cfg = config();
    if (!cfg) throw new Error('storage_unconfigured');
    const res = await fetch(`${cfg.url}/storage/v1/object/sign/${BUCKET}/${path}`, {
        method: 'POST',
        headers: headers(cfg, { 'Content-Type': 'application/json' }),
        body: JSON.stringify({ expiresIn: expiresInSec }),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (res.status === 404 || res.status === 400) return null;
    if (!res.ok) throw new Error(`storage sign failed: ${res.status} ${(await res.text()).slice(0, 200)}`);
    const json = await res.json();
    const rel = json.signedURL || json.signedUrl;
    if (!rel) return null;
    return rel.startsWith('http') ? rel : `${cfg.url}/storage/v1${rel.startsWith('/') ? '' : '/'}${rel}`;
}

async function deleteObjects(paths) {
    const cfg = config();
    if (!cfg || !paths.length) return;
    const res = await fetch(`${cfg.url}/storage/v1/object/${BUCKET}`, {
        method: 'DELETE',
        headers: headers(cfg, { 'Content-Type': 'application/json' }),
        body: JSON.stringify({ prefixes: paths }),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`storage delete failed: ${res.status}`);
}

module.exports = { BUCKET, MAX_BYTES, ALLOWED, RECEIPT_PATH_RE, MAX_RECEIPTS, storageAvailable, putObject, signedUrl, deleteObjects };
