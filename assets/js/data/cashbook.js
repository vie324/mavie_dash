// 現金出納帳のデータ層（/api/cashbook）
// 帳簿残高・過不足の計算はサーバー側（api/_lib/cashbook.js）。ここは取得・保存・キャッシュのみ。

import { state, emit } from '../core/state.js';
import { ApiError } from '../core/api.js';
import { reconMethods } from './manual.js';

async function request(path, options) {
    const res = await fetch(path, { credentials: 'same-origin', cache: 'no-store', ...options });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空 */ }
    if (!res.ok) throw new ApiError(res.status, json.error || 'unknown', json);
    return json;
}

function store() {
    if (!state.data.cashbook) state.data.cashbook = {};
    return state.data.cashbook;
}

const keyOf = (shopId, month) => `${shopId}:${month}`;

export function getCashbook(shopId, month) {
    return store()[keyOf(shopId, month)] || null;
}

export async function loadCashbook(shopId, month, { log = false } = {}) {
    const res = await request(`/api/cashbook?shop=${encodeURIComponent(shopId)}&month=${month}${log ? '&log=1' : ''}`);
    const prev = store()[keyOf(shopId, month)];
    // ログは別取得のことがあるので、取得しなかったときは前回分を残す
    if (!log && prev?.log) res.log = prev.log;
    store()[keyOf(shopId, month)] = res;
    emit('data:cashbook');
    return res;
}

export async function cashbookAction(body) {
    const res = await request('/api/cashbook', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
    });
    if (res.shopId !== undefined && res.month) {
        const prev = store()[keyOf(res.shopId, res.month)];
        store()[keyOf(res.shopId, res.month)] = { ...res, log: prev?.log, logStale: true };
    }
    statusCache = null;
    emit('data:cashbook');
    return res;
}

// ---- 領収書の写真（/api/receipt → Supabase Storage）----
// 画像は端末側で縮小（長辺 1600px・JPEG）してから送る。戻り値は entry.add / entry.edit の receipts に入れる 1件分
export async function uploadReceipt(shopId, month, file) {
    const blob = await compressImage(file);
    const q = new URLSearchParams({ shop: String(shopId), month, type: blob.type, name: (file.name || '').slice(0, 80) });
    const res = await fetch(`/api/receipt?${q}`, {
        method: 'POST',
        credentials: 'same-origin',
        headers: { 'Content-Type': 'application/octet-stream' },
        body: blob,
    });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空 */ }
    if (!res.ok) throw new ApiError(res.status, json.error || 'unknown', json);
    return { id: json.id, path: json.path, type: json.type, size: json.size, name: json.name || file.name || '', at: new Date().toISOString() };
}

export function receiptUrl(shopId, path) {
    return `/api/receipt?shop=${encodeURIComponent(shopId)}&path=${encodeURIComponent(path)}`;
}

const MAX_EDGE = 1600;
const MAX_UPLOAD = 4 * 1024 * 1024;
async function compressImage(file) {
    const type = String(file.type || '');
    let bitmap = null;
    try {
        bitmap = await loadImage(file);
    } catch (_) {
        // 端末が解釈できない形式（HEIC など）はそのまま送る（サーバー側で種類を確認する）
        if (file.size <= MAX_UPLOAD) return file;
        throw new ApiError(413, 'too_large', { detail: '画像が大きすぎます（4MBまで）' });
    }
    const scale = Math.min(1, MAX_EDGE / Math.max(bitmap.width, bitmap.height));
    if (scale === 1 && /^image\/(jpeg|png|webp)$/.test(type) && file.size <= 1.5 * 1024 * 1024) { bitmap.close?.(); return file; }
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(bitmap.width * scale));
    canvas.height = Math.max(1, Math.round(bitmap.height * scale));
    canvas.getContext('2d').drawImage(bitmap, 0, 0, canvas.width, canvas.height);
    bitmap.close?.();
    const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.82));
    if (!blob) throw new ApiError(400, 'invalid_request', { detail: '画像を変換できませんでした' });
    return blob;
}

function loadImage(file) {
    // createImageBitmap は EXIF の向きを反映できる。使えない環境では <img> で読む
    if (typeof createImageBitmap === 'function') {
        return createImageBitmap(file, { imageOrientation: 'from-image' }).catch(() => createImageBitmap(file));
    }
    return new Promise((resolve, reject) => {
        const url = URL.createObjectURL(file);
        const img = new Image();
        img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
        img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('decode failed')); };
        img.src = url;
    });
}

// ---- ホーム・入金突合用のステータス（全店舗分・1分キャッシュ）----
let statusCache = null; // { at, promise }
export function loadCashStatus({ force = false } = {}) {
    if (!force && statusCache && Date.now() - statusCache.at < 60000) return statusCache.promise;
    const promise = request('/api/cashbook?status=1').then(res => {
        state.data.cashStatus = res;
        emit('data:cashstatus');
        return res;
    }).catch(err => {
        statusCache = null;
        throw err;
    });
    statusCache = { at: Date.now(), promise };
    return promise;
}

// 締め済みの日の「実際の現金売上」（入金突合の現金欄の自動入力に使う）
//   = 実査額 − 前日繰越 − 手入力の入金 + 手入力の出金
export function closedCashFor(shopId, date) {
    const view = getCashbook(shopId, String(date).slice(0, 7));
    const row = view?.rows?.find(r => r.date === date);
    if (row?.closed) return { actualCash: row.closed.actualCash, cashMethodIds: (row.closed.cashMethods || []).map(m => m.id) };
    const st = state.data.cashStatus?.shops?.find(s => String(s.shopId) === String(shopId));
    return st?.days?.[date] || null;
}

// 出納帳を使っている店舗か（ステータス取得済みのときだけ判定できる）
export function cashbookConfigured(shopId) {
    const st = state.data.cashStatus?.shops?.find(s => String(s.shopId) === String(shopId));
    return st ? !!st.configured : null;
}

// 入金突合の入力に、出納帳で締めた日の現金を差し込む
//   SalonOneのその日の現金の支払い方法が1つだけのとき、その欄を「実際の現金売上」で埋める
//   戻り値: { entry: 差し込み後の入力, autoKey: 自動で埋めた欄（'m<支払方法ID>'）or null }
export function reconEntryWithCash(shopId, date, dayRow, entry) {
    const base = entry || {};
    const c = closedCashFor(shopId, date);
    if (!c) return { entry: base, autoKey: null };
    const cashRows = reconMethods(dayRow).filter(p => (c.cashMethodIds || []).includes(Number(p.payment_method_id)));
    if (cashRows.length !== 1) return { entry: base, autoKey: null };
    const key = `m${cashRows[0].payment_method_id}`;
    return { entry: { ...base, [key]: c.actualCash }, autoKey: key };
}
