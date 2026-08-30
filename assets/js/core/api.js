// /api/* 呼び出しラッパー
// SalonOneのキーはサーバー側にのみあり、フロントは常に同一オリジンの /api を叩く。

export class ApiError extends Error {
    constructor(status, code, body) {
        super(`API ${status}: ${code}`);
        this.status = status;
        this.code = code;
        this.body = body || {};
    }
}

async function request(path, { method = 'GET', params, body } = {}) {
    const url = new URL(path, location.origin);
    if (params) {
        for (const [k, v] of Object.entries(params)) {
            if (v !== undefined && v !== null && v !== '') url.searchParams.set(k, String(v));
        }
    }
    const res = await fetch(url, {
        method,
        headers: body ? { 'Content-Type': 'application/json' } : undefined,
        body: body ? JSON.stringify(body) : undefined,
        credentials: 'same-origin',
        cache: 'no-store',
    });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空ボディ */ }
    if (!res.ok) throw new ApiError(res.status, json.error || 'unknown', json);
    return json;
}

export function apiGet(path, params) {
    return request(`/api/data/${path}`, { params });
}

export function authGet(params) {
    return request('/api/auth/session', { params });
}

export function authLogin(body) {
    return request('/api/auth/login', { method: 'POST', body });
}

export function authLogout() {
    return request('/api/auth/logout', { method: 'POST' });
}

export function aiGenerate(prompt) {
    return request('/api/ai', { method: 'POST', body: { prompt } });
}

// 簡易メモ化（同一キーの多重リクエスト防止 + セッション内キャッシュ）
const memo = new Map();
export function apiGetCached(path, params, ttlMs = 120000) {
    const key = path + JSON.stringify(params || {});
    const hit = memo.get(key);
    if (hit && hit.expires > Date.now()) return hit.promise;
    const promise = apiGet(path, params).catch(err => {
        memo.delete(key);
        throw err;
    });
    memo.set(key, { promise, expires: Date.now() + ttlMs });
    return promise;
}

export function clearApiCache() {
    memo.clear();
}
