// サーバー保存（手入力データ・入金突合・シフト・アカウント）のキーバリュー層。
// 保存先は環境変数で自動選択（優先順）:
//   1. Supabase (Postgres) : SUPABASE_URL + SUPABASE_SERVICE_ROLE_KEY → テーブル vie_kv（supabase/schema.sql）
//   2. Upstash Redis (REST): KV_REST_API_URL + KV_REST_API_TOKEN（または UPSTASH_REDIS_REST_URL / TOKEN）
//   3. ローカル開発      : DEV_KV_FILE にJSONファイルパス
// 未設定の場合は保存不可（クライアント側はこの端末のみのlocalStorageに退避する）。

'use strict';

const FETCH_TIMEOUT_MS = 8000;

// ---- Supabase ----
function supabaseConfig() {
    const url = process.env.SUPABASE_URL;
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY;
    if (!url || !key) return null;
    return {
        url: url.replace(/\/+$/, ''),
        key,
        table: process.env.SUPABASE_KV_TABLE || 'vie_kv',
    };
}

// service_role 以外のキー（anon / publishable）を貼ってしまった場合に気付けるようにする
function supabaseKeyProblem(key) {
    if (/^sb_publishable_/.test(key)) return 'publishable キー（ブラウザ用）が設定されています。service_role（secret）キーを設定してください';
    const parts = key.split('.');
    if (parts.length === 3) {
        try {
            const payload = JSON.parse(Buffer.from(parts[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
            if (payload && payload.role && payload.role !== 'service_role') {
                return `${payload.role} キーが設定されています。service_role（secret）キーを設定してください`;
            }
        } catch (_) { /* JWTでなければ判定しない */ }
    }
    return null;
}

function supabaseHeaders(cfg, extra) {
    return {
        apikey: cfg.key,
        Authorization: `Bearer ${cfg.key}`,
        Accept: 'application/json',
        ...extra,
    };
}

async function supabaseGet(cfg, key) {
    const problem = supabaseKeyProblem(cfg.key);
    if (problem) throw new Error(`supabase key: ${problem}`);
    const q = `select=value&key=eq.${encodeURIComponent(key)}&limit=1`;
    const res = await fetch(`${cfg.url}/rest/v1/${cfg.table}?${q}`, {
        headers: supabaseHeaders(cfg),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`supabase get failed: ${res.status} ${(await res.text()).slice(0, 200)}`);
    const rows = await res.json();
    if (!Array.isArray(rows) || rows.length === 0) return null;
    return rows[0].value ?? null;
}

async function supabaseSet(cfg, key, value) {
    const problem = supabaseKeyProblem(cfg.key);
    if (problem) throw new Error(`supabase key: ${problem}`);
    const res = await fetch(`${cfg.url}/rest/v1/${cfg.table}?on_conflict=key`, {
        method: 'POST',
        headers: supabaseHeaders(cfg, {
            'Content-Type': 'application/json',
            Prefer: 'resolution=merge-duplicates,return=minimal',
        }),
        body: JSON.stringify([{ key, value, updated_at: new Date().toISOString() }]),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`supabase set failed: ${res.status} ${(await res.text()).slice(0, 200)}`);
    return true;
}

// ---- Upstash ----
function upstashConfig() {
    const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
    const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
    if (!url || !token) return null;
    return { url: url.replace(/\/+$/, ''), token };
}

async function upstashGet(cfg, key) {
    const res = await fetch(`${cfg.url}/get/${encodeURIComponent(key)}`, {
        headers: { Authorization: `Bearer ${cfg.token}` },
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`kv get failed: ${res.status}`);
    const json = await res.json();
    if (json.result === null || json.result === undefined) return null;
    try { return JSON.parse(json.result); } catch (_) { return null; }
}

async function upstashSet(cfg, key, value) {
    const res = await fetch(`${cfg.url}/set/${encodeURIComponent(key)}`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${cfg.token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(value),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`kv set failed: ${res.status}`);
    return true;
}

// ---- ローカル開発用ファイル ----
function devFile() {
    return process.env.DEV_KV_FILE || null;
}

function devLoad() {
    const fs = require('fs');
    try { return JSON.parse(fs.readFileSync(devFile(), 'utf8')); } catch (_) { return {}; }
}

// ---- 共通 ----
function kvBackend() {
    if (supabaseConfig()) return 'supabase';
    if (upstashConfig()) return 'upstash';
    if (devFile()) return 'file';
    return null;
}

function kvAvailable() {
    return kvBackend() !== null;
}

// 設定タブの連携状態に表示する情報（キーの中身は返さない）
function kvStatus() {
    const backend = kvBackend();
    const status = { backend, label: { supabase: 'Supabase', upstash: 'Upstash Redis', file: 'ローカルファイル' }[backend] || null, warning: null };
    if (backend === 'supabase') status.warning = supabaseKeyProblem(supabaseConfig().key);
    return status;
}

async function kvGet(key) {
    const sb = supabaseConfig();
    if (sb) return supabaseGet(sb, key);
    const up = upstashConfig();
    if (up) return upstashGet(up, key);
    if (devFile()) return devLoad()[key] ?? null;
    throw new Error('kv_unconfigured');
}

async function kvSet(key, value) {
    const sb = supabaseConfig();
    if (sb) return supabaseSet(sb, key, value);
    const up = upstashConfig();
    if (up) return upstashSet(up, key, value);
    if (devFile()) {
        const fs = require('fs');
        const all = devLoad();
        all[key] = value;
        fs.writeFileSync(devFile(), JSON.stringify(all));
        return true;
    }
    throw new Error('kv_unconfigured');
}

module.exports = { kvAvailable, kvBackend, kvStatus, kvGet, kvSet };
