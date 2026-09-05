// サーバー保存（手入力データ・入金突合・シフト・目標・アカウント）のキーバリュー層。
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

function supabaseTable(cfg) {
    return `${cfg.url}/rest/v1/${cfg.table}`;
}

function assertSupabaseKey(cfg) {
    const problem = supabaseKeyProblem(cfg.key);
    if (problem) throw new Error(`supabase key: ${problem}`);
}

async function supabaseGet(cfg, key) {
    assertSupabaseKey(cfg);
    const q = `select=value&key=eq.${encodeURIComponent(key)}&limit=1`;
    const res = await fetch(`${supabaseTable(cfg)}?${q}`, {
        headers: supabaseHeaders(cfg),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`supabase get failed: ${res.status} ${(await res.text()).slice(0, 200)}`);
    const rows = await res.json();
    if (!Array.isArray(rows) || rows.length === 0) return null;
    return rows[0].value ?? null;
}

async function supabaseSet(cfg, key, value) {
    assertSupabaseKey(cfg);
    const res = await fetch(`${supabaseTable(cfg)}?on_conflict=key`, {
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

// ロック行（key = "lock:<対象キー>", value = {until: 期限(ミリ秒・13桁ゼロ埋め)}）
//   1) ON CONFLICT DO NOTHING の挿入 → 挿入できた行だけが返るので、返れば取得成功
//   2) 既存ロックの期限切れ（until < now）なら WHERE 付き UPDATE で原子的に奪う
function lockStamp(ms) {
    return String(ms).padStart(13, '0');
}

async function supabaseTryLock(cfg, lockKey, ttlMs) {
    assertSupabaseKey(cfg);
    const now = Date.now();
    const row = { key: lockKey, value: { until: lockStamp(now + ttlMs) }, updated_at: new Date(now).toISOString() };
    let res = await fetch(`${supabaseTable(cfg)}?on_conflict=key`, {
        method: 'POST',
        headers: supabaseHeaders(cfg, { 'Content-Type': 'application/json', Prefer: 'resolution=ignore-duplicates,return=representation' }),
        body: JSON.stringify([row]),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`supabase lock failed: ${res.status}`);
    let rows = await res.json();
    if (Array.isArray(rows) && rows.length > 0) return true;

    const q = `key=eq.${encodeURIComponent(lockKey)}&value->>until=lt.${lockStamp(now)}`;
    res = await fetch(`${supabaseTable(cfg)}?${q}`, {
        method: 'PATCH',
        headers: supabaseHeaders(cfg, { 'Content-Type': 'application/json', Prefer: 'return=representation' }),
        body: JSON.stringify({ value: row.value, updated_at: row.updated_at }),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`supabase lock takeover failed: ${res.status}`);
    rows = await res.json();
    return Array.isArray(rows) && rows.length > 0;
}

async function supabaseUnlock(cfg, lockKey) {
    const res = await fetch(`${supabaseTable(cfg)}?key=eq.${encodeURIComponent(lockKey)}`, {
        method: 'DELETE',
        headers: supabaseHeaders(cfg, { Prefer: 'return=minimal' }),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`supabase unlock failed: ${res.status}`);
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

// 任意のRedisコマンド（Upstash REST: POST / に ["CMD", ...args] を送る）。Upstash 以外では使えない
async function kvCommand(args) {
    const cfg = upstashConfig();
    if (!cfg) throw new Error('kv_command_unsupported');
    const res = await fetch(cfg.url, {
        method: 'POST',
        headers: { Authorization: `Bearer ${cfg.token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(args.map(a => String(a))),
        signal: AbortSignal.timeout(FETCH_TIMEOUT_MS),
    });
    if (!res.ok) throw new Error(`kv command failed: ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(`kv command error: ${json.error}`);
    return json.result;
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

// ---- 簡易ロック（読み取り→加工→書き込みの競合防止）----
// 複数スタッフが同時に日報を保存しても後勝ちで消えないよう、キーごとに短時間のロックを取る。
// 取れなければ少し待って再試行。ロック基盤に問題があっても保存自体は止めない（フェイルオープン）。
//   Supabase: vie_kv の "lock:<key>" 行（期限付き）／ Upstash: SET NX PX ／ 開発サーバー: メモリ
const devLocks = new Map();
const noLock = { release: async () => {} };

async function kvLock(key, { ttlMs = 5000, retries = 12, waitMs = 120 } = {}) {
    const lockKey = `lock:${key}`;
    const backend = kvBackend();
    if (!backend) return noLock;
    for (let i = 0; i <= retries; i++) {
        let acquired = false;
        try {
            if (backend === 'supabase') {
                acquired = await supabaseTryLock(supabaseConfig(), lockKey, ttlMs);
            } else if (backend === 'upstash') {
                acquired = (await kvCommand(['SET', lockKey, '1', 'NX', 'PX', ttlMs])) === 'OK';
            } else {
                // 単一プロセスの開発サーバー: メモリ上で擬似ロック
                const until = devLocks.get(lockKey) || 0;
                if (until < Date.now()) { devLocks.set(lockKey, Date.now() + ttlMs); acquired = true; }
            }
        } catch (e) {
            console.warn('kv lock unavailable (continuing without lock)', e.message);
            return noLock;
        }
        if (acquired) {
            return {
                release: async () => {
                    try {
                        if (backend === 'supabase') await supabaseUnlock(supabaseConfig(), lockKey);
                        else if (backend === 'upstash') await kvCommand(['DEL', lockKey]);
                        else devLocks.delete(lockKey);
                    } catch (_) { /* TTLで自然解放 */ }
                },
            };
        }
        await new Promise(r => setTimeout(r, waitMs + Math.floor(Math.random() * 60)));
    }
    // 待ちきれなかった場合もフェイルオープン（TTL内に前の保存は完了している想定）
    console.warn('kv lock timeout (continuing without lock)', lockKey);
    return noLock;
}

// ロック付きの read-modify-write ヘルパー
//   mutate(current) は新しい値を返す（nullを返すと書き込みしない）
async function kvUpdate(key, mutate, options) {
    const lock = await kvLock(key, options);
    try {
        const current = await kvGet(key);
        const next = await mutate(current);
        if (next !== null && next !== undefined) await kvSet(key, next);
        return next;
    } finally {
        await lock.release();
    }
}

module.exports = { kvAvailable, kvBackend, kvStatus, kvGet, kvSet, kvCommand, kvLock, kvUpdate };
