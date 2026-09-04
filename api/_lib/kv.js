// 手入力データの保存先: Upstash Redis (REST)
// Vercelマーケットプレイスの「Upstash for Redis」を追加すると
// KV_REST_API_URL / KV_REST_API_TOKEN（または UPSTASH_REDIS_REST_URL / TOKEN）が自動設定される。
// 未設定の場合は保存不可（クライアント側はこの端末のみのlocalStorageに退避する）。

'use strict';

function kvConfig() {
    const url = process.env.KV_REST_API_URL || process.env.UPSTASH_REDIS_REST_URL;
    const token = process.env.KV_REST_API_TOKEN || process.env.UPSTASH_REDIS_REST_TOKEN;
    if (!url || !token) return null;
    return { url: url.replace(/\/+$/, ''), token };
}

// ローカル開発用: DEV_KV_FILE にJSONファイルパスを指定するとファイル保存で代用
function devFile() {
    return process.env.DEV_KV_FILE || null;
}

function devLoad() {
    const fs = require('fs');
    try { return JSON.parse(fs.readFileSync(devFile(), 'utf8')); } catch (_) { return {}; }
}

function kvAvailable() {
    return !!kvConfig() || !!devFile();
}

async function kvGet(key) {
    const cfg = kvConfig();
    if (!cfg && devFile()) return devLoad()[key] ?? null;
    if (!cfg) throw new Error('kv_unconfigured');
    const res = await fetch(`${cfg.url}/get/${encodeURIComponent(key)}`, {
        headers: { Authorization: `Bearer ${cfg.token}` },
    });
    if (!res.ok) throw new Error(`kv get failed: ${res.status}`);
    const json = await res.json();
    if (json.result === null || json.result === undefined) return null;
    try {
        return JSON.parse(json.result);
    } catch (_) {
        return null;
    }
}

async function kvSet(key, value) {
    const cfg = kvConfig();
    if (!cfg && devFile()) {
        const fs = require('fs');
        const all = devLoad();
        all[key] = value;
        fs.writeFileSync(devFile(), JSON.stringify(all));
        return true;
    }
    if (!cfg) throw new Error('kv_unconfigured');
    const res = await fetch(`${cfg.url}/set/${encodeURIComponent(key)}`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${cfg.token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(value),
    });
    if (!res.ok) throw new Error(`kv set failed: ${res.status}`);
    return true;
}

// 任意のRedisコマンド（Upstash REST: POST / に ["CMD", ...args] を送る）
async function kvCommand(args) {
    const cfg = kvConfig();
    if (!cfg) throw new Error('kv_unconfigured');
    const res = await fetch(cfg.url, {
        method: 'POST',
        headers: { Authorization: `Bearer ${cfg.token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(args.map(a => String(a))),
    });
    if (!res.ok) throw new Error(`kv command failed: ${res.status}`);
    const json = await res.json();
    if (json.error) throw new Error(`kv command error: ${json.error}`);
    return json.result;
}

// ---- 簡易ロック（読み取り→加工→書き込みの競合防止）----
// 複数スタッフが同時に日報を保存しても後勝ちで消えないよう、キーごとに短時間のロックを取る。
// SET NX PX で取得し、取れなければ少し待って再試行。ロック基盤に問題があっても保存自体は止めない（フェイルオープン）。
const devLocks = new Map();

async function kvLock(key, { ttlMs = 5000, retries = 12, waitMs = 120 } = {}) {
    const lockKey = `lock:${key}`;
    const cfg = kvConfig();
    for (let i = 0; i <= retries; i++) {
        let acquired = false;
        if (!cfg && devFile()) {
            // 単一プロセスの開発サーバー: メモリ上で擬似ロック
            const until = devLocks.get(lockKey) || 0;
            if (until < Date.now()) { devLocks.set(lockKey, Date.now() + ttlMs); acquired = true; }
        } else if (cfg) {
            try {
                acquired = (await kvCommand(['SET', lockKey, '1', 'NX', 'PX', ttlMs])) === 'OK';
            } catch (e) {
                console.warn('kv lock unavailable (continuing without lock)', e.message);
                return { release: async () => {} };
            }
        } else {
            return { release: async () => {} };
        }
        if (acquired) {
            return {
                release: async () => {
                    if (!cfg && devFile()) { devLocks.delete(lockKey); return; }
                    try { await kvCommand(['DEL', lockKey]); } catch (_) { /* TTLで自然解放 */ }
                },
            };
        }
        await new Promise(r => setTimeout(r, waitMs + Math.floor(Math.random() * 60)));
    }
    // 待ちきれなかった場合もフェイルオープン（TTL内に前の保存は完了している想定）
    console.warn('kv lock timeout (continuing without lock)', lockKey);
    return { release: async () => {} };
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

module.exports = { kvAvailable, kvGet, kvSet, kvCommand, kvLock, kvUpdate };
