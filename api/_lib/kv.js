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

module.exports = { kvAvailable, kvGet, kvSet };
