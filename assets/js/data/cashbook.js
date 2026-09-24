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
