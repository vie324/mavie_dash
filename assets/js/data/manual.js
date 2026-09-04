// 手入力データ（SalonOne APIにない項目）の取得・保存
// サーバー保存（Supabase / Upstash）が未設定の場合は、この端末のlocalStorageに退避する。
// 形式: {
//   daily:   {"YYYY-MM-DD:staffId": {nextNew, nextRepeat, blog, sns, reviews, at}},  at=保存時刻(UNIX秒・サーバー付与)
//   monthly: {"staffId": {productSales}},
//   adCosts: {"sourceId": 金額},
//   recon:   {"YYYY-MM-DD:shopId": {"m<支払方法ID>": 実際額, memo}}
// }

import { state, emit } from '../core/state.js';
import { ApiError } from '../core/api.js';

export const BLOG_TARGET = 10; // 月間ブログ更新目標（従来ツールの値を踏襲）
export const DAILY_FIELDS = ['nextNew', 'nextRepeat', 'blog', 'sns', 'reviews'];

const LOCAL_PREFIX = 'vie_manual_';

function emptyData() {
    return { daily: {}, monthly: {}, adCosts: {}, recon: {} };
}

export function monthOf(date) {
    return String(date).slice(0, 7);
}

export function monthKeyOf(y, m) {
    return `${y}-${String(m).padStart(2, '0')}`;
}

function localLoad(month) {
    try {
        return { ...emptyData(), ...(JSON.parse(localStorage.getItem(LOCAL_PREFIX + month) || '{}')) };
    } catch (_) {
        return emptyData();
    }
}

function localSave(month, data) {
    try { localStorage.setItem(LOCAL_PREFIX + month, JSON.stringify(data)); } catch (_) { /* ignore */ }
}

// ローカルへのパッチ適用（サーバー未設定時のフォールバック。サーバーと同じ規則で保存時刻を付ける）
function localApplyPatch(month, patch) {
    const data = localLoad(month);
    for (const section of ['daily', 'monthly', 'recon']) {
        for (const [key, entry] of Object.entries(patch[section] || {})) {
            if (entry === null) { delete data[section][key]; continue; }
            const cur = data[section][key] || {};
            for (const [f, v] of Object.entries(entry)) {
                if (f === 'at') continue;
                if (v === null || v === '') delete cur[f];
                else cur[f] = f === 'memo' ? String(v) : Number(v);
            }
            const keys = Object.keys(cur).filter(k => k !== 'at');
            if (keys.length === 0) delete data[section][key];
            else {
                if (section === 'daily') cur.at = Math.floor(Date.now() / 1000);
                data[section][key] = cur;
            }
        }
    }
    for (const [k, v] of Object.entries(patch.adCosts || {})) {
        if (v === null) delete data.adCosts[k];
        else data.adCosts[k] = Number(v);
    }
    localSave(month, data);
    return data;
}

async function request(path, options) {
    const res = await fetch(path, { credentials: 'same-origin', cache: 'no-store', ...options });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空 */ }
    if (!res.ok) throw new ApiError(res.status, json.error || 'unknown', json);
    return json;
}

function fromResponse(res) {
    return { daily: res.daily || {}, monthly: res.monthly || {}, adCosts: res.adCosts || {}, recon: res.recon || {} };
}

// 同じ月の同時読み込みは1リクエストにまとめる
const inflight = new Map();

// month: 'YYYY-MM'
export function loadManual(month) {
    if (!state.data.manual) state.data.manual = {};
    if (inflight.has(month)) return inflight.get(month);
    const p = (async () => {
        try {
            const res = await request(`/api/manual?month=${month}`);
            if (res.storage === 'none') {
                // サーバー保存なし → ローカル退避分を使う
                state.manualStorage = 'local';
                state.data.manual[month] = localLoad(month);
            } else {
                state.manualStorage = 'kv';
                state.data.manual[month] = fromResponse(res);
            }
        } catch (e) {
            console.warn('manual load', e);
            state.manualStorage = state.manualStorage || 'local';
            state.data.manual[month] = localLoad(month);
        } finally {
            inflight.delete(month);
        }
        emit('data:manual');
        return state.data.manual[month];
    })();
    inflight.set(month, p);
    return p;
}

export async function saveManualPatch(month, patch) {
    if (!state.data.manual) state.data.manual = {};
    if (state.manualStorage === 'local') {
        state.data.manual[month] = localApplyPatch(month, patch);
        emit('data:manual');
        return { storage: 'local' };
    }
    try {
        const res = await request(`/api/manual?month=${month}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ patch }),
        });
        state.data.manual[month] = fromResponse(res);
        emit('data:manual');
        return { storage: 'kv' };
    } catch (e) {
        if (e instanceof ApiError && (e.status === 501 || e.code === 'storage_unconfigured')) {
            state.manualStorage = 'local';
            state.data.manual[month] = localApplyPatch(month, patch);
            emit('data:manual');
            return { storage: 'local' };
        }
        throw e;
    }
}

export function getManual(month) {
    return state.data.manual?.[month] || emptyData();
}

export function isManualLoaded(month) {
    return !!state.data.manual?.[month];
}

// 日次エントリ（無ければnull）
export function getDailyEntry(date, staffId) {
    return getManual(monthOf(date)).daily[`${date}:${staffId}`] || null;
}

// エントリに何か値が入っているか（保存時刻だけの空エントリは無視）
export function hasValues(entry) {
    return !!entry && DAILY_FIELDS.some(f => entry[f] !== undefined && entry[f] !== null);
}

// 月内で入力のある日付 → Set<'YYYY-MM-DD'>
export function daysWithEntry(month, staffId) {
    const out = new Set();
    for (const [key, entry] of Object.entries(getManual(month).daily)) {
        const [date, sid] = key.split(':');
        if (String(sid) === String(staffId) && hasValues(entry)) out.add(date);
    }
    return out;
}

// その日に入力済みのスタッフID → Set<string>
export function staffWithEntryOn(date) {
    const out = new Set();
    for (const [key, entry] of Object.entries(getManual(monthOf(date)).daily)) {
        const [d, sid] = key.split(':');
        if (d === date && hasValues(entry)) out.add(String(sid));
    }
    return out;
}

// 月内のスタッフ別 日次合計 → {"staffId": {blog, sns, reviews, nextNew, nextRepeat, days}}
export function monthlyTotalsByStaff(month) {
    const data = getManual(month);
    const totals = {};
    for (const [key, entry] of Object.entries(data.daily)) {
        const staffId = key.split(':')[1];
        if (!totals[staffId]) totals[staffId] = { blog: 0, sns: 0, reviews: 0, nextNew: 0, nextRepeat: 0, days: 0 };
        totals[staffId].blog += entry.blog || 0;
        totals[staffId].sns += entry.sns || 0;
        totals[staffId].reviews += entry.reviews || 0;
        totals[staffId].nextNew += entry.nextNew || 0;
        totals[staffId].nextRepeat += entry.nextRepeat || 0;
        if (hasValues(entry)) totals[staffId].days++;
    }
    return totals;
}

export function emptyTotals() {
    return { blog: 0, sns: 0, reviews: 0, nextNew: 0, nextRepeat: 0, days: 0 };
}
