// 手入力データ（SalonOne APIにない項目）の取得・保存
// サーバー保存（Upstash Redis）が未設定の場合は、この端末のlocalStorageに退避する。
// 形式: { daily: {"YYYY-MM-DD:staffId": {blog,sns,reviews}}, monthly: {"staffId": {productSales}}, adCosts: {"sourceId": 金額} }

import { state, emit } from '../core/state.js';
import { ApiError } from '../core/api.js';

const LOCAL_PREFIX = 'vie_manual_';

function emptyData() {
    return { daily: {}, monthly: {}, adCosts: {} };
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

// ローカルへのパッチ適用（サーバー未設定時のフォールバック）
function localApplyPatch(month, patch) {
    const data = localLoad(month);
    for (const section of ['daily', 'monthly']) {
        for (const [key, entry] of Object.entries(patch[section] || {})) {
            if (entry === null) { delete data[section][key]; continue; }
            const cur = data[section][key] || {};
            for (const [f, v] of Object.entries(entry)) {
                if (v === null) delete cur[f];
                else cur[f] = Number(v);
            }
            if (Object.keys(cur).length === 0) delete data[section][key];
            else data[section][key] = cur;
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

// month: 'YYYY-MM'
export async function loadManual(month) {
    if (!state.data.manual) state.data.manual = {};
    try {
        const res = await request(`/api/manual?month=${month}`);
        if (res.storage === 'none') {
            // サーバー保存なし → ローカル退避分を使う
            state.manualStorage = 'local';
            state.data.manual[month] = localLoad(month);
        } else {
            state.manualStorage = 'kv';
            state.data.manual[month] = { daily: res.daily || {}, monthly: res.monthly || {}, adCosts: res.adCosts || {} };
        }
    } catch (e) {
        console.warn('manual load', e);
        state.manualStorage = state.manualStorage || 'local';
        state.data.manual[month] = localLoad(month);
    }
    emit('data:manual');
    return state.data.manual[month];
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
        state.data.manual[month] = { daily: res.daily || {}, monthly: res.monthly || {}, adCosts: res.adCosts || {} };
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

// 月内のスタッフ別 日次合計 → {"staffId": {blog, sns, reviews}}
export function monthlyTotalsByStaff(month) {
    const data = getManual(month);
    const totals = {};
    for (const [key, entry] of Object.entries(data.daily)) {
        const staffId = key.split(':')[1];
        if (!totals[staffId]) totals[staffId] = { blog: 0, sns: 0, reviews: 0 };
        totals[staffId].blog += entry.blog || 0;
        totals[staffId].sns += entry.sns || 0;
        totals[staffId].reviews += entry.reviews || 0;
    }
    return totals;
}
