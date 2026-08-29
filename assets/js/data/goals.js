// 月次目標の保存・取得
// SalonOne APIは読み取り専用のため、目標はブラウザのlocalStorageに保存する。
// 形式: { "2026-08": { "all": {...}, "shop:101": {...}, "staff:1001": {...} } }
// 目標項目: sales(売上), newVisits(新規来店), joins(入会数)

import { state, staffsOfShop } from '../core/state.js';

const KEY = 'vie_goals_v3';

function loadAll() {
    try {
        return JSON.parse(localStorage.getItem(KEY) || '{}');
    } catch (_) {
        return {};
    }
}

function saveAll(data) {
    try {
        localStorage.setItem(KEY, JSON.stringify(data));
    } catch (e) {
        console.warn('目標の保存に失敗しました', e);
    }
}

export function monthKey(anchor) {
    return `${anchor.y}-${String(anchor.m).padStart(2, '0')}`;
}

export function scopeKey(shopId, staffId) {
    if (staffId && staffId !== 'all') return `staff:${staffId}`;
    if (shopId && shopId !== 'all') return `shop:${shopId}`;
    return 'all';
}

export function getGoalRaw(mKey, sKey) {
    const all = loadAll();
    return all[mKey]?.[sKey] || null;
}

export function setGoal(mKey, sKey, goal) {
    const all = loadAll();
    if (!all[mKey]) all[mKey] = {};
    const cleaned = {};
    for (const [k, v] of Object.entries(goal || {})) {
        const n = Number(v);
        if (isFinite(n) && n > 0) cleaned[k] = n;
    }
    if (Object.keys(cleaned).length === 0) delete all[mKey][sKey];
    else all[mKey][sKey] = cleaned;
    saveAll(all);
}

// フォールバック付き取得:
//   スタッフ → 明示設定のみ
//   店舗    → 明示設定 or 所属スタッフの合計
//   全店舗  → 明示設定 or 各店舗の合計
export function getGoal(mKey, shopId, staffId) {
    if (staffId && staffId !== 'all') {
        return getGoalRaw(mKey, `staff:${staffId}`);
    }
    if (shopId && shopId !== 'all') {
        const explicit = getGoalRaw(mKey, `shop:${shopId}`);
        if (explicit) return explicit;
        return sumGoals(staffsOfShop(shopId).map(s => getGoalRaw(mKey, `staff:${s.id}`)));
    }
    const explicit = getGoalRaw(mKey, 'all');
    if (explicit) return explicit;
    return sumGoals(state.masters.shops.map(s => getGoal(mKey, s.id, 'all')));
}

function sumGoals(goals) {
    const valid = goals.filter(Boolean);
    if (valid.length === 0) return null;
    const out = {};
    for (const g of valid) {
        for (const [k, v] of Object.entries(g)) out[k] = (out[k] || 0) + v;
    }
    return out;
}

// エクスポート/インポート（端末間の持ち運び用）
export function exportGoals() {
    return JSON.stringify(loadAll(), null, 2);
}

export function importGoals(json) {
    const data = JSON.parse(json);
    if (typeof data !== 'object' || Array.isArray(data)) throw new Error('形式が正しくありません');
    saveAll(data);
}
