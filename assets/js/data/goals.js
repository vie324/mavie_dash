// 月次目標・基本給の保存/取得
// サーバー保存（/api/goals → Upstash Redis）を正とし、未設定時はこの端末のlocalStorageに退避する。
// 形式: goals = { "2026-08": { "all": {...}, "shop:101": {...}, "staff:1001": {...} } }
// 目標項目: sales(売上), newVisits(新規来店), joins(入会数)
// 基本給: salaries = { "<staffId>": 円/月 }（オーナーのみ）

import { state, emit, staffsOfShop, isAdminLike, isStoreLocked } from '../core/state.js';
import { ApiError } from '../core/api.js';

const KEY = 'vie_goals_v3';
const SALARY_KEY = 'vie_base_salary_v1';
const MIGRATED_KEY = 'vie_goals_migrated_v1';

// メモリ上のキャッシュ（描画は同期でここを読む）
const cache = { goals: {}, salaries: {}, storage: null, loaded: false };

function localLoad(key) {
    try {
        return JSON.parse(localStorage.getItem(key) || '{}');
    } catch (_) {
        return {};
    }
}

function localSave(key, data) {
    try {
        localStorage.setItem(key, JSON.stringify(data));
    } catch (e) {
        console.warn('保存に失敗しました', e);
    }
}

async function request(path, options) {
    const res = await fetch(path, { credentials: 'same-origin', cache: 'no-store', ...options });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空 */ }
    if (!res.ok) throw new ApiError(res.status, json.error || 'unknown', json);
    return json;
}

function postPatch(patch) {
    return request('/api/goals', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ patch }),
    });
}

function applyResponse(res) {
    cache.goals = res.goals || {};
    if (res.salaries) cache.salaries = res.salaries;
    cache.storage = res.storage || 'kv';
    cache.loaded = true;
    emit('data:goals');
}

// ---- 読み込み ----
export async function loadGoals() {
    try {
        const res = await request('/api/goals');
        if (res.storage === 'none') {
            cache.storage = 'local';
            cache.goals = localLoad(KEY);
            cache.salaries = localLoad(SALARY_KEY);
            cache.loaded = true;
            emit('data:goals');
        } else {
            applyResponse(res);
            await migrateLocalToServer();
        }
    } catch (e) {
        console.warn('goals load', e);
        cache.storage = cache.storage || 'local';
        cache.goals = localLoad(KEY);
        cache.salaries = localLoad(SALARY_KEY);
        cache.loaded = true;
        emit('data:goals');
    }
    return cache.goals;
}

// 旧バージョン（端末保存）の目標を、サーバーが空のときだけ一度だけ引き継ぐ
async function migrateLocalToServer() {
    try {
        if (localStorage.getItem(MIGRATED_KEY)) return;
        if (!isAdminLike() && !isStoreLocked()) return;
        const local = localLoad(KEY);
        const localSal = localLoad(SALARY_KEY);
        const serverEmpty = Object.keys(cache.goals).length === 0;
        const patch = {};
        if (serverEmpty && Object.keys(local).length) {
            // 店長は自店舗の範囲だけ送る（サーバーで拒否されないように）
            const goals = {};
            const ownShop = state.session?.shopId;
            const ownStaff = new Set(staffsOfShop(ownShop).map(s => String(s.id)));
            for (const [month, scopes] of Object.entries(local)) {
                for (const [scope, goal] of Object.entries(scopes || {})) {
                    if (isStoreLocked()) {
                        const ok = scope === `shop:${ownShop}` || (scope.startsWith('staff:') && ownStaff.has(scope.slice(6)));
                        if (!ok) continue;
                    }
                    if (!goals[month]) goals[month] = {};
                    goals[month][scope] = goal;
                }
            }
            if (Object.keys(goals).length) patch.goals = goals;
        }
        if (state.session?.role === 'admin' && Object.keys(cache.salaries).length === 0 && Object.keys(localSal).length) {
            patch.salaries = localSal;
        }
        if (Object.keys(patch).length) {
            const res = await postPatch(patch);
            applyResponse(res);
        }
        localStorage.setItem(MIGRATED_KEY, '1');
    } catch (e) {
        console.warn('goals migrate', e);
    }
}

export function goalsStorage() {
    return cache.storage;
}

export function goalsLoaded() {
    return cache.loaded;
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
    return cache.goals[mKey]?.[sKey] || null;
}

function cleanGoal(goal) {
    const cleaned = {};
    for (const [k, v] of Object.entries(goal || {})) {
        const n = Number(v);
        if (isFinite(n) && n > 0) cleaned[k] = Math.round(n);
    }
    return cleaned;
}

// 目標の保存（スコープ単位で置き換え）。サーバー保存が無い場合はこの端末に保存。
export async function setGoal(mKey, sKey, goal) {
    const cleaned = cleanGoal(goal);
    const value = Object.keys(cleaned).length ? cleaned : null;
    if (cache.storage === 'local') {
        if (!cache.goals[mKey]) cache.goals[mKey] = {};
        if (value) cache.goals[mKey][sKey] = value; else delete cache.goals[mKey][sKey];
        if (Object.keys(cache.goals[mKey]).length === 0) delete cache.goals[mKey];
        localSave(KEY, cache.goals);
        emit('data:goals');
        return { storage: 'local' };
    }
    try {
        const res = await postPatch({ goals: { [mKey]: { [sKey]: value } } });
        applyResponse(res);
        return { storage: 'kv' };
    } catch (e) {
        if (e instanceof ApiError && (e.status === 501 || e.code === 'storage_unconfigured')) {
            cache.storage = 'local';
            return setGoal(mKey, sKey, goal);
        }
        throw e;
    }
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

// ---- 基本給（オーナーのみ・インセンティブ計算用）----
export function getSalaries() {
    return cache.salaries || {};
}

export async function setSalary(staffId, value) {
    const n = Number(value);
    const v = isFinite(n) && n > 0 ? Math.round(n) : null;
    if (cache.storage === 'local') {
        if (v) cache.salaries[String(staffId)] = v; else delete cache.salaries[String(staffId)];
        localSave(SALARY_KEY, cache.salaries);
        emit('data:goals');
        return { storage: 'local' };
    }
    try {
        const res = await postPatch({ salaries: { [String(staffId)]: v } });
        applyResponse(res);
        return { storage: 'kv' };
    } catch (e) {
        if (e instanceof ApiError && (e.status === 501 || e.code === 'storage_unconfigured')) {
            cache.storage = 'local';
            return setSalary(staffId, value);
        }
        throw e;
    }
}

// ---- エクスポート/インポート（バックアップ・端末間の持ち運び用）----
export function exportGoals() {
    return JSON.stringify(cache.goals, null, 2);
}

export async function importGoals(json) {
    const data = JSON.parse(json);
    if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('形式が正しくありません');
    for (const [month, scopes] of Object.entries(data)) {
        if (!/^\d{4}-\d{2}$/.test(month) || !scopes || typeof scopes !== 'object') throw new Error(`形式が正しくありません（${month}）`);
    }
    if (cache.storage === 'local') {
        cache.goals = data;
        localSave(KEY, data);
        emit('data:goals');
        return { storage: 'local' };
    }
    const res = await postPatch({ goals: data });
    applyResponse(res);
    return { storage: 'kv' };
}
