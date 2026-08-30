// アプリ全体の状態ストア（単一ソース + 購読モデル）

import { todayJst } from './format.js';

const t = todayJst();

export const state = {
    session: null,          // {role, shopId, shopName, staffId, staffName}
    demo: false,
    aiAvailable: false,
    masters: { shops: [], staffs: [], visitSources: [] },
    filters: {
        shopId: 'all',      // 'all' | number
        staffId: 'all',     // 'all' | number
        periodKind: 'month',
        anchor: { y: t.y, m: t.m },
    },
    data: {
        summary: null,       // 現在期間の売上サマリ
        summaryPrev: null,   // 前期間
        summaryYoy: null,    // 前年同期間（単月時のみ）
        today: null,         // 本日分
        channels: null,      // marketing/by-channel
        mkStaff: null,       // marketing/by-staff
        mkStaffPrev: null,
        retention: null,
        ageDist: null,
    },
    ui: {
        activeTab: 'overview',
        lastFetchAt: null,
        loadError: null,
    },
};

const listeners = new Map(); // event -> Set<fn>

export function on(event, fn) {
    if (!listeners.has(event)) listeners.set(event, new Set());
    listeners.get(event).add(fn);
    return () => listeners.get(event).delete(fn);
}

export function emit(event, payload) {
    for (const fn of listeners.get(event) || []) {
        try { fn(payload); } catch (e) { console.error(`listener error [${event}]`, e); }
    }
}

// ---- 役割ヘルパー ----
// 4段階: admin(オーナー) / manager(マネージャー) / store(店長) / staff(スタッフ)
export function isAdmin() {
    return state.session?.role === 'admin';
}
export function isManager() {
    return state.session?.role === 'manager';
}
// 全店舗を横断して見られる役割
export function isAdminLike() {
    return isAdmin() || isManager();
}
export function isStoreLocked() {
    return state.session?.role === 'store';
}
export function isStaffLocked() {
    return state.session?.role === 'staff';
}
export function isLocked() {
    return isStoreLocked() || isStaffLocked();
}

// 現在の対象店舗ID（'all' か 数値）
export function currentShopId() {
    if (isLocked()) return state.session.shopId;
    return state.filters.shopId;
}

// 現在の対象スタッフID（'all' か 数値）
export function currentStaffId() {
    if (isStaffLocked()) return state.session.staffId;
    return state.filters.staffId;
}

export function shopName(id) {
    if (id === 'all' || id === null || id === undefined) return '全店舗';
    const s = state.masters.shops.find(x => String(x.id) === String(id));
    return s ? s.name : `店舗${id}`;
}

export function staffName(id) {
    if (id === 'all' || id === null || id === undefined) return '店舗合計';
    const s = state.masters.staffs.find(x => String(x.id) === String(id));
    return s ? s.name : `スタッフ${id}`;
}

export function staffsOfShop(shopId) {
    if (shopId === 'all') return state.masters.staffs;
    return state.masters.staffs.filter(s => String(s.shop_id) === String(shopId));
}
