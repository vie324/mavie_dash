// シフト希望休のデータ層と自動分配アルゴリズム

import { state, emit } from '../core/state.js';
import { ApiError } from '../core/api.js';
import { dowIndex, daysInMonth, ymd } from '../core/format.js';

async function request(path, options) {
    const res = await fetch(path, { credentials: 'same-origin', cache: 'no-store', ...options });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空 */ }
    if (!res.ok) throw new ApiError(res.status, json.error || 'unknown', json);
    return json;
}

// month: 'YYYY-MM'
export async function loadShift(month) {
    if (!state.data.shift) state.data.shift = {};
    const res = await request(`/api/shift?month=${month}`);
    state.data.shift[month] = res;
    emit('data:shift');
    return res;
}

function post(body) {
    return request('/api/shift', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
    });
}

export async function submitShiftRequest(month, days, staffId) {
    const res = await post({ action: 'request', month, days, staffId });
    mergeShops(month, res.shops);
    return res;
}

export async function saveAssignments(month, shopId, assigned) {
    const res = await post({ action: 'assign', month, shopId, assigned });
    mergeShops(month, res.shops);
    return res;
}

export async function approveShift(month, shopId, staffIds) {
    const res = await post({ action: 'approve', month, shopId, staffIds });
    mergeShops(month, res.shops);
    return res;
}

export async function saveShiftConfig(config) {
    const res = await post({ action: 'config', config });
    for (const m of Object.keys(state.data.shift || {})) {
        state.data.shift[m].config = res.config;
    }
    emit('data:shift');
    return res;
}

function mergeShops(month, shops) {
    if (!state.data.shift) state.data.shift = {};
    if (!state.data.shift[month]) state.data.shift[month] = { month, shops: {}, config: null };
    state.data.shift[month].shops = shops || {};
    emit('data:shift');
}

export function getShift(month) {
    return state.data.shift?.[month] || null;
}

// ---- 自動分配 ----
// ルール: 月offDays日休み・うち土日はweekendOffDays日・同じ日の休みはmaxSameDayOff人まで（自動割当時）
// 希望日は最大限尊重する（枠超過分だけ警告付きで外す）。不足分は人が薄い日から自動で埋める。
export function autoDistribute({ month, staffIds, requests, existingAssigned = {}, config }) {
    const { offDays, weekendOffDays, maxSameDayOff } = config;
    const [y, m] = month.split('-').map(Number);
    const dim = daysInMonth(y, m);
    const allDays = [];
    for (let d = 1; d <= dim; d++) allDays.push(ymd(y, m, d));
    const isWeekend = date => [0, 6].includes(dowIndex(date));
    const weekdays = allDays.filter(d => !isWeekend(d));
    const weekends = allDays.filter(isWeekend);

    // 既存の承認済み割当も混雑カウントに含める
    const dayCount = {};
    const bump = d => { dayCount[d] = (dayCount[d] || 0) + 1; };
    for (const days of Object.values(existingAssigned)) {
        for (const d of days || []) bump(d);
    }

    const assigned = {};
    const warnings = [];

    // 申請が早い順に処理（希望の競合時に先着を優先）
    const order = [...staffIds].sort((a, b) => {
        const ta = requests[a]?.submittedAt || '9999';
        const tb = requests[b]?.submittedAt || '9999';
        return ta < tb ? -1 : ta > tb ? 1 : String(a).localeCompare(String(b));
    });

    // 1) 希望日の受け入れ（枠内）
    for (const staffId of order) {
        const wanted = requests[staffId]?.days || [];
        const mine = [];
        let weekendUsed = 0;
        for (const d of wanted) {
            if (mine.length >= offDays) {
                warnings.push({ staffId, day: d, reason: `月${offDays}日の上限を超えるため外しました` });
                continue;
            }
            if (isWeekend(d)) {
                if (weekendUsed >= weekendOffDays) {
                    warnings.push({ staffId, day: d, reason: `土日休みは${weekendOffDays}日までのため外しました` });
                    continue;
                }
                weekendUsed++;
            }
            if ((dayCount[d] || 0) >= maxSameDayOff) {
                warnings.push({ staffId, day: d, reason: `同日${maxSameDayOff}人の上限を超えています（希望のため残しています）` });
            }
            mine.push(d);
            bump(d);
        }
        assigned[staffId] = mine;
    }

    // 候補日の選び方: ①その日に休む人が少ない → ②本人の休みが少ない週（月内で分散） → ③日付順
    const weekOf = d => Math.floor((Number(d.slice(8)) - 1 + dowIndex(ymd(y, m, 1))) / 7);
    const pick = (staffId, candidates) => {
        const mine = new Set(assigned[staffId]);
        const myWeekCount = {};
        for (const d of mine) myWeekCount[weekOf(d)] = (myWeekCount[weekOf(d)] || 0) + 1;
        const open = candidates.filter(d => !mine.has(d));
        // 上限未満の日を優先、全て埋まっていれば最も空いている日
        const under = open.filter(d => (dayCount[d] || 0) < maxSameDayOff);
        const pool = under.length ? under : open;
        pool.sort((a, b) =>
            (dayCount[a] || 0) - (dayCount[b] || 0)
            || (myWeekCount[weekOf(a)] || 0) - (myWeekCount[weekOf(b)] || 0)
            || a.localeCompare(b));
        return pool[0] || null;
    };

    // 2) 土日枠の充足 → 3) 平日で残りを充足
    for (const staffId of order) {
        const mine = assigned[staffId];
        let weekendHave = mine.filter(isWeekend).length;
        while (weekendHave < weekendOffDays && mine.length < offDays) {
            const d = pick(staffId, weekends);
            if (!d) break;
            mine.push(d); bump(d); weekendHave++;
        }
        while (mine.length < offDays) {
            const d = pick(staffId, weekdays);
            if (!d) break;
            mine.push(d); bump(d);
        }
        mine.sort();
    }

    return { assigned, warnings, dayCount };
}
