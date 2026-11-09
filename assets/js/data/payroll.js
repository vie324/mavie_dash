// 給与・インセンティブ（2026年11月分〜）の設定・勤怠の読み込み/保存と、スタッフ1人分の計算のとりまとめ
// 保存先は /api/goals?kind=payroll（サーバー保存のみ。権限: オーナー=全員分、スタッフ=本人の分の閲覧のみ）
// 給与に使う売上も同じAPIが SalonOne から全店舗の合計で返す（画面の店舗の絞り込みに左右されない・本人画面とオーナー画面で一致する）
// 計算そのものは payroll/calc.js（画面・通信と切り離した関数）、金額・率は payroll/config.js にある。

import { state, emit } from '../core/state.js';
import { ApiError } from '../core/api.js';
import { getManual } from './manual.js';
import { ruleForMonth } from '../payroll/config.js';
import { calcMonthlyPay, settingForMonth, salesForPayroll, getPattern } from '../payroll/calc.js';

// month -> { storage, hours, sales, loadedAt }。settings（「◯月から」の履歴）は月をまたぐので1つだけ持つ
const months = new Map();
let settings = {};
let lastError = null;
const inflight = new Map();

async function request(path, options) {
    const res = await fetch(path, { credentials: 'same-origin', cache: 'no-store', ...options });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空 */ }
    if (!res.ok) throw new ApiError(res.status, json.error || 'unknown', json);
    return json;
}

const endpoint = month => `/api/goals?kind=payroll&month=${month}`;

function applyResponse(month, res) {
    settings = res.settings || {};
    const prev = months.get(month);
    // 保存（POST）の応答には売上が含まれないので、読み込み時の売上を引き継ぐ
    const sales = res.sales !== undefined ? res.sales : (prev ? prev.sales : null);
    months.set(month, { storage: res.storage || 'kv', hours: res.hours || {}, sales, loadedAt: Date.now() });
    lastError = null;
}

// オーナーとスタッフ本人だけが読める（マネージャー・店長は読まない）
export function canViewPayroll() {
    const role = state.session?.role;
    return role === 'admin' || role === 'staff';
}

export function loadPayroll(month) {
    if (!canViewPayroll() || !ruleForMonth(month)) return Promise.resolve(null);
    if (inflight.has(month)) return inflight.get(month);
    const p = (async () => {
        try {
            applyResponse(month, await request(endpoint(month)));
        } catch (e) {
            lastError = e;
            if (!(e instanceof ApiError && e.status === 403)) console.warn('payroll load', e);
        } finally {
            inflight.delete(month);
        }
        emit('data:payroll');
        return months.get(month) || null;
    })();
    inflight.set(month, p);
    return p;
}

export function payrollLoaded(month) {
    return months.has(month);
}

export function payrollStorage(month) {
    return months.get(month)?.storage || null;
}

export function payrollError() {
    return lastError;
}

export function staffSettingHistory(staffId) {
    return settings[String(staffId)] || [];
}

export function staffSetting(staffId, month) {
    return settingForMonth(staffSettingHistory(staffId), month);
}

export function staffHours(staffId, month) {
    return months.get(month)?.hours?.[String(staffId)] || null;
}

/** その月の SalonOne の全店舗のスタッフ別売上（サーバーから）。取れなかったときは null */
export function payrollSalesMap(month) {
    return months.get(month)?.sales || null;
}

async function postPatch(month, patch) {
    const res = await request(endpoint(month), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ patch }),
    });
    applyResponse(month, res);
    emit('data:payroll');
    return res;
}

/** その月から適用する給与パターン・開店締めを保存（pattern: null は「この月から対象外」） */
export function saveStaffSetting(month, staffId, { pattern, storeDuty }) {
    const value = pattern ? { pattern, storeDuty } : { pattern: null };
    return postPatch(month, { settings: { [String(staffId)]: value } });
}

/** その月の勤怠（残業・法定時間外・総労働時間）を保存。空欄は消す */
export function saveStaffHours(month, staffId, hours) {
    const entry = {};
    for (const f of ['overtime', 'statutoryOvertime', 'totalHours']) {
        const v = hours[f];
        entry[f] = v === '' || v === null || v === undefined ? null : Number(v);
    }
    const empty = Object.values(entry).every(v => v === null);
    return postPatch(month, { hours: { [String(staffId)]: empty ? null : entry } });
}

/**
 * スタッフ1人・1か月分の給与を計算する（新ルールの月だけ）
 *   売上はサーバーが返した全店舗の合計を使う。取れなかったときだけ画面の売上（fallbackRow）で代用する
 * @returns {null | { rule, setting, sales, pay, salesFromScreen }}  対象外（旧ルールの月・パターン未設定）なら setting/pay が null
 */
export function computeStaffPay(month, staffId, fallbackRow) {
    const rule = ruleForMonth(month);
    if (!rule) return null;
    const setting = staffSetting(staffId, month);
    const manualRetail = getManual(month).monthly?.[String(staffId)]?.productSales || 0;
    const map = payrollSalesMap(month);
    const row = map ? (map[String(staffId)] || {}) : (fallbackRow || {});
    const salesFromScreen = !map;
    const sales = salesForPayroll(row, rule, manualRetail);
    if (!setting || !getPattern(rule, setting.pattern)) return { rule, setting: null, sales, pay: null, salesFromScreen };
    const pay = calcMonthlyPay(rule, {
        patternKey: setting.pattern,
        storeDuty: setting.storeDuty,
        serviceSales: sales.service,
        retailSales: sales.retail,
        hours: staffHours(staffId, month),
    });
    return { rule, setting, sales, pay, salesFromScreen };
}
