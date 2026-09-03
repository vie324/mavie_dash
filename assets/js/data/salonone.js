// SalonOneデータの取得と派生指標の計算
// 数値の定義（API仕様書より）:
//   gross_sales   = 粗売上（購入時点で計上）
//   consumed_sales = 消化売上（施術提供時点。失効益含む）
//   digest_sales  = 会計済み売上
// 売上の合計は必ず /sales/summary の値を使う（明細からの自前合計は数字がずれるため）。

import { apiGet, apiGetCached } from '../core/api.js';
import { state, emit, currentShopId, currentStaffId, isStaffLocked } from '../core/state.js';
import { periodRange, prevRange, yoyRange, todayJst, todayStr, ymd, daysInMonth, dowIndex } from '../core/format.js';

// ---- マスタ ----
export async function loadMasters() {
    const [shops, staffs, sources] = await Promise.all([
        apiGetCached('shops', {}, 3600000),
        apiGetCached('staffs', {}, 3600000),
        apiGetCached('visit-sources', {}, 3600000).catch(() => ({ data: [] })),
    ]);
    state.masters.shops = shops.data || [];
    state.masters.staffs = staffs.data || [];
    state.masters.visitSources = sources.data || [];
    emit('masters');
}

function shopParam() {
    const id = currentShopId();
    return id === 'all' ? {} : { shop_id: id };
}

export function currentRange() {
    return periodRange(state.filters.periodKind, state.filters.anchor);
}

// ---- 売上サマリ（現在・前期・前年・当月）----
export async function loadCore() {
    const range = currentRange();
    const prev = prevRange(range);
    const isMonth = state.filters.periodKind === 'month';
    const t = todayJst();
    const nowMonthRange = { from: ymd(t.y, t.m, 1), to: ymd(t.y, t.m, daysInMonth(t.y, t.m)) };
    const anchorIsNow = state.filters.anchor.y === t.y && state.filters.anchor.m === t.m;

    const tasks = [
        apiGet('sales/summary', { from: range.from, to: range.to, ...shopParam() }),
        apiGet('sales/summary', { from: prev.from, to: prev.to, ...shopParam() }).catch(() => null),
        isMonth
            ? apiGet('sales/summary', { from: yoyRange(range).from, to: yoyRange(range).to, ...shopParam() }).catch(() => null)
            : Promise.resolve(null),
        (isMonth && anchorIsNow)
            ? null // 現在期間と同じなのでcurを流用
            : apiGet('sales/summary', { from: nowMonthRange.from, to: nowMonthRange.to, ...shopParam() }).catch(() => null),
    ];
    const [cur, prevSum, yoySum, nowMonth] = await Promise.all(tasks);
    state.data.summary = cur;
    state.data.summaryPrev = prevSum;
    state.data.summaryYoy = yoySum;
    state.data.nowMonth = nowMonth || cur;
    state.ui.lastFetchAt = Date.now();
    emit('data:core');
}

// ---- マーケ集計（遅延ロード）----
export async function loadMarketing() {
    const range = currentRange();
    const prev = prevRange(range);
    // staffロールには広告指標(by-channel)をサーバーが返さないため取得しない
    const [channels, mkStaff, mkStaffPrev] = await Promise.all([
        isStaffLocked() ? Promise.resolve(null) : apiGet('marketing/by-channel', { from: range.from, to: range.to, ...shopParam() }),
        apiGet('marketing/by-staff', { from: range.from, to: range.to, ...shopParam() }),
        apiGet('marketing/by-staff', { from: prev.from, to: prev.to, ...shopParam() }).catch(() => null),
    ]);
    state.data.channels = channels ? (channels.data || []) : [];
    state.data.mkStaff = mkStaff.data || [];
    state.data.mkStaffPrev = mkStaffPrev?.data || null;
    emit('data:marketing');
}

export async function loadRetention() {
    const range = currentRange();
    state.data.retention = await apiGet('marketing/retention', { from: range.from, to: range.to, ...shopParam() });
    emit('data:retention');
}

export async function loadAgeDist() {
    state.data.ageDist = await apiGetCached('insights/age-distribution', shopParam(), 600000);
    emit('data:agedist');
}

// ---- 売上の基準 ----
// 実APIでは日別・スタッフ別・支払い内訳が「会計済み売上(digest)」基準で整合しており、
// 粗売上(gross)は月合計にしか存在しない（回数券等の購入時計上分が含まれる）。
// 画面内で数字がズレないよう、既定は会計済み売上で統一する（設定タブで変更可）。
const BASIS_KEY = 'vie_sales_basis';
export const SALES_BASES = {
    digest: { field: 'digest_sales', label: '会計済み売上', short: '会計済み' },
    gross: { field: 'gross_sales', label: '粗売上', short: '粗' },
    consumed: { field: 'consumed_sales', label: '消化売上', short: '消化' },
};

export function salesBasis() {
    try {
        const v = localStorage.getItem(BASIS_KEY);
        return SALES_BASES[v] ? v : 'digest';
    } catch (_) {
        return 'digest';
    }
}

export function setSalesBasis(basis) {
    if (!SALES_BASES[basis]) return;
    try { localStorage.setItem(BASIS_KEY, basis); } catch (_) { /* ignore */ }
}

export function salesBasisLabel() {
    return SALES_BASES[salesBasis()].label;
}

// 行（サマリ/日別/スタッフ別/月次バケット）の「売上」を現在の基準で返す
export function salesOf(row) {
    if (!row) return 0;
    const v = row[SALES_BASES[salesBasis()].field];
    if (v !== undefined && v !== null) return v;
    return row.gross_sales || 0;
}

// ---- 派生指標 ----

// サマリ or by_staff/by_day 行から共通KPIを計算
export function kpisOf(row) {
    if (!row) return null;
    const visits = (row.new_visit_count || 0) + (row.repeat_visit_count || 0);
    const cancels = (row.cancel_count || 0) + (row.no_show_count || 0);
    const sales = salesOf(row);
    return {
        sales,
        gross: row.gross_sales || 0,
        consumed: row.consumed_sales || 0,
        digest: row.digest_sales || 0,
        visits,
        newVisits: row.new_visit_count || 0,
        repeatVisits: row.repeat_visit_count || 0,
        cancels,
        noShows: row.no_show_count || 0,
        unitPrice: visits > 0 ? Math.round(sales / visits) : 0,
        newRate: visits > 0 ? (row.new_visit_count || 0) / visits * 100 : 0,
        cancelRate: (visits + cancels) > 0 ? cancels / (visits + cancels) * 100 : 0,
    };
}

// 選択中スタッフのby_staff行（'all'ならサマリ全体）
export function scopedRow(summary) {
    if (!summary) return null;
    const staffId = currentStaffId();
    if (staffId === 'all') return summary;
    return (summary.by_staff || []).find(s => String(s.staff_id) === String(staffId)) || {
        gross_sales: 0, consumed_sales: 0, digest_sales: 0,
        new_visit_count: 0, repeat_visit_count: 0, cancel_count: 0, no_show_count: 0,
    };
}

// by_day を月単位に集計（3ヶ月/6ヶ月/1年表示用）
export function monthlyBuckets(byDay) {
    const buckets = new Map();
    for (const d of byDay || []) {
        const key = d.date.slice(0, 7);
        if (!buckets.has(key)) {
            buckets.set(key, { month: key, gross_sales: 0, consumed_sales: 0, digest_sales: 0, new_visit_count: 0, repeat_visit_count: 0, cancel_count: 0, no_show_count: 0 });
        }
        const b = buckets.get(key);
        for (const k of ['gross_sales', 'consumed_sales', 'digest_sales', 'new_visit_count', 'repeat_visit_count', 'cancel_count', 'no_show_count']) {
            b[k] += d[k] || 0;
        }
    }
    return [...buckets.values()].sort((a, b) => a.month.localeCompare(b.month));
}

// 月初〜今日までの実績（当月サマリのby_dayから）
export function monthToDate(nowMonthSummary) {
    const today = todayStr();
    const days = (nowMonthSummary?.by_day || []).filter(d => d.date <= today);
    const agg = { sales: 0, gross_sales: 0, digest_sales: 0, consumed_sales: 0, new_visit_count: 0, repeat_visit_count: 0, cancel_count: 0, no_show_count: 0 };
    for (const d of days) {
        agg.sales += salesOf(d);
        agg.gross_sales += d.gross_sales || 0;
        agg.digest_sales += d.digest_sales || 0;
        agg.consumed_sales += d.consumed_sales || 0;
        agg.new_visit_count += d.new_visit_count || 0;
        agg.repeat_visit_count += d.repeat_visit_count || 0;
        agg.cancel_count += d.cancel_count || 0;
        agg.no_show_count += d.no_show_count || 0;
    }
    return agg;
}

// 着地予測（曜日係数ベース + レンジ）
export function forecastMonth(nowMonthSummary) {
    const t = todayJst();
    const today = todayStr();
    const byDay = (nowMonthSummary?.by_day || []);
    const observed = byDay.filter(d => d.date <= today);
    if (observed.length < 3) return null;

    const mtd = observed.reduce((a, d) => a + salesOf(d), 0);
    // 曜日ごとの平均売上
    const dowSum = Array(7).fill(0), dowCnt = Array(7).fill(0);
    for (const d of observed) {
        const w = dowIndex(d.date);
        dowSum[w] += salesOf(d);
        dowCnt[w]++;
    }
    const overallAvg = mtd / observed.length;
    const dim = daysInMonth(t.y, t.m);
    let future = 0;
    for (let day = t.d + 1; day <= dim; day++) {
        const w = dowIndex(ymd(t.y, t.m, day));
        future += dowCnt[w] > 0 ? dowSum[w] / dowCnt[w] : overallAvg;
    }
    const forecast = Math.round(mtd + future);
    return {
        mtd,
        forecast,
        low: Math.round(mtd + future * 0.88),
        high: Math.round(mtd + future * 1.12),
        elapsedRatio: t.d / dim,
    };
}

// 曜日別平均（曜日分析チャート用） → [{dow, avgSales, avgVisits}]
export function dowAnalysis(byDay) {
    const today = todayStr();
    const sum = Array(7).fill(0).map(() => ({ sales: 0, visits: 0, n: 0 }));
    for (const d of byDay || []) {
        if (d.date > today) continue;
        const w = dowIndex(d.date);
        sum[w].sales += salesOf(d);
        sum[w].visits += (d.new_visit_count || 0) + (d.repeat_visit_count || 0);
        sum[w].n++;
    }
    return sum.map((s, w) => ({
        dow: w,
        avgSales: s.n ? Math.round(s.sales / s.n) : 0,
        avgVisits: s.n ? +(s.visits / s.n).toFixed(1) : 0,
    }));
}

// CSVダウンロード
export function downloadCsv(filename, headers, rows) {
    const bom = '﻿';
    const lines = [headers.join(',')];
    for (const row of rows) {
        lines.push(row.map(v => {
            const s = String(v ?? '');
            return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        }).join(','));
    }
    const blob = new Blob([bom + lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}
