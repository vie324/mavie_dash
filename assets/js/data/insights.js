// 予約明細からの自動推定（β）のクライアント側キャッシュ
// サーバー: /api/data/insights/appointments（次回予約・会計未処理の集計のみ。明細は返らない）

import { apiGetCached } from '../core/api.js';
import { state, emit, currentShopId } from '../core/state.js';

const TTL = 10 * 60 * 1000;
const store = new Map(); // `${shop}:${from}:${to}` -> result

function keyOf(from, to) {
    return `${currentShopId()}:${from}:${to}`;
}

export async function loadInsights({ from, to, force = false }) {
    const key = keyOf(from, to);
    if (!force && store.has(key)) return store.get(key);
    const shopId = currentShopId();
    const res = await apiGetCached('insights/appointments', { from, to, ...(shopId === 'all' ? {} : { shop_id: shopId }) }, TTL);
    store.set(key, res);
    state.data.insights = res;
    emit('data:insights');
    return res;
}

// 月（'YYYY-MM'）単位で読む
export function loadInsightsForMonth(month, opts = {}) {
    const [y, m] = month.split('-').map(Number);
    const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return loadInsights({ from: `${month}-01`, to: `${month}-${String(dim).padStart(2, '0')}`, ...opts });
}

// 直近に読み込んだ結果（ホーム・日報サマリ用）。推定に使えない場合は null
export function getInsights(month) {
    let res = null;
    if (month) {
        for (const [k, v] of store) {
            if (k.startsWith(`${currentShopId()}:${month}-01:`)) { res = v; break; }
        }
    } else {
        res = state.data.insights || null;
    }
    return res;
}

export function insightsUsable(res) {
    return !!res && res.reliable === true;
}

// スタッフのその日の推定: {visits, withNext, newVisits, newWithNext}
export function staffDayEstimate(res, staffId, date) {
    if (!insightsUsable(res)) return null;
    const cell = res.byStaffDay?.[String(staffId)]?.[date];
    if (!cell) return { visits: 0, withNext: 0, newVisits: 0, newWithNext: 0 };
    return { visits: cell[0], withNext: cell[1], newVisits: cell[2], newWithNext: cell[3] };
}

// スタッフの期間合計の推定
export function staffEstimate(res, staffId) {
    if (!insightsUsable(res)) return null;
    return res.byStaff?.[String(staffId)] || { visits: 0, withNext: 0, newVisits: 0, newWithNext: 0 };
}

export const REASON_LABELS = {
    no_data: '予約データが取得できませんでした',
    no_start_field: '予約の開始日時の項目が見つかりません',
    status_unknown: '予約のステータス（来店・会計済み）を判別できません',
    no_customer_id: '顧客IDがないため、次回予約を紐付けできません',
    no_created_field: '予約の作成日時がないため、来店時に取った予約か判定できません',
    visit_count_mismatch: '予約データで会計済みと判定した件数が、SalonOneの売上サマリの来店数と合いません',
};
