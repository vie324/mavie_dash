// 予約明細（/appointments）からの自動推定（β）
//
//   次回予約: その日に来店（会計済み）した顧客が、その日の終わり（日本時間）までに
//            「来店日時がそれより後」の予約を持っていたか。＝来店時に次の予約が入っていたか。
//   会計未処理: 開始日時が今日より前なのに、会計済みでもキャンセル/無断キャンセルでもない予約。
//
// 仕様書は明細の項目名・ステータス値を明記していないため、候補名を広く受け付けて判定する。
// 判定に使った項目名・ステータスの内訳は diagnostics（オーナーのみ）で返し、実データで検証できるようにする。
// 個人情報は返さない（顧客IDは集計の内部でのみ使い、応答には含めない）。
//
// 実APIで確認できたこと（2026-09）:
//   status は数値（0 / 2 / 3）。会計済みは accounting_confirmed_at、取消は cancelled_at に日時が入る。
//   明細には来店ではない行（サブスク課金のみ・プリペイド入金のみ・キャンセル料のみ・枠ブロック）が混ざる。
//   → 日時項目で判定し、来店ではない行は除外。さらに売上サマリの来店数と突き合わせて、
//     大きくずれる場合は推定に使わない（reliable: false）。

'use strict';

const { fetchAllPages, fetchSalonOne } = require('./salonone');

const CACHE_TTL_MS = 10 * 60 * 1000;
const LOOKBACK_DAYS = 45;     // 期間開始より前に作られた予約（前回来店時の次回予約など）も拾うための遡り日数
const MAX_PAGES = 25;         // 1000件 × 25ページ = 25,000件で打ち切り
const cache = new Map();      // key -> { expires, value }

const START_KEYS = ['start_at', 'starts_at', 'start_time', 'start_datetime', 'reserved_at', 'reservation_at', 'visit_at', 'scheduled_at'];
const CREATED_KEYS = ['created_at', 'booked_at', 'reserved_on'];
const STATUS_KEYS = ['status', 'state', 'appointment_status', 'reservation_status'];
const CANCEL_TIME_KEYS = ['cancelled_at', 'canceled_at'];
const DONE_TIME_KEYS = ['accounting_confirmed_at', 'salon_board_checkout_at', 'checked_out_at', 'checkout_at', 'paid_at', 'accounted_at', 'completed_at', 'settled_at', 'visited_at'];
// 来店ではない行（課金・入金・キャンセル料だけの行）
const NON_VISIT_FLAGS = ['subscription_billing_only', 'prepaid_charge_only', 'cancellation_fee_only', 'subscription_anchor_only'];
const MATCH_TOLERANCE = 0.25; // 売上サマリの来店数との許容ずれ（±25%）
const NEW_FLAG_KEYS = ['is_new_customer', 'is_first_visit', 'is_first', 'first_visit', 'is_new'];
const TYPE_KEYS = ['visit_type', 'customer_type', 'customer_kind'];

function pad(n) { return String(n).padStart(2, '0'); }

function pick(row, keys) {
    for (const k of keys) {
        const v = row[k];
        if (v !== undefined && v !== null && v !== '') return { key: k, value: v };
    }
    return { key: null, value: null };
}

// 明細の日時はUTC（仕様書§9）。タイムゾーン表記がない文字列もUTCとして解釈する
function parseUtc(v) {
    if (v === null || v === undefined || v === '') return NaN;
    const s = String(v).trim();
    const withTz = /[zZ]|[+-]\d{2}:?\d{2}$/.test(s) ? s : `${s.replace(' ', 'T')}Z`;
    return Date.parse(withTz);
}

function jstDate(ms) {
    const d = new Date(ms + 9 * 3600 * 1000);
    return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
}

function endOfJstDay(date) {
    return Date.parse(`${date}T23:59:59.999+09:00`);
}

const RE_NOSHOW = /no[\s_-]?show|absent|無断/i;
const RE_CANCEL = /cancel|取消|キャンセル/i;
const RE_DONE = /visit|done|complete|finish|check(ed)?[\s_-]?out|paid|account|settle|来店|会計|完了/i;

function truthy(v) {
    return v === true || v === 1 || v === '1' || v === 'true';
}

// 来店として数えない行: 顧客なし・枠ブロック・課金/入金/キャンセル料だけの行
function isNonVisit(row) {
    if (row.customer_id === undefined || row.customer_id === null || row.customer_id === '') return true;
    const block = row.slot_block_type;
    if (block !== undefined && block !== null && block !== '' && block !== 0 && block !== '0' && block !== false) return true;
    return NON_VISIT_FLAGS.some(k => truthy(row[k]));
}

// kind: done（来店・会計済み）/ canceled / no_show / open（予約中・未処理）
function classify(row) {
    const raw = pick(row, STATUS_KEYS).value;
    const status = raw === null ? '' : String(raw);
    const cancelType = row.cancel_type === undefined || row.cancel_type === null ? '' : String(row.cancel_type);
    if (RE_NOSHOW.test(status) || RE_NOSHOW.test(cancelType) || row.is_no_show === true || row.no_show === true) return 'no_show';
    if (pick(row, CANCEL_TIME_KEYS).value || RE_CANCEL.test(status)) return 'canceled';
    if (pick(row, DONE_TIME_KEYS).value || RE_DONE.test(status)) return 'done';
    return 'open';
}

function newFlag(row) {
    const f = pick(row, NEW_FLAG_KEYS);
    if (f.key) return f.value === true || f.value === 1 || f.value === '1' || f.value === 'true';
    const t = pick(row, TYPE_KEYS).value;
    if (t !== null) {
        const v = String(t);
        if (/new|first|新規|初回/i.test(v)) return true;
        if (/repeat|existing|return|再来|既存/i.test(v)) return false;
    }
    return null;
}

function emptyAgg() {
    return { visits: 0, withNext: 0, newVisits: 0, newWithNext: 0 };
}

function add(agg, isNew, hasNext) {
    agg.visits++;
    if (hasNext) agg.withNext++;
    if (isNew) {
        agg.newVisits++;
        if (hasNext) agg.newWithNext++;
    }
}

// 顧客の初回来店日（予約行に新規フラグが無い場合のフォールバック）
async function firstVisitDates(shopId) {
    const { data } = await fetchAllPages('customers', shopId ? { shop_id: shopId } : {}, 10);
    const map = new Map();
    for (const c of data) {
        if (c.id === undefined || c.id === null || c.deleted_at) continue;
        const ms = parseUtc(c.first_visit_at);
        if (isFinite(ms)) map.set(String(c.id), jstDate(ms));
    }
    return map;
}

/**
 * @param {{from:string,to:string,shopId?:string|number|null,today:string}} opts  from/to/today は日本時間の暦日
 */
async function appointmentInsights({ from, to, shopId = null, today }) {
    const key = `${shopId || 'all'}:${from}:${to}:${today}`;
    const hit = cache.get(key);
    if (hit && hit.expires > Date.now()) return hit.value;

    const sinceMs = Date.parse(`${from}T00:00:00+09:00`) - LOOKBACK_DAYS * 86400000;
    const params = { updated_since: new Date(sinceMs).toISOString().replace(/\.\d{3}Z$/, 'Z') };
    if (shopId) params.shop_id = shopId;
    const { data, truncated } = await fetchAllPages('appointments', params, MAX_PAGES);

    // 同じIDは updated_at の新しい方で上書き（仕様書§7）。削除済みは除外
    const byId = new Map();
    for (const r of data) {
        if (!r || r.id === undefined || r.id === null) continue;
        const k = String(r.id);
        const prev = byId.get(k);
        if (!prev || String(r.updated_at || '') >= String(prev.updated_at || '')) byId.set(k, r);
    }
    const rows = [...byId.values()].filter(r => !r.deleted_at);

    const fieldSet = new Set();
    const statusCounts = {};
    const valueCounts = {};
    let startKey = null, createdKey = null, newFlagFound = false;
    const appts = [];
    for (const r of rows) {
        if (fieldSet.size < 80) for (const k of Object.keys(r)) fieldSet.add(k);
        const st = pick(r, START_KEYS);
        const cr = pick(r, CREATED_KEYS);
        startKey = startKey || st.key;
        createdKey = createdKey || cr.key;
        const statusRaw = pick(r, STATUS_KEYS).value;
        const sk = statusRaw === null ? '(なし)' : String(statusRaw);
        statusCounts[sk] = (statusCounts[sk] || 0) + 1;
        for (const k of ['type', 'cancel_type', 'slot_block_type']) {
            if (!(k in r)) continue;
            const v = r[k] === null || r[k] === undefined || r[k] === '' ? '(なし)' : String(r[k]);
            if (!valueCounts[k]) valueCounts[k] = {};
            valueCounts[k][v] = (valueCounts[k][v] || 0) + 1;
        }
        const startMs = parseUtc(st.value);
        if (!isFinite(startMs)) continue;
        const isNew = newFlag(r);
        if (isNew !== null) newFlagFound = true;
        const cancelMs = parseUtc(pick(r, CANCEL_TIME_KEYS).value);
        appts.push({
            nonVisit: isNonVisit(r),
            visitSource: r.visit_source_id !== undefined && r.visit_source_id !== null && r.visit_source_id !== '',
            id: String(r.id),
            customer: r.customer_id === undefined || r.customer_id === null ? null : String(r.customer_id),
            staff: r.staff_id === undefined || r.staff_id === null ? null : String(r.staff_id),
            startMs,
            createdMs: parseUtc(cr.value),
            cancelMs,
            date: jstDate(startMs),
            kind: classify(r),
            isNew,
        });
    }

    const kindCounts = { done: 0, canceled: 0, no_show: 0, open: 0, non_visit: 0 };
    for (const a of appts) {
        if (a.nonVisit) kindCounts.non_visit++;
        else kindCounts[a.kind]++;
    }

    // 新規判定: 予約行のフラグ → 無ければ顧客の初回来店日
    let newSource = newFlagFound ? 'appointment' : null;
    let firstVisit = null;
    if (!newFlagFound) {
        try {
            firstVisit = await firstVisitDates(shopId);
            if (firstVisit.size > 0) newSource = 'customer_first_visit';
        } catch (_) { firstVisit = null; }
    }

    // 売上サマリ（公式の来店数）との突き合わせ: 今日までの範囲
    const toCap = to < today ? to : today;
    const doneRange = appts.filter(a => !a.nonVisit && a.kind === 'done' && a.date >= from && a.date <= toCap);
    const calibration = { summaryVisits: null, summaryNew: null, done: doneRange.length, doneWithSource: doneRange.filter(a => a.visitSource).length };
    try {
        const sum = await fetchSalonOne('sales/summary', { from, to: toCap, ...(shopId ? { shop_id: shopId } : {}) });
        const row = sum?.data && !Array.isArray(sum.data) ? sum.data : sum;
        const nv = Number(row?.new_visit_count), rv = Number(row?.repeat_visit_count);
        if (isFinite(nv) && isFinite(rv)) { calibration.summaryVisits = nv + rv; calibration.summaryNew = nv; }
    } catch (_) { /* 突き合わせできなくても推定は続ける */ }
    const near = (a, b) => b > 0 && Math.abs(a - b) / b <= MATCH_TOLERANCE;
    // 新規フラグも初回来店日も無い場合: 流入元が入っている来店を新規とみなす（新規来店数と一致する場合のみ）
    let useSourceAsNew = false;
    if (!newSource && calibration.summaryNew !== null && near(calibration.doneWithSource, calibration.summaryNew)) {
        useSourceAsNew = true;
        newSource = 'visit_source';
    }

    const byCustomer = new Map();
    for (const a of appts) {
        if (!a.customer) continue;
        if (!byCustomer.has(a.customer)) byCustomer.set(a.customer, []);
        byCustomer.get(a.customer).push(a);
    }

    const total = emptyAgg();
    const byStaff = {};
    const byStaffDay = {}; // staffId -> date -> [visits, withNext, newVisits, newWithNext]
    const byDay = {};
    for (const v of appts) {
        if (v.nonVisit || v.kind !== 'done' || v.date < from || v.date > to) continue;
        const eod = endOfJstDay(v.date);
        const hasNext = !!v.customer && (byCustomer.get(v.customer) || []).some(a =>
            !a.nonVisit
            && a.id !== v.id
            && a.date > v.date // 同じ日の別予約（メニュー追加など）は次回予約に数えない
            && a.startMs > v.startMs
            && isFinite(a.createdMs) && a.createdMs <= eod
            // 後でキャンセルされた予約も「その日の時点では入っていた」なら数える。取消日時が無ければ数えない
            && (a.kind !== 'canceled' || (isFinite(a.cancelMs) && a.cancelMs > eod)));
        let isNew = v.isNew;
        if (isNew === null && firstVisit && firstVisit.size > 0 && v.customer) isNew = firstVisit.get(v.customer) === v.date;
        if (isNew === null && useSourceAsNew) isNew = v.visitSource;
        const sid = v.staff || 'none';
        add(total, isNew, hasNext);
        if (!byStaff[sid]) byStaff[sid] = emptyAgg();
        add(byStaff[sid], isNew, hasNext);
        if (!byDay[v.date]) byDay[v.date] = emptyAgg();
        add(byDay[v.date], isNew, hasNext);
        if (!byStaffDay[sid]) byStaffDay[sid] = {};
        const cell = byStaffDay[sid][v.date] || [0, 0, 0, 0];
        cell[0]++;
        if (hasNext) cell[1]++;
        if (isNew) { cell[2]++; if (hasNext) cell[3]++; }
        byStaffDay[sid][v.date] = cell;
    }

    // 会計未処理: 期間内・今日より前・予約中のまま
    const unsettled = { count: 0, byDate: {}, byStaff: {} };
    for (const a of appts) {
        if (a.nonVisit || a.kind !== 'open' || a.date < from || a.date > to || a.date >= today) continue;
        unsettled.count++;
        unsettled.byDate[a.date] = (unsettled.byDate[a.date] || 0) + 1;
        const sid = a.staff || 'none';
        unsettled.byStaff[sid] = (unsettled.byStaff[sid] || 0) + 1;
    }

    // 判定の信頼性: 来店（会計済み）を1件も判別できない・顧客IDが無い場合は推定値を使わない
    const pastCount = appts.filter(a => a.date < today).length;
    let reliable = true, reason = null;
    if (appts.length === 0) { reliable = false; reason = 'no_data'; }
    else if (!startKey) { reliable = false; reason = 'no_start_field'; }
    else if (pastCount > 0 && kindCounts.done === 0) { reliable = false; reason = 'status_unknown'; }
    else if (!appts.some(a => a.customer)) { reliable = false; reason = 'no_customer_id'; }
    else if (!createdKey) { reliable = false; reason = 'no_created_field'; }
    else if (calibration.summaryVisits !== null && calibration.summaryVisits >= 20 && !near(calibration.done, calibration.summaryVisits)) {
        reliable = false; reason = 'visit_count_mismatch';
    }

    const value = {
        from, to,
        reliable, reason,
        newSplit: newSource !== null,
        total, byStaff, byStaffDay, byDay, unsettled,
        diagnostics: {
            fetched: rows.length,
            usable: appts.length,
            truncated: !!truncated,
            fields: [...fieldSet].sort(),
            startKey, createdKey,
            newSource,
            statusCounts: Object.fromEntries(Object.entries(statusCounts).sort((a, b) => b[1] - a[1]).slice(0, 20)),
            kindCounts,
            calibration,
            valueCounts: Object.fromEntries(Object.entries(valueCounts).map(([k, v]) => [k, Object.fromEntries(Object.entries(v).sort((a, b) => b[1] - a[1]).slice(0, 8))])),
        },
    };
    cache.set(key, { expires: Date.now() + CACHE_TTL_MS, value });
    if (cache.size > 100) {
        for (const [k, v] of cache) if (v.expires < Date.now()) cache.delete(k);
    }
    return value;
}

module.exports = { appointmentInsights, _internal: { classify, newFlag, parseUtc, jstDate, isNonVisit } };
