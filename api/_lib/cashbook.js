// 現金出納帳の計算ロジック（サーバー側）
//
// 帳簿残高の考え方:
//   その日の帳簿残高 = 前日繰越 + SalonOneの現金売上 + 手入力の入金 − 手入力の出金
//   締めた日は「実査額（数えた現金）」が翌日の繰越になる（過不足はその日で確定）
//   締めていない日は帳簿残高がそのまま繰り越される
// 前日繰越は「直前に締めた日の実査額」から毎回計算し直す（保存値に頼らないので、途中の修正が自動で反映される）。
// 締めた日が一度もない場合は、設定の開始日・開始時の残高から計算する。
//
// 保存構造（api/_lib/kv.js のキー）:
//   vie:cashcfg                      → { categories, cashMethodIds, shops: { "<shopId>": { startDate, initialBalance, float } }, log }
//   vie:cash:<shopId>:<YYYY-MM>      → { entries: [...], days: { "YYYY-MM-DD": { count, closed, history } }, log: [...] }

'use strict';

const crypto = require('crypto');
const { kvGet } = require('./kv');
const { fetchSalonOne } = require('./salonone');

const MONTH_RE = /^\d{4}-\d{2}$/;
const DAY_RE = /^\d{4}-\d{2}-\d{2}$/;
const CFG_KEY = 'vie:cashcfg';
const DENOMS = [10000, 5000, 2000, 1000, 500, 100, 50, 10, 5, 1];
const MAX_AMOUNT = 10000000;       // 1件・残高の上限（1,000万円）
const LOOKBACK_MONTHS = 24;        // 繰越の起点（直近の締め）を探す最大月数
const CFG_LOG_LIMIT = 500;

// 既定の科目（名前の変更・非表示・追加は設定から）
const DEFAULT_CATEGORIES = [
    { id: 'in_change', type: 'in', name: '釣り銭の補充（両替・本部から）' },
    { id: 'in_prepaid', type: 'in', name: '回数券・サブスク等の販売（現金）' },
    { id: 'in_product', type: 'in', name: '物販（SalonOne外）' },
    { id: 'in_other', type: 'in', name: 'その他の入金' },
    { id: 'out_deposit', type: 'out', name: '銀行へ預け入れ' },
    { id: 'out_refund', type: 'out', name: '返金（現金）' },
    { id: 'out_supplies', type: 'out', name: '消耗品・備品' },
    { id: 'out_transport', type: 'out', name: '交通費' },
    { id: 'out_postage', type: 'out', name: '通信・郵送' },
    { id: 'out_meeting', type: 'out', name: '会議・接待' },
    { id: 'out_hq', type: 'out', name: '本部・オーナーへ渡した' },
    { id: 'out_other', type: 'out', name: 'その他の出金' },
];

// ---- 日付 ----
function todayJst() {
    return new Date(Date.now() + 9 * 3600e3).toISOString().slice(0, 10);
}
function addDays(date, n) {
    const d = new Date(`${date}T00:00:00Z`);
    d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
}
function monthOf(date) {
    return String(date).slice(0, 7);
}
function prevMonth(month) {
    const [y, m] = month.split('-').map(Number);
    return m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, '0')}`;
}
function monthEnd(month) {
    const [y, m] = month.split('-').map(Number);
    return `${month}-${String(new Date(Date.UTC(y, m, 0)).getUTCDate()).padStart(2, '0')}`;
}

function newId(prefix) {
    return prefix + crypto.randomBytes(6).toString('hex');
}

function docKey(shopId, month) {
    return `vie:cash:${shopId}:${month}`;
}

function normDoc(raw) {
    const d = raw && typeof raw === 'object' ? raw : {};
    return {
        entries: Array.isArray(d.entries) ? d.entries : [],
        days: d.days && typeof d.days === 'object' ? d.days : {},
        log: Array.isArray(d.log) ? d.log : [],
    };
}

function intIn(v, min, max) {
    const n = Number(v);
    if (!Number.isFinite(n) || !Number.isInteger(n) || n < min || n > max) return null;
    return n;
}

function cleanText(v, max) {
    if (v === undefined || v === null) return '';
    if (typeof v !== 'string') return null;
    const s = v.replace(/[\u0000-\u001f\u007f]/g, ' ').trim();
    return s.length > max ? null : s;
}

// ---- 設定 ----
function sanitizeCategories(raw) {
    const out = [];
    const seen = new Set();
    for (const c of Array.isArray(raw) ? raw : []) {
        if (!c || typeof c !== 'object') continue;
        const id = String(c.id || '');
        const def = DEFAULT_CATEGORIES.find(d => d.id === id);
        if ((!def && !/^c_[0-9a-f]{12}$/.test(id)) || seen.has(id)) continue;
        const type = def ? def.type : (c.type === 'in' || c.type === 'out' ? c.type : null);
        if (!type) continue;
        const name = cleanText(c.name, 30);
        seen.add(id);
        out.push({ id, type, name: name || (def ? def.name : '（名称なし）'), disabled: !!c.disabled });
    }
    // 既定の科目は必ず残す（後から既定に追加した科目も自動で入る）
    for (const d of DEFAULT_CATEGORIES) {
        if (!seen.has(d.id)) out.push({ ...d, disabled: false });
    }
    return out.slice(0, 60);
}

function sanitizeShopSettings(raw) {
    if (!raw || typeof raw !== 'object') return null;
    const startDate = DAY_RE.test(String(raw.startDate || '')) ? raw.startDate : null;
    if (!startDate) return null;
    return {
        startDate,
        initialBalance: intIn(raw.initialBalance, 0, MAX_AMOUNT) ?? 0,
        float: intIn(raw.float, 0, MAX_AMOUNT) ?? 0,
        updatedAt: raw.updatedAt || null,
        updatedBy: raw.updatedBy || null,
    };
}

function sanitizeCfg(raw) {
    const r = raw && typeof raw === 'object' ? raw : {};
    const shops = {};
    for (const [id, s] of Object.entries(r.shops || {})) {
        if (!/^\d+$/.test(id)) continue;
        const v = sanitizeShopSettings(s);
        if (v) shops[id] = v;
    }
    return {
        categories: sanitizeCategories(r.categories),
        cashMethodIds: Array.isArray(r.cashMethodIds)
            ? [...new Set(r.cashMethodIds.map(Number).filter(n => Number.isInteger(n) && n > 0))].slice(0, 20)
            : [],
        shops,
        log: Array.isArray(r.log) ? r.log.slice(-CFG_LOG_LIMIT) : [],
    };
}

async function loadCfg() {
    return sanitizeCfg(await kvGet(CFG_KEY));
}

// ---- 権限・記入者 ----
function isAdminLike(session) {
    return session.role === 'admin' || session.role === 'manager';
}

function canAccessShop(session, shopId) {
    if (isAdminLike(session)) return true;
    return session.role === 'store' && String(session.shopId) === String(shopId);
}

function actorOf(session) {
    if (session.role === 'admin') return 'オーナー';
    if (session.role === 'manager') return 'マネージャー';
    if (session.role === 'store') return `店長（${session.shopName || session.shopId}）`;
    return session.staffName || 'スタッフ';
}

// ---- SalonOne の現金 ----
function isCashMethod(p, cfg) {
    const id = p.payment_method_id;
    if (id === null || id === undefined) return false;
    if (cfg.cashMethodIds.length) return cfg.cashMethodIds.includes(Number(id));
    return /現金|cash|キャッシュ/i.test(String(p.name || ''));
}

// 日別の現金売上（支払い方法の内訳のうち現金とみなすもの）
async function salonCash(shopId, from, to, cfg) {
    if (from > to) return { ok: true, byDay: {}, methods: [] };
    try {
        const raw = await fetchSalonOne('sales/summary', { from, to, shop_id: shopId });
        const d = raw?.data && !Array.isArray(raw.data) ? raw.data : raw;
        const byDay = {};
        const methods = new Map();
        for (const r of d?.by_day || []) {
            let cash = 0, unset = 0;
            const cashMethods = [];
            for (const p of r.payment_breakdown || []) {
                const amount = Number(p.amount) || 0;
                const id = p.payment_method_id;
                if (id === null || id === undefined) { unset += amount; continue; }
                const cashLike = isCashMethod(p, cfg);
                methods.set(String(id), { id: Number(id), name: p.name || `ID ${id}`, isCash: cashLike });
                if (cashLike) { cash += amount; cashMethods.push({ id: Number(id), name: p.name || '', amount }); }
            }
            byDay[r.date] = { cash, cashMethods, unset, refund: Number(r.refund_amount) || 0 };
        }
        return { ok: true, byDay, methods: [...methods.values()] };
    } catch (e) {
        console.warn('cashbook salon fetch failed', e.message);
        return { ok: false, byDay: {}, methods: [] };
    }
}

// ---- 帳簿の計算 ----
// ctx: { today, docs: Map<month, doc>（読み込み済みのキャッシュ。書き込み中の最新版で上書きできる）}
async function getDoc(shopId, month, ctx) {
    if (!ctx.docs.has(month)) ctx.docs.set(month, normDoc(await kvGet(docKey(shopId, month))));
    return ctx.docs.get(month);
}

function sumEntries(entries, date) {
    let inSum = 0, outSum = 0, count = 0;
    for (const e of entries) {
        if (e.date !== date || e.voided) continue;
        count++;
        if (e.type === 'in') inSum += e.amount; else outSum += e.amount;
    }
    return { inSum, outSum, count };
}

async function buildView(shopId, month, cfg, ctx) {
    const today = ctx.today;
    const sc = cfg.shops[String(shopId)];
    const base = {
        shopId: Number(shopId), month, today,
        settings: { shop: sc || null, categories: cfg.categories, cashMethodIds: cfg.cashMethodIds, denoms: DENOMS },
    };
    const doc = await getDoc(shopId, month, ctx);
    base.entries = [...doc.entries].sort((a, b) => (a.date === b.date ? String(a.createdAt).localeCompare(String(b.createdAt)) : a.date.localeCompare(b.date)));
    base.logCount = doc.log.length;
    if (!sc) return { ...base, setupRequired: true, rows: [], summary: null, methods: [] };

    const mStart = `${month}-01`;
    const mEnd = monthEnd(month);
    if (mEnd < sc.startDate) return { ...base, beforeStart: true, rows: [], summary: null, methods: [] };
    const rangeStart = sc.startDate > mStart ? sc.startDate : mStart;

    // 繰越の起点: 対象月より前で最後に締めた日（なければ開始日）
    let anchorDate = null, anchorBal = null, openingUncertain = false;
    if (rangeStart === sc.startDate) {
        anchorDate = addDays(sc.startDate, -1);
        anchorBal = sc.initialBalance;
    } else {
        const startMonth = monthOf(sc.startDate);
        let m = prevMonth(month);
        let walked = 0;
        while (m >= startMonth && walked < LOOKBACK_MONTHS) {
            const d = await getDoc(shopId, m, ctx);
            const closed = Object.keys(d.days).filter(x => d.days[x]?.closed && x >= sc.startDate && x < rangeStart).sort();
            if (closed.length) {
                anchorDate = closed[closed.length - 1];
                anchorBal = d.days[anchorDate].closed.counted;
                break;
            }
            m = prevMonth(m);
            walked++;
        }
        if (anchorDate === null) {
            anchorDate = addDays(sc.startDate, -1);
            anchorBal = sc.initialBalance;
            if (m >= startMonth) openingUncertain = true; // 遡りきれなかった（2年以上締めていない）
        }
    }

    const walkFrom = addDays(anchorDate, 1);
    const dataTo = mEnd < today ? mEnd : today;
    const salon = await salonCash(shopId, walkFrom, dataTo, cfg);

    const rows = [];
    let bal = anchorBal;
    for (let d = walkFrom; d <= dataTo; d = addDays(d, 1)) {
        const dDoc = await getDoc(shopId, monthOf(d), ctx);
        const rec = dDoc.days[d] || {};
        const ent = sumEntries(dDoc.entries, d);
        const s = salon.ok ? (salon.byDay[d] || { cash: 0, cashMethods: [], unset: 0, refund: 0 }) : null;
        const opening = bal;
        const expected = opening === null || !s ? null : opening + s.cash + ent.inSum - ent.outSum;
        const row = {
            date: d,
            opening,
            salonCash: s ? s.cash : null,
            cashMethods: s ? s.cashMethods : [],
            unset: s ? s.unset : 0,
            refund: s ? s.refund : 0,
            manualIn: ent.inSum,
            manualOut: ent.outSum,
            entryCount: ent.count,
            expected,
            count: rec.count || null,
            counted: null,
            diff: null,
            closed: rec.closed || null,
            reopenCount: Array.isArray(rec.history) ? rec.history.length : 0,
            salonChanged: false,
            openingChanged: false,
        };
        if (rec.closed) {
            row.counted = rec.closed.counted;
            row.diff = rec.closed.diff;
            row.salonChanged = !!s && s.cash !== rec.closed.salonCash;
            row.openingChanged = opening !== null && opening !== rec.closed.opening;
            row.currentDiff = expected === null ? null : rec.closed.counted - expected;
            row.status = rec.closed.diff === 0 ? 'closed' : 'closed_diff';
            bal = rec.closed.counted;
        } else {
            if (rec.count) {
                row.counted = rec.count.total;
                row.diff = expected === null ? null : rec.count.total - expected;
                row.status = 'counted';
            } else {
                row.status = (s && s.cash) || ent.count ? 'open' : 'idle';
            }
            bal = expected;
        }
        if (d >= rangeStart) rows.push(row);
    }

    // 月の集計
    const byCat = {};
    for (const e of doc.entries) {
        if (e.voided || e.date < rangeStart || e.date > dataTo) continue;
        byCat[e.cat] = (byCat[e.cat] || 0) + e.amount;
    }
    const last = rows[rows.length - 1];
    const summary = {
        opening: rows.length ? rows[0].opening : (dataTo < rangeStart ? null : anchorBal),
        closing: last ? (last.closed ? last.closed.counted : last.expected) : null,
        salonCash: salon.ok ? rows.reduce((a, r) => a + (r.salonCash || 0), 0) : null,
        manualIn: rows.reduce((a, r) => a + r.manualIn, 0),
        manualOut: rows.reduce((a, r) => a + r.manualOut, 0),
        byCat,
        diffTotal: rows.reduce((a, r) => a + (r.closed ? r.closed.diff : 0), 0),
        diffDays: rows.filter(r => r.closed && r.closed.diff !== 0).length,
        closedDays: rows.filter(r => r.closed).length,
        pendingDays: rows.filter(r => !r.closed && r.date < today && (r.status === 'open' || r.status === 'counted')).map(r => r.date),
        changedDays: rows.filter(r => r.salonChanged || r.openingChanged).map(r => r.date),
    };

    return {
        ...base,
        rows,
        summary,
        methods: salon.methods,
        salonError: !salon.ok,
        openingUncertain,
        anchor: { date: anchorDate, balance: anchorBal },
    };
}

module.exports = {
    MONTH_RE, DAY_RE, CFG_KEY, DENOMS, MAX_AMOUNT, CFG_LOG_LIMIT, DEFAULT_CATEGORIES,
    todayJst, addDays, monthOf, prevMonth, monthEnd, newId, docKey, normDoc, intIn, cleanText,
    sanitizeCfg, sanitizeShopSettings, sanitizeCategories, loadCfg,
    isAdminLike, canAccessShop, actorOf, isCashMethod, salonCash, getDoc, buildView,
};
