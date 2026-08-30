// デモデータ生成器
// SALONONE_API_KEY 未設定時に、実APIと同じ形のダミーデータを返す。
// 決定的（同じパラメータ→同じ値）にするため日付文字列ベースのシード付き乱数を使う。

'use strict';

const DEMO_SHOPS = [
    { id: 101, name: '千葉店', deleted_at: null },
    { id: 102, name: '本厚木店', deleted_at: null },
    { id: 103, name: '大和店', deleted_at: null },
];

const DEMO_STAFFS = [
    { id: 1001, name: 'Kiki', shop_id: 101, deleted_at: null },
    { id: 1002, name: 'Karin', shop_id: 101, deleted_at: null },
    { id: 1003, name: 'Nanami', shop_id: 101, deleted_at: null },
    { id: 1004, name: 'Kanon', shop_id: 101, deleted_at: null },
    { id: 1005, name: 'Ayami', shop_id: 101, deleted_at: null },
    { id: 1006, name: 'Miki', shop_id: 101, deleted_at: null },
    { id: 1007, name: 'Vienna', shop_id: 102, deleted_at: null },
    { id: 1008, name: 'Rin', shop_id: 102, deleted_at: null },
    { id: 1009, name: 'Mio', shop_id: 103, deleted_at: null },
    { id: 1010, name: 'Yua', shop_id: 103, deleted_at: null },
];

const DEMO_SOURCES = [
    { id: 11, name: 'Hot Pepper Beauty', platform_type: null },
    { id: 12, name: 'Instagram広告', platform_type: 'meta' },
    { id: 13, name: 'TikTok広告', platform_type: 'tiktok' },
    { id: 14, name: 'minimo', platform_type: null },
    { id: 15, name: 'ご紹介', platform_type: null },
];

// ---- シード付き乱数（mulberry32）----
function hashStr(s) {
    let h = 1779033703 ^ s.length;
    for (let i = 0; i < s.length; i++) {
        h = Math.imul(h ^ s.charCodeAt(i), 3432918353);
        h = (h << 13) | (h >>> 19);
    }
    return h >>> 0;
}
function rng(seed) {
    let a = hashStr(seed);
    return function () {
        a |= 0; a = (a + 0x6D2B79F5) | 0;
        let t = Math.imul(a ^ (a >>> 15), 1 | a);
        t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
        return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
}

function pad2(n) { return String(n).padStart(2, '0'); }

function* eachDay(from, to) {
    const d = new Date(`${from}T00:00:00Z`);
    const end = new Date(`${to}T00:00:00Z`);
    while (d <= end) {
        yield `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
        d.setUTCDate(d.getUTCDate() + 1);
    }
}

function todayJst() {
    const now = new Date(Date.now() + 9 * 3600 * 1000);
    return `${now.getUTCFullYear()}-${pad2(now.getUTCMonth() + 1)}-${pad2(now.getUTCDate())}`;
}

// 店舗ごとの日次売上・来店（未来日は0）
function dayStats(shopId, date) {
    const r = rng(`day:${shopId}:${date}`);
    if (date > todayJst()) {
        return { gross: 0, consumed: 0, digest: 0, newV: 0, repeatV: 0, cancel: 0, noShow: 0 };
    }
    const dow = new Date(`${date}T00:00:00Z`).getUTCDay();
    const weekendBoost = (dow === 0 || dow === 6) ? 1.45 : 1;
    const shopScale = shopId === 101 ? 1.5 : shopId === 102 ? 0.9 : 1.1;
    const visits = Math.max(0, Math.round((4 + r() * 9) * weekendBoost * shopScale));
    const newV = Math.min(visits, Math.round(visits * (0.18 + r() * 0.22)));
    const repeatV = visits - newV;
    const unit = 7200 + Math.round(r() * 4200);
    const gross = visits * unit + Math.round(r() * 15000);
    const consumed = Math.round(gross * (0.86 + r() * 0.2));
    const digest = Math.round(gross * (0.9 + r() * 0.1));
    const cancel = Math.round(r() * 2.4);
    const noShow = r() > 0.85 ? 1 : 0;
    return { gross, consumed, digest, newV, repeatV, cancel, noShow };
}

function sum(arr, f) { return arr.reduce((a, x) => a + f(x), 0); }

function salesSummary({ from, to, shop_id }) {
    const shops = shop_id ? DEMO_SHOPS.filter(s => String(s.id) === String(shop_id)) : DEMO_SHOPS;
    const byDay = [];
    for (const date of eachDay(from, to)) {
        const per = shops.map(s => dayStats(s.id, date));
        const gross = sum(per, p => p.gross);
        const r = rng(`pay:${date}`);
        byDay.push({
            date,
            gross_sales: gross,
            consumed_sales: sum(per, p => p.consumed),
            digest_sales: sum(per, p => p.digest),
            new_visit_count: sum(per, p => p.newV),
            repeat_visit_count: sum(per, p => p.repeatV),
            cancel_count: sum(per, p => p.cancel),
            no_show_count: sum(per, p => p.noShow),
            // 実APIと同様の支払い方法内訳
            payment_breakdown: gross > 0 ? (() => {
                const cash = Math.round(gross * (0.3 + r() * 0.2));
                const credit = Math.round(gross * (0.25 + r() * 0.15));
                const qr = Math.round(gross * (0.1 + r() * 0.1));
                return [
                    { payment_method_id: 1, name: '現金', amount: cash, is_sales: true },
                    { payment_method_id: 2, name: 'クレジットカード', amount: credit, is_sales: true },
                    { payment_method_id: 3, name: 'PayPay', amount: qr, is_sales: true },
                    { payment_method_id: 4, name: 'HPBポイント', amount: Math.max(0, gross - cash - credit - qr), is_sales: true },
                ];
            })() : [],
        });
    }
    const staffs = DEMO_STAFFS.filter(st => shops.some(s => s.id === st.shop_id));
    const totalGross = sum(byDay, d => d.gross_sales);
    const totalVisits = sum(byDay, d => d.new_visit_count + d.repeat_visit_count);
    // スタッフ別: 店舗合計を重み付きで按分
    const weights = staffs.map(st => 0.6 + rng(`staff:${st.id}:${from}`)() * 0.9);
    const wSum = weights.reduce((a, b) => a + b, 0) || 1;
    const byStaff = staffs.map((st, i) => {
        const ratio = weights[i] / wSum;
        const visits = Math.round(totalVisits * ratio);
        const newV = Math.round(visits * (0.2 + rng(`sn:${st.id}:${from}`)() * 0.2));
        const r = rng(`sx:${st.id}:${from}`);
        const gross = Math.round(totalGross * ratio);
        return {
            staff_id: st.id,
            staff_name: st.name,
            gross_sales: gross,
            consumed_sales: Math.round(gross * 0.92),
            digest_sales: Math.round(gross * 0.95),
            new_visit_count: newV,
            repeat_visit_count: Math.max(0, visits - newV),
            cancel_count: Math.round(rng(`sc:${st.id}:${from}`)() * 4),
            no_show_count: rng(`sns:${st.id}:${from}`)() > 0.8 ? 1 : 0,
            // 実APIの拡張フィールド相当
            product_sales: Math.round(gross * r() * 0.08),
            utilization_rate: Math.round(35 + r() * 50),
            operating_minutes: visits * 75,
            available_minutes: Math.round(visits * 75 / (0.35 + r() * 0.5)),
            google_review_count: Math.round(r() * 5),
            hotpepper_review_count: Math.round(r() * 8),
        };
    });
    return {
        from, to,
        shop_id: shop_id ? Number(shop_id) : null,
        gross_sales: totalGross,
        consumed_sales: sum(byDay, d => d.consumed_sales),
        digest_sales: sum(byDay, d => d.digest_sales),
        new_visit_count: sum(byDay, d => d.new_visit_count),
        repeat_visit_count: sum(byDay, d => d.repeat_visit_count),
        cancel_count: sum(byDay, d => d.cancel_count),
        no_show_count: sum(byDay, d => d.no_show_count),
        product_sales: sum(byStaff, s => s.product_sales),
        period_utilization_rate: 62.5,
        period_operating_minutes: sum(byStaff, s => s.operating_minutes),
        period_available_minutes: sum(byStaff, s => s.available_minutes),
        by_day: byDay,
        by_staff: byStaff.sort((a, b) => b.gross_sales - a.gross_sales),
    };
}

function marketingByChannel({ from, to, shop_id }) {
    const scale = shop_id ? 0.45 : 1;
    return DEMO_SOURCES.map(src => {
        const r = rng(`ch:${src.id}:${from}:${to}:${shop_id || 'all'}`);
        const booking = Math.round((14 + r() * 40) * scale);
        const cancel = Math.round(booking * (0.06 + r() * 0.12));
        const remaining = Math.round(booking * r() * 0.15);
        const visit = Math.max(0, booking - cancel - remaining);
        const join = Math.round(visit * (0.25 + r() * 0.4));
        const joinInPeriod = Math.min(join, Math.round(join * (0.6 + r() * 0.4)));
        const sales = visit * (9000 + Math.round(r() * 6000));
        const hasAds = src.platform_type !== null;
        const adSpend = hasAds ? Math.round((40000 + r() * 120000) * scale) : null;
        return {
            visit_source_id: src.id,
            name: src.name,
            platform_type: src.platform_type,
            booking_count: booking,
            visit_count: visit,
            remaining_count: remaining,
            cancel_count: cancel,
            cancel_rate: booking ? +(cancel / booking * 100).toFixed(1) : 0,
            join_count: join,
            join_rate: visit ? +(join / visit * 100).toFixed(1) : 0,
            join_rate_by_booking: booking ? +(join / booking * 100).toFixed(1) : 0,
            join_in_period_count: joinInPeriod,
            join_in_period_rate: visit ? +(joinInPeriod / visit * 100).toFixed(1) : 0,
            join_in_period_rate_by_booking: booking ? +(joinInPeriod / booking * 100).toFixed(1) : 0,
            sales,
            ad_spend: adSpend,
            impressions: hasAds ? Math.round(adSpend * (14 + r() * 20)) : null,
            clicks: hasAds ? Math.round(adSpend / (60 + r() * 90)) : null,
            cpa: hasAds && visit ? Math.round(adSpend / visit) : null,
            roas: hasAds && adSpend ? +(sales / adSpend * 100).toFixed(0) : null,
        };
    });
}

function marketingByStaff({ from, to, shop_id }) {
    const staffs = shop_id ? DEMO_STAFFS.filter(s => String(s.shop_id) === String(shop_id)) : DEMO_STAFFS;
    const rows = staffs.map(st => {
        const r = rng(`ms:${st.id}:${from}:${to}`);
        const booking = Math.round(6 + r() * 18);
        const cancel = Math.round(booking * r() * 0.15);
        const remaining = Math.round(booking * r() * 0.12);
        const visit = Math.max(0, booking - cancel - remaining);
        const purchase = Math.round(visit * (0.3 + r() * 0.45));
        const purchaseInPeriod = Math.min(purchase, Math.round(purchase * (0.55 + r() * 0.45)));
        const amount = purchase * (38000 + Math.round(r() * 40000));
        return {
            staff_id: st.id,
            staff_name: st.name,
            is_total: false,
            new_booking_count: booking,
            new_visit_count: visit,
            remaining_count: remaining,
            cancel_count: cancel,
            purchase_count: purchase,
            purchase_rate: booking ? +(purchase / booking * 100).toFixed(1) : 0,
            purchase_amount: amount,
            purchase_unit_price: purchase ? Math.round(amount / purchase) : 0,
            purchase_in_period_count: purchaseInPeriod,
            purchase_in_period_rate: booking ? +(purchaseInPeriod / booking * 100).toFixed(1) : 0,
            purchase_in_period_amount: Math.round(amount * (purchase ? purchaseInPeriod / purchase : 0)),
            new_customer_sales_total: visit * (11000 + Math.round(r() * 9000)),
        };
    });
    const total = {
        staff_id: null, staff_name: null, is_total: true,
        new_booking_count: sum(rows, x => x.new_booking_count),
        new_visit_count: sum(rows, x => x.new_visit_count),
        remaining_count: sum(rows, x => x.remaining_count),
        cancel_count: sum(rows, x => x.cancel_count),
        purchase_count: sum(rows, x => x.purchase_count),
        purchase_amount: sum(rows, x => x.purchase_amount),
        purchase_in_period_count: sum(rows, x => x.purchase_in_period_count),
        purchase_in_period_amount: sum(rows, x => x.purchase_in_period_amount),
        new_customer_sales_total: sum(rows, x => x.new_customer_sales_total),
    };
    total.purchase_rate = total.new_booking_count ? +(total.purchase_count / total.new_booking_count * 100).toFixed(1) : 0;
    total.purchase_in_period_rate = total.new_booking_count ? +(total.purchase_in_period_count / total.new_booking_count * 100).toFixed(1) : 0;
    total.purchase_unit_price = total.purchase_count ? Math.round(total.purchase_amount / total.purchase_count) : 0;
    return [total, ...rows.sort((a, b) => b.purchase_amount - a.purchase_amount)];
}

function marketingRetention({ from, to, shop_id }) {
    const r = rng(`ret:${from}:${to}:${shop_id || 'all'}`);
    const joined = Math.round((30 + r() * 50) * (shop_id ? 0.45 : 1));
    const active = Math.round(joined * (0.55 + r() * 0.3));
    const churned = joined - active;
    const summary = {
        join_count: joined,
        active_count: active,
        churn_count: churned,
        retention_rate: joined ? +(active / joined * 100).toFixed(1) : 0,
        avg_purchase_count: +(2.1 + r() * 3.2).toFixed(1),
    };
    const bySource = DEMO_SOURCES.map(src => {
        const rr = rng(`rets:${src.id}:${from}`);
        const j = Math.round(joined * (0.08 + rr() * 0.25));
        const a = Math.round(j * (0.5 + rr() * 0.4));
        return { visit_source_id: src.id, name: src.name, join_count: j, active_count: a, churn_count: j - a, retention_rate: j ? +(a / j * 100).toFixed(1) : 0 };
    });
    const staffs = shop_id ? DEMO_STAFFS.filter(s => String(s.shop_id) === String(shop_id)) : DEMO_STAFFS;
    const byStaff = staffs.map(st => {
        const rr = rng(`retst:${st.id}:${from}`);
        const j = Math.round(2 + rr() * 12);
        const a = Math.round(j * (0.45 + rr() * 0.45));
        return { staff_id: st.id, staff_name: st.name, join_count: j, active_count: a, churn_count: j - a, retention_rate: j ? +(a / j * 100).toFixed(1) : 0 };
    });
    const byPurchases = [1, 2, 3, 4, 5].map(n => {
        const rr = rng(`retp:${n}:${from}`);
        const c = Math.round(joined * Math.pow(0.62, n - 1) * (0.85 + rr() * 0.3));
        return { purchases: n, count: c, churn_count: Math.round(c * (0.32 - n * 0.04)) };
    });
    return { summary, by_source: bySource, by_staff: byStaff, by_purchases: byPurchases };
}

function customers({ limit = 200, cursor, shop_id }) {
    // 年代分析に必要な最小項目のみ（実APIの個人情報なしキー相当）
    const nowYear = new Date().getFullYear();
    const rows = [];
    for (let i = 0; i < 480; i++) {
        const r = rng(`cust:${i}`);
        const age = 18 + Math.floor(r() * 45);
        rows.push({
            id: 50000 + i,
            shop_id: DEMO_SHOPS[i % DEMO_SHOPS.length].id,
            visit_source_id: DEMO_SOURCES[Math.floor(r() * DEMO_SOURCES.length)].id,
            birth_year: nowYear - age,
            age_bracket: Math.floor(age / 10) * 10,
            deleted_at: r() > 0.96 ? '2026-01-15T00:00:00+00:00' : null,
        });
    }
    const filtered = shop_id ? rows.filter(c => String(c.shop_id) === String(shop_id)) : rows;
    const total = filtered.length;
    const start = cursor ? parseInt(Buffer.from(cursor, 'base64').toString(), 10) || 0 : 0;
    const n = Math.min(Number(limit) || 200, 1000);
    const data = filtered.slice(start, start + n);
    const next = start + n;
    return {
        data,
        meta: {
            returned: data.length,
            has_more: next < total,
            next_cursor: next < total ? Buffer.from(String(next)).toString('base64') : null,
            pii_included: false,
            schema_version: 'demo',
        },
    };
}

function listResponse(data) {
    return { data, meta: { returned: data.length, has_more: false, next_cursor: null, pii_included: false, schema_version: 'demo' } };
}

// エンドポイント別ディスパッチ
function demoFetch(path, params) {
    switch (path) {
        case 'meta':
            return { brand: { id: 1, name: 'vie (デモ)' }, key: { pii_included: false, expires_at: null }, schema_version: 'demo' };
        case 'sales/summary':
            return salesSummary(params);
        case 'marketing/by-channel':
            return { data: marketingByChannel(params) };
        case 'marketing/by-staff':
            return { data: marketingByStaff(params) };
        case 'marketing/retention':
            return marketingRetention(params);
        case 'shops':
            return listResponse(DEMO_SHOPS);
        case 'staffs':
            return listResponse(params.shop_id ? DEMO_STAFFS.filter(s => String(s.shop_id) === String(params.shop_id)) : DEMO_STAFFS);
        case 'visit-sources':
            return listResponse(DEMO_SOURCES);
        case 'customers':
            return customers(params);
        default:
            return listResponse([]);
    }
}

module.exports = { demoFetch, DEMO_SHOPS, DEMO_STAFFS };
