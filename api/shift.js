// /api/shift — シフト希望休の申請・分配・承認（月単位・店舗ごと）
//
// 保存構造:
//   vie:shiftconfig            → { offDays, weekendOffDays, maxSameDayOff }
//   vie:shift:<YYYY-MM>        → { shops: { "<shopId>": {
//       requests: { "<staffId>": { days: ["YYYY-MM-DD", ...], submittedAt } },
//       assigned: { "<staffId>": ["YYYY-MM-DD", ...] },
//       status:   { "<staffId>": "requested" | "proposed" | "approved" },
//   } } }
//
// 権限:
//   staff  … 自分の希望申請のみ。自店舗のシフト状況は閲覧可
//   store  … 自店舗の申請代行・分配編集・承認
//   manager/admin … 全店舗 + ルール設定の変更

'use strict';

const { getSession, readJsonBody } = require('./_lib/auth');
const { kvAvailable, kvGet, kvSet } = require('./_lib/kv');
const { fetchSalonOne } = require('./_lib/salonone');

const MONTH_RE = /^\d{4}-\d{2}$/;
const DAY_RE = /^\d{4}-\d{2}-\d{2}$/;
const CONFIG_KEY = 'vie:shiftconfig';

const DEFAULT_CONFIG = { offDays: 8, weekendOffDays: 1, maxSameDayOff: 2 };

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function isAdminLike(session) {
    return session.role === 'admin' || session.role === 'manager';
}

function monthKey(month) {
    return `vie:shift:${month}`;
}

function emptyShop() {
    return { requests: {}, assigned: {}, status: {} };
}

function validDays(days, month) {
    if (!Array.isArray(days) || days.length > 31) return null;
    const out = [];
    const seen = new Set();
    for (const d of days) {
        if (typeof d !== 'string' || !DAY_RE.test(d) || !d.startsWith(month) || seen.has(d)) return null;
        seen.add(d);
        out.push(d);
    }
    return out;
}

function sanitizeConfig(raw) {
    const cfg = { ...DEFAULT_CONFIG };
    if (raw && typeof raw === 'object') {
        for (const k of ['offDays', 'weekendOffDays', 'maxSameDayOff']) {
            const n = Number(raw[k]);
            if (isFinite(n) && n >= 0 && n <= 31) cfg[k] = Math.round(n);
        }
    }
    if (cfg.weekendOffDays > cfg.offDays) cfg.weekendOffDays = cfg.offDays;
    if (cfg.maxSameDayOff < 1) cfg.maxSameDayOff = 1;
    return cfg;
}

async function staffShopId(session, staffId) {
    const raw = await fetchSalonOne('staffs', {});
    const staffs = Array.isArray(raw) ? raw : (raw.data || []);
    const st = staffs.find(s => String(s.id) === String(staffId) && !s.deleted_at);
    return st ? String(st.shop_id) : null;
}

// 閲覧スコープ: staff/storeは自店舗のみ
function scopeShops(shops, session) {
    if (isAdminLike(session)) return shops;
    const own = String(session.shopId);
    return shops[own] ? { [own]: shops[own] } : {};
}

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');

    const session = getSession(req);
    if (!session) return bad(res, 401, 'auth_required');

    try {
        if (req.method === 'GET') {
            const url = new URL(req.url, 'http://localhost');
            const month = url.searchParams.get('month') || '';
            if (!MONTH_RE.test(month)) return bad(res, 400, 'invalid_request', { fields: ['month'] });
            if (!kvAvailable()) {
                return res.end(JSON.stringify({ month, storage: 'none', config: sanitizeConfig(null), shops: {} }));
            }
            const [data, cfg] = await Promise.all([kvGet(monthKey(month)), kvGet(CONFIG_KEY)]);
            const shops = (data && data.shops) || {};
            return res.end(JSON.stringify({
                month, storage: 'kv',
                config: sanitizeConfig(cfg),
                shops: scopeShops(shops, session),
            }));
        }

        if (req.method !== 'POST') return bad(res, 405, 'method_not_allowed');
        if (!kvAvailable()) return bad(res, 501, 'storage_unconfigured', {
            detail: 'シフト管理にはサーバー保存が必要です。Supabase（または Upstash）のサーバー保存を設定してください（docs/SALONONE_INTEGRATION.md 参照）',
        });

        const body = await readJsonBody(req);
        const action = body.action;

        // ---- ルール設定の変更（オーナー・マネージャー）----
        if (action === 'config') {
            if (!isAdminLike(session)) return bad(res, 403, 'forbidden');
            const cfg = sanitizeConfig(body.config);
            await kvSet(CONFIG_KEY, cfg);
            return res.end(JSON.stringify({ ok: true, config: cfg }));
        }

        const month = body.month || '';
        if (!MONTH_RE.test(month)) return bad(res, 400, 'invalid_request', { fields: ['month'] });
        const key = monthKey(month);
        const data = (await kvGet(key)) || { shops: {} };
        if (!data.shops) data.shops = {};

        // ---- 希望休の申請 ----
        if (action === 'request') {
            let staffId = String(body.staffId || '');
            if (session.role === 'staff') staffId = String(session.staffId); // 本人のみ
            if (!/^\d+$/.test(staffId)) return bad(res, 400, 'invalid_request', { fields: ['staffId'] });

            const shopId = session.role === 'staff' ? String(session.shopId) : await staffShopId(session, staffId);
            if (!shopId) return bad(res, 404, 'unknown_staff');
            if (session.role === 'store' && shopId !== String(session.shopId)) {
                return bad(res, 403, 'forbidden', { detail: '他店舗のスタッフです' });
            }

            const days = validDays(body.days, month);
            if (days === null) return bad(res, 400, 'invalid_request', { fields: ['days'] });

            if (!data.shops[shopId]) data.shops[shopId] = emptyShop();
            const shop = data.shops[shopId];
            const approved = shop.status[staffId] === 'approved';
            if (approved && session.role === 'staff') {
                return bad(res, 409, 'already_approved', { detail: '承認済みのシフトです。変更は店長にご相談ください' });
            }
            shop.requests[staffId] = { days, submittedAt: new Date().toISOString() };
            delete shop.assigned[staffId];
            shop.status[staffId] = 'requested';
            await kvSet(key, data);
            return res.end(JSON.stringify({ ok: true, month, shops: scopeShops(data.shops, session) }));
        }

        // ---- 分配結果の保存（自動分配・手調整）----
        if (action === 'assign') {
            if (session.role === 'staff') return bad(res, 403, 'forbidden');
            const shopId = String(body.shopId || '');
            if (session.role === 'store' && shopId !== String(session.shopId)) return bad(res, 403, 'forbidden');
            if (!/^\d+$/.test(shopId)) return bad(res, 400, 'invalid_request', { fields: ['shopId'] });

            if (!data.shops[shopId]) data.shops[shopId] = emptyShop();
            const shop = data.shops[shopId];
            for (const [staffId, days] of Object.entries(body.assigned || {})) {
                if (!/^\d+$/.test(staffId)) return bad(res, 400, 'invalid_request', { fields: ['assigned'] });
                if (days === null) {
                    delete shop.assigned[staffId];
                    if (shop.status[staffId] === 'proposed') shop.status[staffId] = shop.requests[staffId] ? 'requested' : undefined;
                    continue;
                }
                const valid = validDays(days, month);
                if (valid === null) return bad(res, 400, 'invalid_request', { fields: ['assigned'] });
                shop.assigned[staffId] = valid.sort();
                if (shop.status[staffId] !== 'approved') shop.status[staffId] = 'proposed';
            }
            await kvSet(key, data);
            return res.end(JSON.stringify({ ok: true, month, shops: scopeShops(data.shops, session) }));
        }

        // ---- 承認 ----
        if (action === 'approve') {
            if (session.role === 'staff') return bad(res, 403, 'forbidden');
            const shopId = String(body.shopId || '');
            if (session.role === 'store' && shopId !== String(session.shopId)) return bad(res, 403, 'forbidden');
            const shop = data.shops[shopId];
            if (!shop) return bad(res, 404, 'not_found');
            const ids = Array.isArray(body.staffIds) ? body.staffIds.map(String) : [];
            for (const staffId of ids) {
                if (shop.assigned[staffId]) shop.status[staffId] = 'approved';
            }
            await kvSet(key, data);
            return res.end(JSON.stringify({ ok: true, month, shops: scopeShops(data.shops, session) }));
        }

        return bad(res, 400, 'invalid_request', { fields: ['action'] });
    } catch (e) {
        console.error('shift api error', e);
        return bad(res, 500, 'internal_error');
    }
};
