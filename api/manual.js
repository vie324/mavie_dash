// /api/manual — SalonOne APIにないデータの手入力（月単位で保存）
//   daily:   { "YYYY-MM-DD:<staffId>": { blog, sns, reviews } }   ブログ/SNS更新・★5口コミ
//   monthly: { "<staffId>": { productSales } }                     物販売上（税込・インセンティブ用）
//   adCosts: { "<visitSourceId>"|"other": 金額 }                   広告費の手入力（APIにない媒体用）
// 保存先はUpstash Redis（Vercelの環境変数）。未設定時は storage:'none' を返し、
// クライアントはこの端末のみのlocalStorageに退避する。
// 権限: staffは自分のdailyのみ書き込み可 / storeは自店舗スタッフのdaily+monthly / adminは全て。

'use strict';

const { getSession, readJsonBody } = require('./_lib/auth');
const { kvAvailable, kvGet, kvSet } = require('./_lib/kv');
const { fetchSalonOne } = require('./_lib/salonone');

const MONTH_RE = /^\d{4}-\d{2}$/;

// admin(オーナー)とmanager(マネージャー)は手入力データを全店舗分扱える
function isAdminLike(session) {
    return session.role === 'admin' || session.role === 'manager';
}
const DAILY_KEY_RE = /^\d{4}-\d{2}-\d{2}:\d+$/;
const DAILY_FIELDS = new Set(['blog', 'sns', 'reviews']);
const MONTHLY_FIELDS = new Set(['productSales']);

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function emptyData() {
    return { daily: {}, monthly: {}, adCosts: {} };
}

function validNum(v) {
    const n = Number(v);
    return isFinite(n) && n >= 0 && n <= 1e9 ? Math.round(n) : null;
}

async function shopStaffIds(shopId) {
    const raw = await fetchSalonOne('staffs', { shop_id: shopId });
    const staffs = Array.isArray(raw) ? raw : (raw.data || []);
    return new Set(staffs.filter(s => !s.deleted_at).map(s => String(s.id)));
}

// patch を検証しつつ data にマージする。不正なキー・値は拒否。
function applyPatch(data, patch, session, allowedStaffIds) {
    for (const [key, entry] of Object.entries(patch.daily || {})) {
        if (!DAILY_KEY_RE.test(key)) throw { code: 'invalid_request', detail: `daily key: ${key}` };
        const staffId = key.split(':')[1];
        if (session.role === 'staff' && String(session.staffId) !== staffId) {
            throw { code: 'forbidden', detail: '自分以外の日報は入力できません' };
        }
        if (session.role === 'store' && allowedStaffIds && !allowedStaffIds.has(staffId)) {
            throw { code: 'forbidden', detail: '他店舗のスタッフです' };
        }
        if (entry === null) { delete data.daily[key]; continue; }
        const cur = data.daily[key] || {};
        for (const [f, v] of Object.entries(entry)) {
            if (!DAILY_FIELDS.has(f)) throw { code: 'invalid_request', detail: `daily field: ${f}` };
            if (v === null) { delete cur[f]; continue; }
            const n = validNum(v);
            if (n === null) throw { code: 'invalid_request', detail: `daily value: ${f}` };
            cur[f] = n;
        }
        if (Object.keys(cur).length === 0) delete data.daily[key];
        else data.daily[key] = cur;
    }

    for (const [staffId, entry] of Object.entries(patch.monthly || {})) {
        if (session.role === 'staff') throw { code: 'forbidden', detail: '月次項目は管理者・店舗のみ入力できます' };
        if (!/^\d+$/.test(staffId)) throw { code: 'invalid_request', detail: `monthly key: ${staffId}` };
        if (session.role === 'store' && allowedStaffIds && !allowedStaffIds.has(staffId)) {
            throw { code: 'forbidden', detail: '他店舗のスタッフです' };
        }
        if (entry === null) { delete data.monthly[staffId]; continue; }
        const cur = data.monthly[staffId] || {};
        for (const [f, v] of Object.entries(entry)) {
            if (!MONTHLY_FIELDS.has(f)) throw { code: 'invalid_request', detail: `monthly field: ${f}` };
            if (v === null) { delete cur[f]; continue; }
            const n = validNum(v);
            if (n === null) throw { code: 'invalid_request', detail: `monthly value: ${f}` };
            cur[f] = n;
        }
        if (Object.keys(cur).length === 0) delete data.monthly[staffId];
        else data.monthly[staffId] = cur;
    }

    for (const [sourceId, v] of Object.entries(patch.adCosts || {})) {
        if (!isAdminLike(session)) throw { code: 'forbidden', detail: '広告費はオーナー・マネージャーのみ入力できます' };
        if (!/^\d+$|^other$/.test(sourceId)) throw { code: 'invalid_request', detail: `adCosts key: ${sourceId}` };
        if (v === null) { delete data.adCosts[sourceId]; continue; }
        const n = validNum(v);
        if (n === null) throw { code: 'invalid_request', detail: 'adCosts value' };
        data.adCosts[sourceId] = n;
    }
}

// ロック済みロールには自店舗スタッフ分のみ返す（adCostsは管理者のみ）
function scopeData(data, session, allowedStaffIds) {
    if (isAdminLike(session)) return data;
    const out = emptyData();
    for (const [key, entry] of Object.entries(data.daily)) {
        if (allowedStaffIds.has(key.split(':')[1])) out.daily[key] = entry;
    }
    for (const [staffId, entry] of Object.entries(data.monthly)) {
        if (allowedStaffIds.has(staffId)) out.monthly[staffId] = entry;
    }
    return out;
}

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');

    const session = getSession(req);
    if (!session) return bad(res, 401, 'auth_required');

    const url = new URL(req.url, 'http://localhost');
    const month = url.searchParams.get('month') || '';
    if (!MONTH_RE.test(month)) return bad(res, 400, 'invalid_request', { fields: ['month'] });
    const key = `vie:manual:${month}`;

    try {
        if (req.method === 'GET') {
            if (!kvAvailable()) {
                return res.end(JSON.stringify({ month, storage: 'none', ...emptyData() }));
            }
            const data = { ...emptyData(), ...(await kvGet(key) || {}) };
            const allowed = isAdminLike(session) ? null : await shopStaffIds(session.shopId);
            const scoped = isAdminLike(session) ? data : scopeData(data, session, allowed);
            return res.end(JSON.stringify({ month, storage: 'kv', ...scoped }));
        }

        if (req.method === 'POST') {
            if (!kvAvailable()) return bad(res, 501, 'storage_unconfigured', {
                detail: 'Vercelで Upstash for Redis 連携を追加すると全端末で共有保存できます（docs/SALONONE_INTEGRATION.md 参照）',
            });
            const body = await readJsonBody(req);
            const patch = body.patch || {};
            const allowed = isAdminLike(session) ? null : await shopStaffIds(session.shopId);
            const data = { ...emptyData(), ...(await kvGet(key) || {}) };
            try {
                applyPatch(data, patch, session, allowed);
            } catch (e) {
                if (e.code) return bad(res, e.code === 'forbidden' ? 403 : 400, e.code, { detail: e.detail });
                throw e;
            }
            // サイズ暴走の防止（1ヶ月あたり200KB上限）
            const serialized = JSON.stringify(data);
            if (serialized.length > 200 * 1024) return bad(res, 413, 'too_large');
            await kvSet(key, data);
            const scoped = isAdminLike(session) ? data : scopeData(data, session, allowed);
            return res.end(JSON.stringify({ ok: true, month, storage: 'kv', ...scoped }));
        }

        return bad(res, 405, 'method_not_allowed');
    } catch (e) {
        console.error('manual api error', e);
        return bad(res, 500, 'internal_error');
    }
};
