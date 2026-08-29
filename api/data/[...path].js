// GET /api/data/* — SalonOne 分析APIの認証付きプロキシ
// - セッション必須（クッキー）。store/staffロックのセッションは shop_id をサーバー側で強制
// - エンドポイントとパラメータはホワイトリスト方式
// - /customers は直接公開せず、集計済みの insights のみ返す（PII防御）

'use strict';

const { fetchSalonOne, fetchAllPages, stripCustomerPii, isDemo, UpstreamError } = require('../_lib/salonone');
const { getSession } = require('../_lib/auth');

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function parseSubPath(req) {
    // Vercelでは req.query.path が配列。開発サーバーではURLから切り出す。
    const q = req.query && req.query.path;
    if (Array.isArray(q)) return q.join('/');
    if (typeof q === 'string') return q;
    const url = new URL(req.url, 'http://localhost');
    const m = url.pathname.match(/\/api\/data\/(.+)$/);
    return m ? decodeURIComponent(m[1]).replace(/\/+$/, '') : '';
}

function pickParams(url, allowed) {
    const out = {};
    for (const k of allowed) {
        const v = url.searchParams.get(k);
        if (v !== null && v !== '') out[k] = v;
    }
    return out;
}

function validRange(params) {
    if (params.from && !DATE_RE.test(params.from)) return false;
    if (params.to && !DATE_RE.test(params.to)) return false;
    if (params.from && params.to) {
        const span = (new Date(params.to) - new Date(params.from)) / 86400000;
        if (span < 0 || span > 400) return false;
    }
    return true;
}

// 年代分布: /customers を全ページ集計して返す
// 仕様書§7「同じIDなら上書き」に従い、ページ間の重複行はIDで排除する。削除済み行も除外。
async function ageDistribution(params) {
    const { data, truncated } = await fetchAllPages('customers', params.shop_id ? { shop_id: params.shop_id } : {});
    const byId = new Map();
    for (const raw of data) {
        const row = stripCustomerPii(raw);
        if (row.id !== undefined && row.id !== null) byId.set(String(row.id), row);
    }
    const buckets = {};
    let total = 0, unknown = 0;
    for (const row of byId.values()) {
        if (row.deleted_at) continue;
        total++;
        if (row.age_bracket === null || row.age_bracket === undefined) { unknown++; continue; }
        const b = String(row.age_bracket);
        buckets[b] = (buckets[b] || 0) + 1;
    }
    return { total, unknown, buckets, truncated: !!truncated };
}

// 役割別に呼べるエンドポイントを制限する（クライアントのタブ非表示に頼らない）
// staffロールには広告経済指標（by-channelのad_spend/cpa/roas）や顧客集計を返さない。
const ROLE_DENIED_PATHS = {
    staff: new Set(['marketing/by-channel', 'insights/age-distribution', 'menus', 'menu-categories']),
    store: new Set([]),
    admin: new Set([]),
};

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'private, max-age=60');
    if (req.method !== 'GET') return bad(res, 405, 'method_not_allowed');

    const session = getSession(req);
    if (!session) {
        res.setHeader('Cache-Control', 'no-store');
        return bad(res, 401, 'auth_required');
    }

    const path = parseSubPath(req);
    const url = new URL(req.url, 'http://localhost');
    const locked = session.role === 'store' || session.role === 'staff';

    const denied = ROLE_DENIED_PATHS[session.role] || ROLE_DENIED_PATHS.staff;
    if (denied.has(path)) return bad(res, 403, 'forbidden', { path });

    try {
        switch (path) {
            case 'meta': {
                let brand = null, schemaVersion = null, piiIncluded = null;
                try {
                    const meta = await fetchSalonOne('meta', {});
                    brand = meta.brand || meta.data?.brand || null;
                    schemaVersion = meta.schema_version || meta.meta?.schema_version || null;
                    piiIncluded = meta.key?.pii_included ?? null;
                } catch (_) { /* 疎通不可でも他の情報は返す */ }
                return res.end(JSON.stringify({
                    demo: isDemo(),
                    aiAvailable: !!process.env.GEMINI_API_KEY,
                    brand, schemaVersion, piiIncluded,
                }));
            }

            case 'sales/summary':
            case 'marketing/by-channel':
            case 'marketing/by-staff':
            case 'marketing/retention': {
                const params = pickParams(url, ['from', 'to', 'shop_id']);
                if (!params.from || !params.to || !validRange(params)) return bad(res, 400, 'invalid_request', { fields: ['from', 'to'] });
                if (locked) params.shop_id = session.shopId; // 店舗スコープを強制
                const raw = await fetchSalonOne(path, params);
                // 実APIの形の揺れを吸収してから返す
                if (path === 'sales/summary' || path === 'marketing/retention') {
                    const obj = (raw && !Array.isArray(raw) && raw.data && !Array.isArray(raw.data)) ? raw.data : raw;
                    return res.end(JSON.stringify(obj));
                }
                let arr = Array.isArray(raw) ? raw : (raw.data || []);
                // 防御: ロックセッションのby-staffは自店舗スタッフの行だけに絞る
                // （実APIがマーケ集計のshop_idを無視した場合でも他店舗スタッフの数値を返さない）
                if (locked && path === 'marketing/by-staff') {
                    try {
                        const staffsRaw = await fetchSalonOne('staffs', { shop_id: session.shopId });
                        const ids = new Set((Array.isArray(staffsRaw) ? staffsRaw : staffsRaw.data || []).map(s => String(s.id)));
                        arr = arr.filter(r => r.is_total || r.staff_id === null || ids.has(String(r.staff_id)));
                    } catch (_) { /* スタッフ一覧が取れない場合はそのまま */ }
                }
                return res.end(JSON.stringify({ data: arr }));
            }

            case 'shops': {
                const raw = await fetchSalonOne('shops', {});
                let shops = (Array.isArray(raw) ? raw : raw.data || []).filter(s => !s.deleted_at);
                if (locked) shops = shops.filter(s => String(s.id) === String(session.shopId));
                return res.end(JSON.stringify({ data: shops }));
            }

            case 'staffs': {
                const params = pickParams(url, ['shop_id']);
                if (locked) params.shop_id = session.shopId;
                const raw = await fetchSalonOne('staffs', params);
                const staffs = (Array.isArray(raw) ? raw : raw.data || []).filter(s => !s.deleted_at);
                return res.end(JSON.stringify({ data: staffs }));
            }

            case 'visit-sources':
            case 'menu-categories':
            case 'menus': {
                const raw = await fetchSalonOne(path, {});
                const rows = (Array.isArray(raw) ? raw : raw.data || []).filter(s => !s.deleted_at);
                return res.end(JSON.stringify({ data: rows }));
            }

            case 'insights/age-distribution': {
                const params = pickParams(url, ['shop_id']);
                if (locked) params.shop_id = session.shopId;
                const dist = await ageDistribution(params);
                return res.end(JSON.stringify(dist));
            }

            default:
                return bad(res, 404, 'not_found', { path });
        }
    } catch (e) {
        if (e instanceof UpstreamError) {
            if (e.status === 429) {
                res.setHeader('Retry-After', String(e.retryAfter || 30));
                return bad(res, 429, 'rate_limited', { retryAfter: e.retryAfter || 30 });
            }
            if (e.status === 401 || e.status === 403) {
                // ダッシュボードの認証ではなく、SalonOneキーの設定問題
                return bad(res, 502, 'upstream_auth', { detail: 'SalonOne APIキーが無効か失効しています。運営に再発行を依頼してください。' });
            }
            return bad(res, 502, 'upstream_error', { status: e.status, code: e.code });
        }
        console.error('data proxy error', path, e);
        return bad(res, 500, 'internal_error');
    }
};
