// GET /api/data/* — SalonOne 分析APIの認証付きプロキシ
// - セッション必須（クッキー）。store/staffロックのセッションは shop_id をサーバー側で強制
// - エンドポイントとパラメータはホワイトリスト方式
// - /customers は直接公開せず、集計済みの insights のみ返す（PII防御）

'use strict';

const { fetchSalonOne, fetchAllPages, stripCustomerPii, isDemo, UpstreamError } = require('./salonone');
const { getSession, passwordConfigStatus, accountsSummary } = require('./auth');
const { kvAvailable } = require('./kv');

const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function parseSubPath(req) {
    // 到達経路は3通り: Vercelのrewrite(?path=…) / キャッチオールのreq.query.path / 開発サーバーのURL直
    const q = req.query && req.query.path;
    if (Array.isArray(q)) return q.join('/');
    if (typeof q === 'string' && q) return q.replace(/\/+$/, '');
    const url = new URL(req.url, 'http://localhost');
    const fromQuery = url.searchParams.get('path');
    if (fromQuery) return fromQuery.replace(/\/+$/, '');
    const m = url.pathname.match(/\/api\/data\/(.+)$/);
    return m ? decodeURIComponent(m[1]).replace(/\/+$/, '') : '';
}

// ---- 実APIレスポンスの形状・型の正規化 ----
// 数値がJSON文字列で来ても集計が壊れないようNumberに揃える。nullはnullのまま。
function toNum(v) {
    if (v === null || v === undefined || v === '') return v ?? null;
    const n = Number(v);
    return isNaN(n) ? v : n;
}

function coerceNumbers(obj, keys) {
    if (!obj || typeof obj !== 'object') return obj;
    for (const k of keys) {
        if (k in obj) obj[k] = toNum(obj[k]);
    }
    return obj;
}

const SUMMARY_NUM_KEYS = [
    'gross_sales', 'consumed_sales', 'digest_sales', 'new_visit_count', 'repeat_visit_count', 'cancel_count', 'no_show_count',
    // 実APIの拡張フィールド（仕様書に未記載だが返ってくるもの）
    'product_sales', 'new_customer_sales', 'repeat_customer_sales', 'new_sales', 'repeat_sales', 'refund_amount', 'recovered_sales',
    'visit_count', 'new_count', 'repeat_count', 'treatment_count', 'item_count', 'avg_ticket',
    'utilization_rate', 'period_utilization_rate', 'operating_minutes', 'available_minutes',
    'google_review_count', 'hotpepper_review_count', 'gross_count', 'digest_count',
];
const CHANNEL_NUM_KEYS = ['booking_count', 'visit_count', 'remaining_count', 'cancel_count', 'cancel_rate', 'join_count', 'join_rate', 'join_rate_by_booking', 'join_in_period_count', 'join_in_period_rate', 'join_in_period_rate_by_booking', 'sales', 'ad_spend', 'impressions', 'clicks', 'cpa', 'roas'];
const MKSTAFF_NUM_KEYS = ['new_booking_count', 'new_visit_count', 'remaining_count', 'cancel_count', 'purchase_count', 'purchase_rate', 'purchase_amount', 'purchase_unit_price', 'purchase_in_period_count', 'purchase_in_period_rate', 'purchase_in_period_amount', 'new_customer_sales_total'];

function unwrapObject(raw) {
    // {data:{…}} 形式で包まれていたら中身を取り出す
    if (raw && !Array.isArray(raw) && raw.data && !Array.isArray(raw.data) && typeof raw.data === 'object') return raw.data;
    return raw;
}

// 実APIの by_day / by_staff は仕様書と異なるフィールド名を使う
// （by_day: new_count/repeat_count/visit_count、by_staff: new_count/treatment_count）。
// フロントが期待する new_visit_count / repeat_visit_count に揃える。
function aliasVisitCounts(row, { isStaff = false } = {}) {
    if (row.new_visit_count === undefined || row.new_visit_count === null) {
        if (row.new_count !== undefined) row.new_visit_count = row.new_count;
    }
    if (row.repeat_visit_count === undefined || row.repeat_visit_count === null) {
        if (row.repeat_count !== undefined && row.repeat_count !== null) {
            row.repeat_visit_count = row.repeat_count;
        } else if (isStaff && row.treatment_count !== undefined && row.treatment_count !== null) {
            row.repeat_visit_count = Math.max(0, row.treatment_count - (row.new_count || 0));
        } else if (row.visit_count !== undefined && row.visit_count !== null) {
            row.repeat_visit_count = Math.max(0, row.visit_count - (row.new_count || 0));
        }
    }
    return row;
}

function normalizeSummary(raw) {
    const obj = unwrapObject(raw) || {};
    coerceNumbers(obj, SUMMARY_NUM_KEYS);
    // 日別・スタッフ別のキー名の揺れを吸収
    const byDay = obj.by_day || obj.daily || obj.days || [];
    const byStaff = obj.by_staff || obj.staffs || obj.by_staffs || [];
    obj.by_day = (Array.isArray(byDay) ? byDay : []).map(d => aliasVisitCounts(coerceNumbers({ ...d }, SUMMARY_NUM_KEYS)));
    obj.by_staff = (Array.isArray(byStaff) ? byStaff : []).map(s => aliasVisitCounts(coerceNumbers({ ...s }, SUMMARY_NUM_KEYS), { isStaff: true }));
    return obj;
}

// スタッフ一覧は必要最小限の項目だけ返す（実APIは誕生日・電話番号などスタッフの個人情報を含むため）
const STAFF_ALLOWED_FIELDS = new Set(['id', 'shop_id', 'brand_id', 'name', 'is_public', 'employment_type', 'deleted_at']);
function stripStaffFields(row) {
    const out = {};
    for (const k of Object.keys(row || {})) {
        if (STAFF_ALLOWED_FIELDS.has(k)) out[k] = row[k];
    }
    return out;
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
        // 未設定はNumber(null)=0で「0歳代」に化けるため先に弾く
        if (row.age_bracket === null || row.age_bracket === undefined || row.age_bracket === '') { unknown++; continue; }
        const b = Number(row.age_bracket);
        // 年代として妥当な範囲（0〜90代）以外は入力ミス（誕生年の混入等）として「不明」に寄せる
        if (!isFinite(b) || b < 0 || b > 90) { unknown++; continue; }
        buckets[String(b)] = (buckets[String(b)] || 0) + 1;
    }
    return { total, unknown, buckets, truncated: !!truncated };
}

// 役割別に呼べるエンドポイントを制限する（クライアントのタブ非表示に頼らない）
// staffロールには広告経済指標（by-channelのad_spend/cpa/roas）や顧客集計を返さない。
const ROLE_DENIED_PATHS = {
    staff: new Set(['marketing/by-channel', 'insights/age-distribution', 'menus', 'menu-categories']),
    store: new Set([]),
    manager: new Set([]),
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
                    const raw = await fetchSalonOne('meta', {});
                    const meta = unwrapObject(raw) || {};
                    const b = meta.brand || meta.data?.brand;
                    brand = b ? { id: b.id ?? null, name: b.name ?? b.brand_name ?? null }
                        : (meta.brand_name || meta.name) ? { id: meta.brand_id ?? null, name: meta.brand_name || meta.name } : null;
                    schemaVersion = meta.schema_version || raw?.meta?.schema_version || null;
                    piiIncluded = meta.key?.pii_included ?? meta.pii_included ?? raw?.meta?.pii_included ?? null;
                } catch (_) { /* 疎通不可でも他の情報は返す */ }
                return res.end(JSON.stringify({
                    demo: isDemo(),
                    aiAvailable: !!process.env.GEMINI_API_KEY,
                    manualStorage: kvAvailable(),
                    // 設定タブの警告表示用（設定有無のみ）
                    passwords: session.role === 'admin' ? { ...passwordConfigStatus(), ...(await accountsSummary().catch(() => ({}))) } : undefined,
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
                if (path === 'sales/summary') {
                    return res.end(JSON.stringify(normalizeSummary(raw)));
                }
                if (path === 'marketing/retention') {
                    return res.end(JSON.stringify(unwrapObject(raw)));
                }
                let arr = Array.isArray(raw) ? raw : (raw.data || []);
                const numKeys = path === 'marketing/by-channel' ? CHANNEL_NUM_KEYS : MKSTAFF_NUM_KEYS;
                arr = arr.map(r => coerceNumbers({ ...r }, numKeys));
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
                const staffs = (Array.isArray(raw) ? raw : raw.data || [])
                    .filter(s => !s.deleted_at)
                    .map(stripStaffFields);
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

// テスト用に内部関数を公開
module.exports._internal = { normalizeSummary, aliasVisitCounts, stripStaffFields };
