// /api/manual — SalonOne APIにないデータの手入力（月単位で保存）
//   daily:   { "YYYY-MM-DD:<staffId>": { nextNew, nextRepeat, repeatNo, src, blog, sns, reviews } }
//            次回予約(新規/2回目以降)・ブログ/SNS更新・★5口コミ
//            src = 新規の次回予約の媒体別の内訳 { "<visitSourceId>"|"other": { n: 新規, r: 2回目以降 } }。
//            src がある日報は nextNew をサーバーが src の合計から作る（内訳と合計がずれないように）。
//            2回目以降は媒体を問わず、nextRepeat（取れた人数）と repeatNo（取れなかった人数）だけを記録する。
//            r は媒体別に2回目以降を入力していた頃の値（nextRepeat が送られない古い画面からの保存では r の合計を使う）
//   monthly: { "<staffId>": { productSales } }                     物販売上（税込・インセンティブ用）
//   adCosts: { "<visitSourceId>"|"other": 金額 }                   広告費の手入力（APIにない媒体用）
//   recon:   { "YYYY-MM-DD:<shopId>": { "m<支払方法ID>": 実際額, memo } } 入金突合の実際額（レジ実査・端末集計）
// 保存先はSupabase / Upstash（api/_lib/kv.js、Vercelの環境変数で選択）。未設定時は storage:'none' を返し、
// クライアントはこの端末のみのlocalStorageに退避する。
// 保存はキー単位のロック付き read-modify-write（複数スタッフの同時保存で後勝ち消失しないように）。
// daily の各エントリにはサーバーが保存時刻 at（UNIX秒）を付与する（「保存済み 20:10」表示用）。
// 権限: staffは自分のdailyのみ書き込み可 / storeは自店舗スタッフのdaily+monthly / adminは全て。

'use strict';

const { getSession, readJsonBody } = require('./_lib/auth');
const { kvAvailable, kvGet, kvUpdate } = require('./_lib/kv');
const { fetchSalonOne } = require('./_lib/salonone');

const MONTH_RE = /^\d{4}-\d{2}$/;

// admin(オーナー)とmanager(マネージャー)は手入力データを全店舗分扱える
function isAdminLike(session) {
    return session.role === 'admin' || session.role === 'manager';
}
const DAILY_KEY_RE = /^\d{4}-\d{2}-\d{2}:\d+$/;
const DAILY_FIELDS = new Set(['blog', 'sns', 'reviews', 'nextNew', 'nextRepeat', 'repeatNo']);
const SRC_KEY_RE = /^(\d+|other)$/;   // 媒体: SalonOneの流入元ID / "other"（その他・不明）
const MAX_SRC_KEYS = 60;
const MAX_NEXT_PER_DAY = 999;
const MAX_MONTH_BYTES = 400 * 1024;
const MONTHLY_FIELDS = new Set(['productSales']);
const RECON_KEY_RE = /^\d{4}-\d{2}-\d{2}:\d+$/;       // "日付:shopId"
const RECON_FIELD_RE = /^m\d+$/;                        // "m<payment_method_id>"

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function emptyData() {
    return { daily: {}, monthly: {}, adCosts: {}, recon: {} };
}

function validNum(v) {
    const n = Number(v);
    return isFinite(n) && n >= 0 && n <= 1e9 ? Math.round(n) : null;
}

// 次回予約の媒体別の内訳を検証して整える（0件の媒体は保存しない）
function cleanSrc(src) {
    if (!src || typeof src !== 'object' || Array.isArray(src)) throw { code: 'invalid_request', detail: 'daily src' };
    const keys = Object.keys(src);
    if (keys.length > MAX_SRC_KEYS) throw { code: 'invalid_request', detail: 'daily src: too many sources' };
    const out = {};
    for (const k of keys) {
        if (!SRC_KEY_RE.test(k)) throw { code: 'invalid_request', detail: `daily src key: ${k}` };
        const cell = src[k];
        if (cell === null || cell === undefined) continue;
        if (typeof cell !== 'object' || Array.isArray(cell)) throw { code: 'invalid_request', detail: `daily src value: ${k}` };
        const count = v => {
            if (v === undefined || v === null || v === '') return 0;
            const n = Number(v);
            return Number.isInteger(n) && n >= 0 && n <= MAX_NEXT_PER_DAY ? n : null;
        };
        const n = count(cell.n), r = count(cell.r);
        if (n === null || r === null) throw { code: 'invalid_request', detail: `daily src value: ${k}` };
        if (n > 0 || r > 0) out[k] = { n, r };
    }
    return out;
}

function srcTotals(src) {
    let n = 0, r = 0;
    for (const cell of Object.values(src || {})) { n += cell.n || 0; r += cell.r || 0; }
    return { n, r };
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
        if (typeof entry !== 'object' || Array.isArray(entry)) throw { code: 'invalid_request', detail: `daily entry: ${key}` };
        const cur = data.daily[key] || {};
        let totalsTouched = false;
        for (const [f, v] of Object.entries(entry)) {
            if (f === 'at' || f === 'src') continue; // 保存時刻はサーバーが付与する / 内訳は下で処理
            if (!DAILY_FIELDS.has(f)) throw { code: 'invalid_request', detail: `daily field: ${f}` };
            if (f === 'nextNew' || f === 'nextRepeat') totalsTouched = true;
            if (v === null) { delete cur[f]; continue; }
            const n = validNum(v);
            if (n === null) throw { code: 'invalid_request', detail: `daily value: ${f}` };
            cur[f] = n;
        }
        if ('src' in entry) {
            if (entry.src === null) {
                delete cur.src;
            } else {
                const src = cleanSrc(entry.src);
                if (Object.keys(src).length) cur.src = src;
                else delete cur.src;
                const t = srcTotals(src);
                cur.nextNew = t.n;
                // 2回目以降は媒体を問わない入力（nextRepeat）を優先。無ければ内訳の r の合計（旧画面）
                if (entry.nextRepeat === undefined || entry.nextRepeat === null) cur.nextRepeat = t.r;
            }
        } else if (totalsTouched && cur.src) {
            // 合計だけを書き換える古い画面からの保存: 内訳と合わなくなったら内訳を外す
            const t = srcTotals(cur.src);
            if (t.n !== (cur.nextNew || 0)) delete cur.src;
        }
        const hasValue = [...DAILY_FIELDS].some(f => cur[f] !== undefined);
        if (!hasValue) delete data.daily[key];
        else { cur.at = Math.floor(Date.now() / 1000); data.daily[key] = cur; }
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

    // 入金突合の実際額（スタッフは不可、店長は自店舗のみ）
    for (const [key, entry] of Object.entries(patch.recon || {})) {
        if (session.role === 'staff') throw { code: 'forbidden', detail: '入金突合はスタッフは入力できません' };
        if (!RECON_KEY_RE.test(key)) throw { code: 'invalid_request', detail: `recon key: ${key}` };
        const shopId = key.split(':')[1];
        if (session.role === 'store' && String(session.shopId) !== shopId) {
            throw { code: 'forbidden', detail: '他店舗の入金突合は入力できません' };
        }
        if (entry === null) { delete data.recon[key]; continue; }
        const cur = data.recon[key] || {};
        for (const [f, v] of Object.entries(entry)) {
            if (f === 'memo') {
                if (v === null || v === '') { delete cur.memo; continue; }
                if (typeof v !== 'string' || v.length > 200) throw { code: 'invalid_request', detail: 'recon memo' };
                cur.memo = v;
                continue;
            }
            if (!RECON_FIELD_RE.test(f)) throw { code: 'invalid_request', detail: `recon field: ${f}` };
            if (v === null) { delete cur[f]; continue; }
            const n = validNum(v);
            if (n === null) throw { code: 'invalid_request', detail: `recon value: ${f}` };
            cur[f] = n;
        }
        if (Object.keys(cur).length === 0) delete data.recon[key];
        else data.recon[key] = cur;
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
    if (session.role === 'store') {
        for (const [key, entry] of Object.entries(data.recon || {})) {
            if (key.split(':')[1] === String(session.shopId)) out.recon[key] = entry;
        }
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
                detail: 'Supabase（または Upstash）のサーバー保存を設定すると全端末で共有保存できます（docs/SALONONE_INTEGRATION.md 参照）',
            });
            const body = await readJsonBody(req);
            const patch = body.patch || {};
            const allowed = isAdminLike(session) ? null : await shopStaffIds(session.shopId);
            let rejected = null;
            let data = await kvUpdate(key, current => {
                const next = { ...emptyData(), ...(current || {}) };
                try {
                    applyPatch(next, patch, session, allowed);
                } catch (e) {
                    if (e.code) { rejected = e; return null; }
                    throw e;
                }
                // サイズ暴走の防止（1ヶ月あたり400KB上限。媒体別の内訳込みで約30名×31日でも余裕がある）
                if (JSON.stringify(next).length > MAX_MONTH_BYTES) { rejected = { code: 'too_large' }; return null; }
                return next;
            });
            if (rejected) {
                const status = rejected.code === 'forbidden' ? 403 : rejected.code === 'too_large' ? 413 : 400;
                return bad(res, status, rejected.code, { detail: rejected.detail });
            }
            if (!data) data = { ...emptyData(), ...(await kvGet(key) || {}) };
            const scoped = isAdminLike(session) ? data : scopeData(data, session, allowed);
            return res.end(JSON.stringify({ ok: true, month, storage: 'kv', ...scoped }));
        }

        return bad(res, 405, 'method_not_allowed');
    } catch (e) {
        console.error('manual api error', e);
        return bad(res, 500, 'internal_error');
    }
};

// テスト用に内部関数を公開
module.exports._internal = { applyPatch, cleanSrc, emptyData };
