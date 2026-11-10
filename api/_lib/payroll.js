// 給与・インセンティブ（2026年11月分〜の新ルール）の設定・勤怠の手入力・給与計算用の売上
// 呼び出し: /api/goals?kind=payroll&month=YYYY-MM （api/goals.js から振り分け）
//   ※ Vercel の Hobby プランは関数（api/ 直下のファイル）が12個までのため、新しい関数ファイルは作らず既存の goals に同居させている
//
// 保存構造（Supabase の vie_kv に新しいキーを足すだけ。テーブルの形は変えない）:
//   vie:payroll:staff           → { "<staffId>": [ { from: "YYYY-MM", pattern: "full" | "short_atsugi" | "short_chiba" | null, storeDuty: true | false, rates?: [段階ごとの率(%)] } ] }
//                                  「◯月から」の設定の履歴。ある月には from がその月以前でいちばん新しい設定を使う
//                                  （過去の月は当時の設定のまま計算される）。pattern: null は「その月から対象外」
//                                  rates はスタッフ個別の施術インセンティブの率（段階の金額はパターンのまま）。無ければパターンの率
//   vie:payroll:hours:<YYYY-MM> → { "<staffId>": { overtime, statutoryOvertime, totalHours } }
//                                  残業時間（所定超え）・法定時間外の時間（1日8h/週44h超え）・総労働時間（勤怠データがないため手入力）
// 既存のキー（vie:salaries の基本給、vie:manual:* の日報・物販など）には一切さわらない。
//
// 権限（サーバー側で強制）:
//   admin（オーナー） … 全員分の閲覧・変更
//   staff（スタッフ） … 本人の分の閲覧のみ（変更不可）
//   manager・store    … 不可
//
// GET  /api/goals?kind=payroll&month=YYYY-MM
//   → { storage, month, settings: { "<staffId>": 履歴[] }, hours: { "<staffId>": {...} }, sales: { "<staffId>": {...} } | null }
//   sales = その月の SalonOne のスタッフ別売上（全店舗の合計。画面の店舗の絞り込みに左右されないように、ここで取る）
//           スタッフには本人の行だけ返す。どの売上項目を使うかは画面側の payroll/config.js（salesField）で決める
// POST /api/goals?kind=payroll&month=YYYY-MM  body: { patch: { settings?: { "<staffId>": { pattern, storeDuty } | null }, hours?: { "<staffId>": {...} | null } } }
//   settings の値: その月から適用する設定（同じ月の設定は置き換え）。null はその月の設定を取り消す（前の設定に戻る）
//                  { pattern, storeDuty, rates }。rates を送らない・null ならパターンの率（個別の率なし）

'use strict';

const { getSession, readJsonBody } = require('./auth');
const { kvAvailable, kvGet, kvUpdate } = require('./kv');
const { fetchSalonOne } = require('./salonone');
const { normalizeSummary } = require('./data-handler')._internal;

const SETTINGS_KEY = 'vie:payroll:staff';
const hoursKey = month => `vie:payroll:hours:${month}`;
const MONTH_RE = /^\d{4}-(0[1-9]|1[0-2])$/;
const STAFF_ID_RE = /^\d+$/;
const PATTERN_RE = /^[a-z][a-z0-9_]{0,31}$/;
const HOURS_FIELDS = ['overtime', 'statutoryOvertime', 'totalHours'];
const MAX_TIERS = 10;
const MAX_HOURS = 744; // 31日 × 24時間
const MAX_BYTES = 256 * 1024;

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function sameSetting(a, b) {
    return !!a && !!b
        && (a.pattern || null) === (b.pattern || null)
        && (a.storeDuty ?? null) === (b.storeDuty ?? null)
        && JSON.stringify(a.rates || null) === JSON.stringify(b.rates || null);
}

// 個別の掛け率（段階ごとの率・%）: 1〜10個、どれも 0〜100（小数第2位まで）。null・未指定は「個別の率なし」
function cleanRates(v) {
    if (v === null || v === undefined) return null;
    if (!Array.isArray(v) || v.length < 1 || v.length > MAX_TIERS) throw { code: 'invalid_request', detail: 'rates' };
    return v.map(x => {
        const n = Number(x);
        if (x === '' || x === null || !isFinite(n) || n < 0 || n > 100) throw { code: 'invalid_request', detail: 'rates value' };
        return Math.round(n * 100) / 100;
    });
}

// 履歴を月順に並べ、意味のない行（最初の「対象外」・直前と同じ設定）を取り除く
function normalizeHistory(list) {
    const sorted = [...list].sort((a, b) => a.from.localeCompare(b.from));
    const out = [];
    for (const e of sorted) {
        const prev = out[out.length - 1];
        if (!prev && !e.pattern) continue;
        if (prev && sameSetting(prev, e)) continue;
        out.push(e);
    }
    return out;
}

// patch: { "<staffId>": { pattern, storeDuty } | null } を month から適用する
function applySettingsPatch(settings, patch, month) {
    for (const [staffId, v] of Object.entries(patch || {})) {
        if (!STAFF_ID_RE.test(staffId)) throw { code: 'invalid_request', detail: `staff: ${staffId}` };
        const list = (Array.isArray(settings[staffId]) ? settings[staffId] : []).filter(e => e && e.from !== month);
        if (v !== null) {
            if (typeof v !== 'object' || Array.isArray(v)) throw { code: 'invalid_request', detail: `setting: ${staffId}` };
            const pattern = v.pattern === null || v.pattern === '' || v.pattern === undefined ? null : String(v.pattern);
            if (pattern !== null && !PATTERN_RE.test(pattern)) throw { code: 'invalid_request', detail: `pattern: ${pattern}` };
            if (v.storeDuty !== undefined && v.storeDuty !== null && typeof v.storeDuty !== 'boolean') throw { code: 'invalid_request', detail: 'storeDuty' };
            const rates = cleanRates(v.rates);
            const entry = { from: month, pattern };
            if (pattern !== null && typeof v.storeDuty === 'boolean') entry.storeDuty = v.storeDuty;
            if (pattern !== null && rates) entry.rates = rates;
            list.push(entry);
        }
        const next = normalizeHistory(list);
        if (next.length) settings[staffId] = next;
        else delete settings[staffId];
    }
    return settings;
}

function validHours(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    if (!isFinite(n) || n < 0 || n > MAX_HOURS) return undefined;
    return Math.round(n * 100) / 100;
}

// patch: { "<staffId>": { overtime, statutoryOvertime, totalHours } | null }
function applyHoursPatch(hours, patch) {
    for (const [staffId, v] of Object.entries(patch || {})) {
        if (!STAFF_ID_RE.test(staffId)) throw { code: 'invalid_request', detail: `staff: ${staffId}` };
        if (v === null) { delete hours[staffId]; continue; }
        if (typeof v !== 'object' || Array.isArray(v)) throw { code: 'invalid_request', detail: `hours: ${staffId}` };
        const cur = { ...(hours[staffId] || {}) };
        for (const [f, raw] of Object.entries(v)) {
            if (!HOURS_FIELDS.includes(f)) throw { code: 'invalid_request', detail: `hours field: ${f}` };
            const n = validHours(raw);
            if (n === undefined) throw { code: 'invalid_request', detail: `hours value: ${f}` };
            if (n === null) delete cur[f]; else cur[f] = n;
        }
        if (Object.keys(cur).length) hours[staffId] = cur;
        else delete hours[staffId];
    }
    return hours;
}

// スタッフには本人の分だけ返す
function scopeFor(session, settings, hours, sales) {
    if (session.role === 'admin') return sales === undefined ? { settings, hours } : { settings, hours, sales };
    const own = String(session.staffId);
    const out = {
        settings: settings[own] ? { [own]: settings[own] } : {},
        hours: hours[own] ? { [own]: hours[own] } : {},
    };
    if (sales !== undefined) out.sales = sales === null ? null : (sales[own] ? { [own]: sales[own] } : {});
    return out;
}

function monthRange(month) {
    const [y, m] = month.split('-').map(Number);
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    return { from: `${month}-01`, to: `${month}-${String(last).padStart(2, '0')}` };
}

// 給与計算に使う売上の項目だけを取り出す（来店数などは返さない）
const SALES_FIELDS = ['digest_sales', 'consumed_sales', 'gross_sales', 'product_sales', 'refund_amount'];
function pickSales(summary) {
    const out = {};
    for (const r of summary.by_staff || []) {
        if (r.staff_id === null || r.staff_id === undefined) continue;
        const row = { staff_name: r.staff_name ?? null };
        for (const f of SALES_FIELDS) if (r[f] !== undefined) row[f] = r[f];
        out[String(r.staff_id)] = row;
    }
    return out;
}

// その月の全店舗のスタッフ別売上（SalonOne。サーバー側で5分キャッシュ）。取れなければ null
async function payrollSales(month) {
    try {
        return pickSales(normalizeSummary(await fetchSalonOne('sales/summary', monthRange(month))));
    } catch (e) {
        console.warn('payroll sales fetch failed', e && e.message);
        return null;
    }
}

function canRead(session) {
    return session.role === 'admin' || (session.role === 'staff' && session.staffId !== null && session.staffId !== undefined);
}

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');

    const session = getSession(req);
    if (!session) return bad(res, 401, 'auth_required');
    if (!canRead(session)) return bad(res, 403, 'forbidden', { detail: '給与はオーナーと本人のみ閲覧できます' });

    const url = new URL(req.url, 'http://localhost');
    const month = url.searchParams.get('month') || '';
    if (!MONTH_RE.test(month)) return bad(res, 400, 'invalid_request', { fields: ['month'] });

    try {
        if (req.method === 'GET') {
            const salesP = payrollSales(month);
            if (!kvAvailable()) {
                return res.end(JSON.stringify({ month, storage: 'none', ...scopeFor(session, {}, {}, await salesP) }));
            }
            const [settings, hours, sales] = await Promise.all([kvGet(SETTINGS_KEY), kvGet(hoursKey(month)), salesP]);
            return res.end(JSON.stringify({ month, storage: 'kv', ...scopeFor(session, settings || {}, hours || {}, sales) }));
        }

        if (req.method !== 'POST') return bad(res, 405, 'method_not_allowed');
        if (session.role !== 'admin') return bad(res, 403, 'forbidden', { detail: '給与の設定・勤怠の入力はオーナーのみ可能です' });
        if (!kvAvailable()) return bad(res, 501, 'storage_unconfigured', {
            detail: '給与の設定はサーバー保存（Supabase または Upstash）が必要です（docs/SALONONE_INTEGRATION.md 参照）',
        });

        const body = await readJsonBody(req);
        const patch = body.patch || {};
        let rejected = null;
        const guarded = apply => current => {
            try {
                const next = apply({ ...(current || {}) });
                if (JSON.stringify(next).length > MAX_BYTES) { rejected = { code: 'too_large' }; return null; }
                return next;
            } catch (e) {
                if (e && e.code) { rejected = e; return null; }
                throw e;
            }
        };

        let settings = null;
        if (patch.settings && Object.keys(patch.settings).length) {
            settings = await kvUpdate(SETTINGS_KEY, guarded(cur => applySettingsPatch(cur, patch.settings, month)));
            if (rejected) return bad(res, rejected.code === 'too_large' ? 413 : 400, rejected.code, { detail: rejected.detail });
        }
        let hours = null;
        if (patch.hours && Object.keys(patch.hours).length) {
            hours = await kvUpdate(hoursKey(month), guarded(cur => applyHoursPatch(cur, patch.hours)));
            if (rejected) return bad(res, rejected.code === 'too_large' ? 413 : 400, rejected.code, { detail: rejected.detail });
        }
        if (settings === null) settings = (await kvGet(SETTINGS_KEY)) || {};
        if (hours === null) hours = (await kvGet(hoursKey(month))) || {};
        return res.end(JSON.stringify({ ok: true, month, storage: 'kv', settings, hours }));
    } catch (e) {
        console.error('payroll api error', e);
        return bad(res, 500, 'internal_error');
    }
};

// テスト用に内部関数を公開
module.exports._internal = { applySettingsPatch, applyHoursPatch, normalizeHistory, scopeFor, canRead, pickSales, monthRange };
