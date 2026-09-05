// /api/goals — 月次目標（売上・新規来店・入会数）と基本給のサーバー保存
//
// 保存構造:
//   vie:goals    → { "YYYY-MM": { "all" | "shop:<shopId>" | "staff:<staffId>": { sales, newVisits, joins } } }
//   vie:salaries → { "<staffId>": 基本給（円/月） }
//
// 権限:
//   閲覧 … admin/manager = 全体、store = 自店舗 + 所属スタッフ、staff = 自店舗 + 自分
//   更新 … admin/manager = 全体、store = 自店舗 + 所属スタッフ、staff = 不可
//   基本給 … admin のみ（閲覧・更新とも）
// 保存先はSupabase / Upstash（api/_lib/kv.js）。未設定時は storage:'none' を返し、クライアントはlocalStorageに退避する。

'use strict';

const { getSession, readJsonBody } = require('./_lib/auth');
const { kvAvailable, kvGet, kvUpdate } = require('./_lib/kv');
const { fetchSalonOne } = require('./_lib/salonone');

const GOALS_KEY = 'vie:goals';
const SALARIES_KEY = 'vie:salaries';
const MONTH_RE = /^\d{4}-\d{2}$/;
const SCOPE_RE = /^(all|shop:\d+|staff:\d+)$/;
const GOAL_FIELDS = new Set(['sales', 'newVisits', 'joins']);

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function isAdminLike(session) {
    return session.role === 'admin' || session.role === 'manager';
}

function validNum(v) {
    const n = Number(v);
    return isFinite(n) && n >= 0 && n <= 1e10 ? Math.round(n) : null;
}

async function shopStaffIds(shopId) {
    const raw = await fetchSalonOne('staffs', { shop_id: shopId });
    const staffs = Array.isArray(raw) ? raw : (raw.data || []);
    return new Set(staffs.filter(s => !s.deleted_at).map(s => String(s.id)));
}

// そのスコープの目標を扱えるか（閲覧・更新共通）
function scopeAllowed(scope, session, allowedStaffIds) {
    if (isAdminLike(session)) return true;
    if (scope === `shop:${session.shopId}`) return true;
    if (session.role === 'store') return scope.startsWith('staff:') && !!allowedStaffIds && allowedStaffIds.has(scope.slice(6));
    if (session.role === 'staff') return scope === `staff:${session.staffId}`;
    return false;
}

function scopeGoals(goals, session, allowedStaffIds) {
    if (isAdminLike(session)) return goals;
    const out = {};
    for (const [month, scopes] of Object.entries(goals || {})) {
        for (const [scope, goal] of Object.entries(scopes || {})) {
            if (!scopeAllowed(scope, session, allowedStaffIds)) continue;
            if (!out[month]) out[month] = {};
            out[month][scope] = goal;
        }
    }
    return out;
}

// patch: { "YYYY-MM": { "<scope>": {sales,newVisits,joins} | null } | null }
// スコープ単位で置き換える（空の目標は削除）。
function applyGoalsPatch(goals, patch, session, allowedStaffIds) {
    for (const [month, scopes] of Object.entries(patch || {})) {
        if (!MONTH_RE.test(month)) throw { code: 'invalid_request', detail: `month: ${month}` };
        if (scopes === null) {
            if (!isAdminLike(session)) throw { code: 'forbidden', detail: '月単位の削除は管理者のみ可能です' };
            delete goals[month];
            continue;
        }
        if (typeof scopes !== 'object') throw { code: 'invalid_request', detail: `month value: ${month}` };
        for (const [scope, goal] of Object.entries(scopes)) {
            if (!SCOPE_RE.test(scope)) throw { code: 'invalid_request', detail: `scope: ${scope}` };
            if (!scopeAllowed(scope, session, allowedStaffIds)) throw { code: 'forbidden', detail: '権限のない目標です' };
            if (!goals[month]) goals[month] = {};
            if (goal === null) {
                delete goals[month][scope];
            } else {
                if (typeof goal !== 'object') throw { code: 'invalid_request', detail: `goal: ${scope}` };
                const cleaned = {};
                for (const [f, v] of Object.entries(goal)) {
                    if (!GOAL_FIELDS.has(f)) throw { code: 'invalid_request', detail: `field: ${f}` };
                    if (v === null || v === '' || v === undefined) continue;
                    const n = validNum(v);
                    if (n === null) throw { code: 'invalid_request', detail: `value: ${f}` };
                    if (n > 0) cleaned[f] = n;
                }
                if (Object.keys(cleaned).length) goals[month][scope] = cleaned;
                else delete goals[month][scope];
            }
            if (Object.keys(goals[month]).length === 0) delete goals[month];
        }
    }
}

// patch: { "<staffId>": 金額 | null }
function applySalariesPatch(salaries, patch) {
    for (const [staffId, v] of Object.entries(patch || {})) {
        if (!/^\d+$/.test(staffId)) throw { code: 'invalid_request', detail: `salary key: ${staffId}` };
        if (v === null || v === '' || v === 0) { delete salaries[staffId]; continue; }
        const n = validNum(v);
        if (n === null) throw { code: 'invalid_request', detail: 'salary value' };
        salaries[staffId] = n;
    }
}

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');

    const session = getSession(req);
    if (!session) return bad(res, 401, 'auth_required');

    try {
        if (!kvAvailable()) {
            if (req.method === 'GET') return res.end(JSON.stringify({ storage: 'none', goals: {}, salaries: {} }));
            return bad(res, 501, 'storage_unconfigured', {
                detail: 'Supabase（または Upstash）のサーバー保存を設定すると全端末で共有保存できます（docs/SALONONE_INTEGRATION.md 参照）',
            });
        }

        const allowed = session.role === 'store' ? await shopStaffIds(session.shopId) : null;

        if (req.method === 'GET') {
            const [goals, salaries] = await Promise.all([
                kvGet(GOALS_KEY),
                session.role === 'admin' ? kvGet(SALARIES_KEY) : Promise.resolve(null),
            ]);
            return res.end(JSON.stringify({
                storage: 'kv',
                goals: scopeGoals(goals || {}, session, allowed),
                salaries: session.role === 'admin' ? (salaries || {}) : undefined,
            }));
        }

        if (req.method !== 'POST') return bad(res, 405, 'method_not_allowed');
        if (session.role === 'staff') return bad(res, 403, 'forbidden', { detail: '目標の変更は店長・マネージャー・オーナーのみ可能です' });

        const body = await readJsonBody(req);
        const patch = body.patch || {};
        let rejected = null;
        let goals = null;
        if (patch.goals && Object.keys(patch.goals).length) {
            goals = await kvUpdate(GOALS_KEY, current => {
                const next = { ...(current || {}) };
                try {
                    applyGoalsPatch(next, patch.goals, session, allowed);
                } catch (e) {
                    if (e.code) { rejected = e; return null; }
                    throw e;
                }
                if (JSON.stringify(next).length > 512 * 1024) { rejected = { code: 'too_large' }; return null; }
                return next;
            });
            if (rejected) return bad(res, rejected.code === 'forbidden' ? 403 : rejected.code === 'too_large' ? 413 : 400, rejected.code, { detail: rejected.detail });
        }
        let salaries = null;
        if (patch.salaries && Object.keys(patch.salaries).length) {
            if (session.role !== 'admin') return bad(res, 403, 'forbidden', { detail: '基本給はオーナーのみ設定できます' });
            salaries = await kvUpdate(SALARIES_KEY, current => {
                const next = { ...(current || {}) };
                try {
                    applySalariesPatch(next, patch.salaries);
                } catch (e) {
                    if (e.code) { rejected = e; return null; }
                    throw e;
                }
                return next;
            });
            if (rejected) return bad(res, 400, rejected.code, { detail: rejected.detail });
        }
        if (goals === null) goals = (await kvGet(GOALS_KEY)) || {};
        if (salaries === null && session.role === 'admin') salaries = (await kvGet(SALARIES_KEY)) || {};
        return res.end(JSON.stringify({
            ok: true,
            storage: 'kv',
            goals: scopeGoals(goals, session, allowed),
            salaries: session.role === 'admin' ? (salaries || {}) : undefined,
        }));
    } catch (e) {
        console.error('goals api error', e);
        return bad(res, 500, 'internal_error');
    }
};
