// /api/accounts — スタッフ・店長アカウント（パスワード）の発行・管理
// 設定タブから発行でき、Vercelの環境変数を触らずに済む。
// 保存: vie:accounts → { "staff:<staffId>": { salt, hash, updatedAt }, "store:<shopId>": {...} }
// パスワードはscryptでハッシュ化して保存（平文は保持しない）。
// 権限: オーナー/マネージャー = 全て、店長 = 自店舗スタッフのみ、スタッフ = 不可
//
// action:
//   set    … 1件のパスワードを設定（{kind, id, password}）
//   delete … 1件のパスワードを解除（{kind, id}）
//   bulk   … まだアカウントがないスタッフ全員に、パスワードを自動で発行（{shopId?}）。
//            発行したパスワード（平文）はこの応答でだけ返す（サーバーには残らない）

'use strict';

const crypto = require('crypto');
const { promisify } = require('util');
const { getSession, readJsonBody, invalidateAccountsCache } = require('./_lib/auth');
const { kvAvailable, kvGet, kvUpdate } = require('./_lib/kv');
const { fetchSalonOne } = require('./_lib/salonone');

const ACCOUNTS_KEY = 'vie:accounts';
const scryptAsync = promisify(crypto.scrypt);
// 自動発行のパスワード: 見間違えやすい文字（0/o/1/l/i）を除いた英小文字+数字 8文字（スマホで打ちやすいよう記号なし）
const PASS_CHARS = 'abcdefghjkmnpqrstuvwxyz23456789';
const PASS_LENGTH = 8;

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

function isAdminLike(session) {
    return session.role === 'admin' || session.role === 'manager';
}

function hashPassword(password, salt) {
    return crypto.scryptSync(String(password), salt, 64).toString('hex');
}

function generatePassword() {
    let out = '';
    for (let i = 0; i < PASS_LENGTH; i++) out += PASS_CHARS[crypto.randomInt(PASS_CHARS.length)];
    return out;
}

async function newEntry(password) {
    const salt = crypto.randomBytes(16).toString('hex');
    const hash = (await scryptAsync(String(password), salt, 64)).toString('hex');
    return { salt, hash, updatedAt: new Date().toISOString() };
}

async function activeStaffs() {
    const raw = await fetchSalonOne('staffs', {});
    const staffs = Array.isArray(raw) ? raw : (raw.data || []);
    return staffs.filter(s => !s.deleted_at && s.id !== undefined && s.id !== null);
}

async function staffShopMap() {
    const map = {};
    for (const s of await activeStaffs()) map[String(s.id)] = String(s.shop_id);
    return map;
}

// ハッシュを含まない公開用の一覧
function publicList(accounts, filter) {
    const out = {};
    for (const [key, a] of Object.entries(accounts || {})) {
        if (filter && !filter(key)) continue;
        out[key] = { updatedAt: a.updatedAt || null };
    }
    return out;
}

// 未発行のスタッフに一括発行。ハッシュ計算（時間がかかる）はロックの外で済ませ、ロック中は書き込むだけにする
async function bulkIssue(session, body, allowedKey) {
    const shopFilter = body.shopId === undefined || body.shopId === null || body.shopId === '' || body.shopId === 'all' ? null : String(body.shopId);
    if (shopFilter !== null && !/^\d+$/.test(shopFilter)) throw { status: 400, code: 'invalid_request', detail: 'shopId' };
    let targets = await activeStaffs();
    if (session.role === 'store') targets = targets.filter(s => String(s.shop_id) === String(session.shopId));
    if (shopFilter !== null) targets = targets.filter(s => String(s.shop_id) === shopFilter);

    const before = (await kvGet(ACCOUNTS_KEY)) || {};
    const prepared = await Promise.all(targets
        .filter(s => !before[`staff:${s.id}`])
        .map(async s => {
            const password = generatePassword();
            return { staff: s, password, entry: await newEntry(password) };
        }));

    const issued = [];
    const saved = await kvUpdate(ACCOUNTS_KEY, current => {
        const accounts = { ...(current || {}) };
        issued.length = 0;
        for (const p of prepared) {
            const key = `staff:${p.staff.id}`;
            if (accounts[key]) continue; // 同時に別の画面から発行された分は上書きしない
            accounts[key] = p.entry;
            issued.push({ kind: 'staff', id: String(p.staff.id), shopId: String(p.staff.shop_id), name: p.staff.name || '', password: p.password });
        }
        return issued.length ? accounts : null;
    });
    if (issued.length) invalidateAccountsCache();
    const accounts = saved || (await kvGet(ACCOUNTS_KEY)) || {};
    return { ok: true, issued, skipped: targets.length - issued.length, accounts: publicList(accounts, allowedKey) };
}

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');

    const session = getSession(req);
    if (!session) return bad(res, 401, 'auth_required');
    if (session.role === 'staff') return bad(res, 403, 'forbidden');

    try {
        // 店長は自店舗スタッフのアカウントだけ扱える
        let allowedKey = null;
        if (session.role === 'store') {
            const map = await staffShopMap();
            const own = String(session.shopId);
            allowedKey = key => key.startsWith('staff:') && map[key.slice(6)] === own;
        }

        if (req.method === 'GET') {
            if (!kvAvailable()) return res.end(JSON.stringify({ storage: 'none', accounts: {} }));
            const accounts = (await kvGet(ACCOUNTS_KEY)) || {};
            return res.end(JSON.stringify({ storage: 'kv', accounts: publicList(accounts, allowedKey) }));
        }

        if (req.method !== 'POST') return bad(res, 405, 'method_not_allowed');
        if (!kvAvailable()) return bad(res, 501, 'storage_unconfigured', {
            detail: 'アカウント発行にはサーバー保存が必要です。Supabase（または Upstash）のサーバー保存を設定してください',
        });

        const body = await readJsonBody(req);

        if (body.action === 'bulk') {
            try {
                return res.end(JSON.stringify(await bulkIssue(session, body, allowedKey)));
            } catch (e) {
                if (e && e.code && e.status) return bad(res, e.status, e.code, { detail: e.detail });
                throw e;
            }
        }

        const kind = body.kind === 'store' ? 'store' : body.kind === 'staff' ? 'staff' : null;
        const id = String(body.id || '');
        if (!kind || !/^\d+$/.test(id)) return bad(res, 400, 'invalid_request', { fields: ['kind', 'id'] });
        const key = `${kind}:${id}`;
        if (allowedKey && !allowedKey(key)) return bad(res, 403, 'forbidden', { detail: '他店舗のアカウントは操作できません' });
        if (kind === 'store' && !isAdminLike(session)) return bad(res, 403, 'forbidden');

        if (body.action === 'set') {
            const password = String(body.password || '');
            if (password.length < 4 || password.length > 64) {
                return bad(res, 400, 'invalid_request', { fields: ['password'], detail: 'パスワードは4〜64文字で設定してください' });
            }
            const entry = await newEntry(password);
            const accounts = await kvUpdate(ACCOUNTS_KEY, current => ({ ...(current || {}), [key]: entry }));
            invalidateAccountsCache();
            return res.end(JSON.stringify({ ok: true, accounts: publicList(accounts, allowedKey) }));
        }

        if (body.action === 'delete') {
            const accounts = await kvUpdate(ACCOUNTS_KEY, current => {
                const next = { ...(current || {}) };
                delete next[key];
                return next;
            });
            invalidateAccountsCache();
            return res.end(JSON.stringify({ ok: true, accounts: publicList(accounts, allowedKey) }));
        }

        return bad(res, 400, 'invalid_request', { fields: ['action'] });
    } catch (e) {
        console.error('accounts api error', e);
        return bad(res, 500, 'internal_error');
    }
};

module.exports.ACCOUNTS_KEY = ACCOUNTS_KEY;
module.exports.hashPassword = hashPassword;
module.exports._internal = { generatePassword, PASS_CHARS, PASS_LENGTH };
