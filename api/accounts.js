// /api/accounts — スタッフ・店長アカウント（パスワード）の発行・管理
// 設定タブから発行でき、Vercelの環境変数を触らずに済む。
// 保存: vie:accounts → { "staff:<staffId>": { salt, hash, updatedAt }, "store:<shopId>": {...} }
// パスワードはscryptでハッシュ化して保存（平文は保持しない）。
// 権限: オーナー/マネージャー = 全て、店長 = 自店舗スタッフのみ、スタッフ = 不可

'use strict';

const crypto = require('crypto');
const { getSession, readJsonBody, invalidateAccountsCache } = require('./_lib/auth');
const { kvAvailable, kvGet, kvSet } = require('./_lib/kv');
const { fetchSalonOne } = require('./_lib/salonone');

const ACCOUNTS_KEY = 'vie:accounts';

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

async function staffShopMap() {
    const raw = await fetchSalonOne('staffs', {});
    const staffs = Array.isArray(raw) ? raw : (raw.data || []);
    const map = {};
    for (const s of staffs) if (!s.deleted_at) map[String(s.id)] = String(s.shop_id);
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
            detail: 'アカウント発行にはサーバー保存が必要です。Vercelで Upstash for Redis 連携を追加してください',
        });

        const body = await readJsonBody(req);
        const kind = body.kind === 'store' ? 'store' : body.kind === 'staff' ? 'staff' : null;
        const id = String(body.id || '');
        if (!kind || !/^\d+$/.test(id)) return bad(res, 400, 'invalid_request', { fields: ['kind', 'id'] });
        const key = `${kind}:${id}`;
        if (allowedKey && !allowedKey(key)) return bad(res, 403, 'forbidden', { detail: '他店舗のアカウントは操作できません' });
        if (kind === 'store' && !isAdminLike(session)) return bad(res, 403, 'forbidden');

        const accounts = (await kvGet(ACCOUNTS_KEY)) || {};

        if (body.action === 'set') {
            const password = String(body.password || '');
            if (password.length < 4 || password.length > 64) {
                return bad(res, 400, 'invalid_request', { fields: ['password'], detail: 'パスワードは4〜64文字で設定してください' });
            }
            const salt = crypto.randomBytes(16).toString('hex');
            accounts[key] = { salt, hash: hashPassword(password, salt), updatedAt: new Date().toISOString() };
            await kvSet(ACCOUNTS_KEY, accounts);
            invalidateAccountsCache();
            return res.end(JSON.stringify({ ok: true, accounts: publicList(accounts, allowedKey) }));
        }

        if (body.action === 'delete') {
            delete accounts[key];
            await kvSet(ACCOUNTS_KEY, accounts);
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
