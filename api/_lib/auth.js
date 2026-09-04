// セッション認証（サーバー側で強制する権限モデル）
// 役割は従来ツールと同じ3種:
//   admin  … 全店舗・全機能
//   store  … ?store=XXX でロックされた店舗ビュー
//   staff  … ?store=XXX&staff=YYY でロックされた個人ビュー（パスワードは STAFF_PASSWORDS で設定）
// 従来との互換: パスワード未設定の対象はそのまま入れる（フェイルオープン挙動を踏襲）。
// 従来と違い、セッションはHMAC署名付きクッキーで、データAPI側でも店舗スコープを強制する。

'use strict';

const crypto = require('crypto');
const { fetchSalonOne, isDemo } = require('./salonone');
const { kvAvailable, kvGet } = require('./kv');

const COOKIE_NAME = 'vie_session';
const SESSION_TTL_SEC = 24 * 3600;

function secret() {
    if (process.env.AUTH_SECRET) return process.env.AUTH_SECRET;
    if (process.env.SALONONE_API_KEY) {
        return crypto.createHash('sha256').update('vie-dash:' + process.env.SALONONE_API_KEY).digest('hex');
    }
    return 'vie-dash-demo-insecure-secret';
}

function b64url(buf) {
    return Buffer.from(buf).toString('base64').replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function unb64url(s) {
    return Buffer.from(s.replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString();
}

function signSession(payload) {
    const body = b64url(JSON.stringify({ ...payload, exp: Math.floor(Date.now() / 1000) + SESSION_TTL_SEC }));
    const mac = b64url(crypto.createHmac('sha256', secret()).update(body).digest());
    return `${body}.${mac}`;
}

function verifyToken(token) {
    if (!token || typeof token !== 'string') return null;
    const dot = token.lastIndexOf('.');
    if (dot < 0) return null;
    const body = token.slice(0, dot);
    const mac = token.slice(dot + 1);
    const expect = b64url(crypto.createHmac('sha256', secret()).update(body).digest());
    const a = Buffer.from(mac), b = Buffer.from(expect);
    if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
    try {
        const payload = JSON.parse(unb64url(body));
        if (!payload.exp || payload.exp < Math.floor(Date.now() / 1000)) return null;
        return payload;
    } catch (_) {
        return null;
    }
}

function parseCookies(req) {
    const header = req.headers.cookie || '';
    const out = {};
    for (const part of header.split(';')) {
        const i = part.indexOf('=');
        if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
    }
    return out;
}

function getSession(req) {
    return verifyToken(parseCookies(req)[COOKIE_NAME]);
}

function setSessionCookie(res, payload) {
    const token = signSession(payload);
    const secure = process.env.NODE_ENV !== 'development' && !process.env.DEV_SERVER;
    res.setHeader('Set-Cookie',
        `${COOKIE_NAME}=${encodeURIComponent(token)}; Path=/; Max-Age=${SESSION_TTL_SEC}; HttpOnly; SameSite=Lax${secure ? '; Secure' : ''}`);
}

function clearSessionCookie(res) {
    res.setHeader('Set-Cookie', `${COOKIE_NAME}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax`);
}

// ---- パスワード設定 ----
function staffPasswordMap() {
    try {
        return JSON.parse(process.env.STAFF_PASSWORDS || '{}');
    } catch (_) {
        console.error('STAFF_PASSWORDS is not valid JSON');
        return {};
    }
}

// 設定済みパスワードの照合。
// 重要: URLの指定形式（旧スラッグ / ID / 名前）が違っても同じスタッフには同じパスワードが
// かかるよう、キーの形式に依存せず「解決済みの店舗・スタッフに一致するエントリ」を探す。
// （形式一致だけで探すと、別形式のURLでパスワードを素通りできてしまう）
const LEGACY_SLUGS = { chiba: '千葉', honatsugi: '本厚木', yamato: '大和' };

function storePartMatches(part, ctx) {
    if (!part) return false;
    const p = part.toLowerCase();
    if (p === String(ctx.shopId)) return true;
    const shopName = String(ctx.shopName || '').toLowerCase();
    if (shopName && (p === shopName || shopName.includes(p))) return true;
    if ((ctx.storeParam || '').toLowerCase() === p) return true;
    const slugMap = storeSlugMap();
    if (slugMap[p] !== undefined && String(slugMap[p]) === String(ctx.shopId)) return true;
    if (LEGACY_SLUGS[p] && String(ctx.shopName || '').includes(LEGACY_SLUGS[p])) return true;
    return false;
}

function staffPartMatches(part, ctx) {
    if (!part) return false;
    const p = part.toLowerCase();
    if (p === String(ctx.staffId)) return true;
    const staffName = String(ctx.staffName || '').toLowerCase();
    if (staffName && (p === staffName || staffName.includes(p))) return true;
    if ((ctx.staffParam || '').toLowerCase() === p) return true;
    return false;
}

function requiredStaffPassword(ctx) {
    const map = staffPasswordMap();
    for (const [key, pass] of Object.entries(map)) {
        const k = String(key);
        const sep = k.indexOf('_');
        if (sep > 0) {
            const storePart = k.slice(0, sep);
            const staffPart = k.slice(sep + 1);
            if (storePartMatches(storePart, ctx) && staffPartMatches(staffPart, ctx)) return pass || '';
        } else if (staffPartMatches(k, ctx)) {
            // 店舗指定なしのキー（スタッフ名 or ID のみ）
            return pass || '';
        }
    }
    return '';
}

function timingSafeEq(a, b) {
    const ba = Buffer.from(String(a)), bb = Buffer.from(String(b));
    if (ba.length !== bb.length) return false;
    return crypto.timingSafeEqual(ba, bb);
}

// ---- 店舗・スタッフの解決（レガシーURL ?store=chiba&staff=kiki 互換）----
function storeSlugMap() {
    try {
        return JSON.parse(process.env.STORE_SLUGS || '{}');
    } catch (_) {
        return {};
    }
}

function norm(s) {
    return String(s || '').trim().toLowerCase();
}

async function resolveContext(storeParam, staffParam, modeParam) {
    // ?mode=manager → マネージャー（全店舗閲覧・給与と設定以外）
    if (String(modeParam || '').toLowerCase() === 'manager') return { role: 'manager' };
    if (!storeParam) return { role: 'admin' };
    const shopsRes = await fetchSalonOne('shops', {});
    const shops = (Array.isArray(shopsRes) ? shopsRes : (shopsRes?.data || [])).filter(s => !s.deleted_at);
    const slugMap = storeSlugMap();
    const want = norm(storeParam);

    let shop = null;
    if (slugMap[want] !== undefined) {
        shop = shops.find(s => String(s.id) === String(slugMap[want]));
    }
    if (!shop && /^\d+$/.test(want)) {
        shop = shops.find(s => String(s.id) === want);
    }
    if (!shop) {
        shop = shops.find(s => norm(s.name) === want)
            || shops.find(s => norm(s.name).includes(want));
    }
    // レガシーslugの既定マッピング（ローマ字slug → 店名に含まれる漢字）
    if (!shop) {
        const legacy = { chiba: '千葉', honatsugi: '本厚木', yamato: '大和' };
        if (legacy[want]) shop = shops.find(s => String(s.name).includes(legacy[want]));
    }
    if (!shop) return { role: 'invalid', reason: 'unknown_store', storeParam };

    if (!staffParam) {
        return { role: 'store', shopId: shop.id, shopName: shop.name, storeParam };
    }

    const staffsRes = await fetchSalonOne('staffs', { shop_id: shop.id });
    const staffs = (Array.isArray(staffsRes) ? staffsRes : (staffsRes?.data || [])).filter(s => !s.deleted_at);
    const wantStaff = norm(staffParam);
    let staff = null;
    if (/^\d+$/.test(wantStaff)) staff = staffs.find(s => String(s.id) === wantStaff);
    if (!staff) staff = staffs.find(s => norm(s.name) === wantStaff) || staffs.find(s => norm(s.name).includes(wantStaff));
    if (!staff) return { role: 'invalid', reason: 'unknown_staff', storeParam, staffParam };

    return {
        role: 'staff',
        shopId: shop.id, shopName: shop.name,
        staffId: staff.id, staffName: staff.name,
        storeParam, staffParam,
    };
}

function adminPassword() {
    return process.env.ADMIN_PASSWORD || '';
}

function managerPassword() {
    return process.env.MANAGER_PASSWORD || '';
}

// 店長ビュー(?store=X)のパスワード。STORE_PASSWORDS のキーは
// スラッグ / shop_id / 店名 のどれでも可（スタッフ同様、形式に依存せず照合する）。
function storePasswordMap() {
    try {
        return JSON.parse(process.env.STORE_PASSWORDS || '{}');
    } catch (_) {
        console.error('STORE_PASSWORDS is not valid JSON');
        return {};
    }
}

function requiredStorePassword(ctx) {
    const map = storePasswordMap();
    for (const [key, pass] of Object.entries(map)) {
        if (storePartMatches(String(key), ctx)) return pass || '';
    }
    return '';
}

// コンテキストに応じた必要パスワード（未設定=''はパスワード不要＝従来のフェイルオープン挙動）
function requiredPasswordFor(ctx) {
    if (ctx.role === 'admin') return adminPassword();
    if (ctx.role === 'manager') return managerPassword();
    if (ctx.role === 'store') return requiredStorePassword(ctx);
    if (ctx.role === 'staff') return requiredStaffPassword(ctx);
    return '';
}

// ---- 設定タブから発行したアカウント（KV・scryptハッシュ）----
// 環境変数のパスワードより優先する。60秒メモリキャッシュ。
let accountsCache = { at: 0, data: null };
async function loadAccounts() {
    if (!kvAvailable()) return {};
    if (Date.now() - accountsCache.at < 60000 && accountsCache.data) return accountsCache.data;
    try {
        const data = (await kvGet('vie:accounts')) || {};
        accountsCache = { at: Date.now(), data };
        return data;
    } catch (e) {
        console.error('accounts load failed', e);
        return accountsCache.data || {};
    }
}

function invalidateAccountsCache() {
    accountsCache = { at: 0, data: null };
}

// 必要な認証を返す: {type:'none'} | {type:'plain', password} | {type:'hash', salt, hash}
async function requiredAuthFor(ctx) {
    if (ctx.role === 'staff' || ctx.role === 'store') {
        const accounts = await loadAccounts();
        const key = ctx.role === 'staff' ? `staff:${ctx.staffId}` : `store:${ctx.shopId}`;
        const acc = accounts[key];
        if (acc && acc.hash && acc.salt) return { type: 'hash', salt: acc.salt, hash: acc.hash };
    }
    const plain = requiredPasswordFor(ctx);
    return plain === '' ? { type: 'none' } : { type: 'plain', password: plain };
}

function verifyCredential(auth, password) {
    if (!auth || auth.type === 'none') return true;
    if (auth.type === 'plain') return timingSafeEq(password, auth.password);
    if (auth.type === 'hash') {
        const got = crypto.scryptSync(String(password || ''), auth.salt, 64).toString('hex');
        return timingSafeEq(got, auth.hash);
    }
    return false;
}

// 設定タブでの警告表示用（設定有無のみ。値は返さない）
function passwordConfigStatus() {
    return {
        admin: adminPassword() !== '',
        manager: managerPassword() !== '',
        storeCount: Object.keys(storePasswordMap()).length,
        staffCount: Object.keys(staffPasswordMap()).length,
    };
}

async function accountsSummary() {
    const accounts = await loadAccounts();
    const keys = Object.keys(accounts);
    return {
        staffAccounts: keys.filter(k => k.startsWith('staff:')).length,
        storeAccounts: keys.filter(k => k.startsWith('store:')).length,
    };
}

function readJsonBody(req) {
    return new Promise((resolve, reject) => {
        if (req.body !== undefined) {
            // Vercelは既にパース済みの場合がある
            if (typeof req.body === 'string') {
                try { return resolve(JSON.parse(req.body || '{}')); } catch (e) { return resolve({}); }
            }
            return resolve(req.body || {});
        }
        let raw = '';
        req.on('data', c => { raw += c; if (raw.length > 64 * 1024) { reject(new Error('body too large')); req.destroy(); } });
        req.on('end', () => {
            try { resolve(raw ? JSON.parse(raw) : {}); } catch (e) { resolve({}); }
        });
        req.on('error', reject);
    });
}

module.exports = {
    COOKIE_NAME,
    getSession, setSessionCookie, clearSessionCookie,
    resolveContext, requiredStaffPassword, adminPassword,
    requiredPasswordFor, passwordConfigStatus,
    requiredAuthFor, verifyCredential, accountsSummary, invalidateAccountsCache,
    timingSafeEq, readJsonBody, isDemo,
};
