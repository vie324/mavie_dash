// GET /api/auth/session?store=&staff=
// 現在のセッション状態と、URLコンテキストに対して必要な認証を返す。
// 従来ツールのフェイルオープン挙動を踏襲: パスワード未設定の対象は自動ログイン。

'use strict';

const {
    getSession, setSessionCookie, resolveContext,
    requiredStaffPassword, adminPassword, isDemo,
} = require('../_lib/auth');

function sessionMatches(session, ctx) {
    if (!session) return false;
    if (session.role === 'admin') return true; // 管理者は全コンテキスト閲覧可
    if (ctx.role === 'store') {
        return (session.role === 'store' || session.role === 'staff') && String(session.shopId) === String(ctx.shopId);
    }
    if (ctx.role === 'staff') {
        if (session.role === 'staff') {
            return String(session.shopId) === String(ctx.shopId) && String(session.staffId) === String(ctx.staffId);
        }
        return false;
    }
    return session.role === 'admin';
}

module.exports = async (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    if (req.method !== 'GET') {
        res.statusCode = 405;
        return res.end(JSON.stringify({ error: 'method_not_allowed' }));
    }
    try {
        const url = new URL(req.url, 'http://localhost');
        const storeParam = url.searchParams.get('store') || '';
        const staffParam = url.searchParams.get('staff') || '';

        const session = getSession(req);
        let ctx;
        try {
            ctx = await resolveContext(storeParam, staffParam);
        } catch (e) {
            // 上流API未達でも管理画面は開けるようにする（データ取得時に別途エラー表示）
            ctx = storeParam ? { role: 'invalid', reason: 'resolve_failed' } : { role: 'admin' };
        }

        const body = {
            demo: isDemo(),
            context: {
                role: ctx.role,
                shopId: ctx.shopId ?? null,
                shopName: ctx.shopName ?? null,
                staffId: ctx.staffId ?? null,
                staffName: ctx.staffName ?? null,
                reason: ctx.reason ?? null,
            },
        };

        if (ctx.role === 'invalid') {
            res.statusCode = 200;
            return res.end(JSON.stringify({ ...body, authenticated: false, needsPassword: false }));
        }

        if (sessionMatches(session, ctx)) {
            res.statusCode = 200;
            return res.end(JSON.stringify({
                ...body,
                authenticated: true,
                session: { role: session.role, shopId: session.shopId ?? null, shopName: session.shopName ?? null, staffId: session.staffId ?? null, staffName: session.staffName ?? null },
            }));
        }

        // 未認証: パスワード不要なら自動発行（従来挙動の踏襲）
        let needsPassword;
        if (ctx.role === 'admin') {
            needsPassword = adminPassword() !== '';
        } else if (ctx.role === 'staff') {
            needsPassword = requiredStaffPassword(ctx) !== '';
        } else {
            needsPassword = false; // 店舗ビューは従来どおりパスワードなし
        }

        if (!needsPassword) {
            const payload = ctx.role === 'admin'
                ? { role: 'admin' }
                : { role: ctx.role, shopId: ctx.shopId, shopName: ctx.shopName, staffId: ctx.staffId ?? null, staffName: ctx.staffName ?? null };
            setSessionCookie(res, payload);
            res.statusCode = 200;
            return res.end(JSON.stringify({ ...body, authenticated: true, session: payload, autoIssued: true }));
        }

        res.statusCode = 200;
        return res.end(JSON.stringify({ ...body, authenticated: false, needsPassword: true }));
    } catch (e) {
        console.error('session error', e);
        res.statusCode = 500;
        return res.end(JSON.stringify({ error: 'internal_error' }));
    }
};
