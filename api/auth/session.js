// GET /api/auth/session?store=&staff=&mode=
// 現在のセッション状態と、URLコンテキストに対して必要な認証を返す。
// 役割は4段階: admin(オーナー) > manager(マネージャー) > store(店長) > staff(スタッフ)。
// パスワード未設定の対象は自動ログイン（従来ツールのフェイルオープン挙動を踏襲。設定タブに警告表示あり）。

'use strict';

const {
    getSession, setSessionCookie, resolveContext,
    requiredAuthFor, isDemo,
} = require('../_lib/auth');

// 上位ロールのセッションは下位コンテキストをそのまま閲覧できる
// （例: 店長が自店スタッフの専用URLを開いても再ログイン不要）
function sessionMatches(session, ctx) {
    if (!session) return false;
    if (session.role === 'admin') return true;
    if (session.role === 'manager') return ctx.role !== 'admin';
    if (session.role === 'store') {
        if (ctx.role === 'store' || ctx.role === 'staff') {
            return String(session.shopId) === String(ctx.shopId);
        }
        return false;
    }
    if (session.role === 'staff') {
        return ctx.role === 'staff'
            && String(session.shopId) === String(ctx.shopId)
            && String(session.staffId) === String(ctx.staffId);
    }
    return false;
}

function sessionPayload(ctx) {
    if (ctx.role === 'admin' || ctx.role === 'manager') return { role: ctx.role };
    return { role: ctx.role, shopId: ctx.shopId, shopName: ctx.shopName, staffId: ctx.staffId ?? null, staffName: ctx.staffName ?? null };
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
        const modeParam = url.searchParams.get('mode') || '';

        const session = getSession(req);
        let ctx;
        try {
            ctx = await resolveContext(storeParam, staffParam, modeParam);
        } catch (e) {
            // 上流API未達でも管理画面は開けるようにする（データ取得時に別途エラー表示）
            ctx = storeParam ? { role: 'invalid', reason: 'resolve_failed' } : { role: modeParam === 'manager' ? 'manager' : 'admin' };
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
        const needsPassword = (await requiredAuthFor(ctx)).type !== 'none';
        if (!needsPassword) {
            const payload = sessionPayload(ctx);
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
