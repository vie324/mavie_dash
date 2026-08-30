// POST /api/auth/login  { store?, staff?, mode?, password? }
// パスワードを照合してセッションクッキーを発行する。

'use strict';

const {
    setSessionCookie, resolveContext, requiredPasswordFor,
    timingSafeEq, readJsonBody,
} = require('../_lib/auth');

module.exports = async (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    if (req.method !== 'POST') {
        res.statusCode = 405;
        return res.end(JSON.stringify({ error: 'method_not_allowed' }));
    }
    try {
        const { store = '', staff = '', mode = '', password = '' } = await readJsonBody(req);
        let ctx;
        try {
            ctx = await resolveContext(store, staff, mode);
        } catch (e) {
            res.statusCode = 502;
            return res.end(JSON.stringify({ error: 'upstream_unreachable' }));
        }
        if (ctx.role === 'invalid') {
            res.statusCode = 404;
            return res.end(JSON.stringify({ error: 'unknown_target', reason: ctx.reason }));
        }

        const required = requiredPasswordFor(ctx);
        if (required !== '' && !timingSafeEq(password, required)) {
            res.statusCode = 401;
            return res.end(JSON.stringify({ error: 'invalid_password' }));
        }

        const payload = (ctx.role === 'admin' || ctx.role === 'manager')
            ? { role: ctx.role }
            : { role: ctx.role, shopId: ctx.shopId, shopName: ctx.shopName, staffId: ctx.staffId ?? null, staffName: ctx.staffName ?? null };
        setSessionCookie(res, payload);
        res.statusCode = 200;
        return res.end(JSON.stringify({ ok: true, session: payload }));
    } catch (e) {
        console.error('login error', e);
        res.statusCode = 500;
        return res.end(JSON.stringify({ error: 'internal_error' }));
    }
};
