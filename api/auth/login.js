// POST /api/auth/login  { store?, staff?, password? }
// パスワードを照合してセッションクッキーを発行する。

'use strict';

const {
    setSessionCookie, resolveContext, requiredStaffPassword,
    adminPassword, timingSafeEq, readJsonBody,
} = require('../_lib/auth');

module.exports = async (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    if (req.method !== 'POST') {
        res.statusCode = 405;
        return res.end(JSON.stringify({ error: 'method_not_allowed' }));
    }
    try {
        const { store = '', staff = '', password = '' } = await readJsonBody(req);
        let ctx;
        try {
            ctx = await resolveContext(store, staff);
        } catch (e) {
            res.statusCode = 502;
            return res.end(JSON.stringify({ error: 'upstream_unreachable' }));
        }
        if (ctx.role === 'invalid') {
            res.statusCode = 404;
            return res.end(JSON.stringify({ error: 'unknown_target', reason: ctx.reason }));
        }

        let required = '';
        if (ctx.role === 'admin') required = adminPassword();
        else if (ctx.role === 'staff') required = requiredStaffPassword(ctx);

        if (required !== '' && !timingSafeEq(password, required)) {
            res.statusCode = 401;
            return res.end(JSON.stringify({ error: 'invalid_password' }));
        }

        const payload = ctx.role === 'admin'
            ? { role: 'admin' }
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
