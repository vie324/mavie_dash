// POST /api/auth/logout — セッションクッキーを破棄

'use strict';

const { clearSessionCookie } = require('../_lib/auth');

module.exports = async (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    if (req.method !== 'POST') {
        res.statusCode = 405;
        return res.end(JSON.stringify({ error: 'method_not_allowed' }));
    }
    clearSessionCookie(res);
    res.statusCode = 200;
    return res.end(JSON.stringify({ ok: true }));
};
