// POST /api/ai — Gemini APIのサーバー側プロキシ
// キーは GEMINI_API_KEY 環境変数のみ（従来のlocalStorage保存を廃止し、ブラウザに露出させない）。

'use strict';

const { getSession, readJsonBody } = require('./_lib/auth');

const MODEL = process.env.GEMINI_MODEL || 'gemini-2.0-flash';
const MAX_PROMPT = 16000;

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');
    if (req.method !== 'POST') {
        res.statusCode = 405;
        return res.end(JSON.stringify({ error: 'method_not_allowed' }));
    }
    if (!getSession(req)) {
        res.statusCode = 401;
        return res.end(JSON.stringify({ error: 'auth_required' }));
    }
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
        res.statusCode = 503;
        return res.end(JSON.stringify({ error: 'ai_unavailable' }));
    }
    try {
        const { prompt } = await readJsonBody(req);
        if (!prompt || typeof prompt !== 'string') {
            res.statusCode = 400;
            return res.end(JSON.stringify({ error: 'invalid_request', fields: ['prompt'] }));
        }
        const upstream = await fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${MODEL}:generateContent?key=${key}`,
            {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    contents: [{ parts: [{ text: prompt.slice(0, MAX_PROMPT) }] }],
                    generationConfig: { temperature: 0.6, maxOutputTokens: 1024 },
                }),
            }
        );
        if (!upstream.ok) {
            console.error('gemini error', upstream.status);
            res.statusCode = 502;
            return res.end(JSON.stringify({ error: 'ai_upstream_error' }));
        }
        const json = await upstream.json();
        const text = json.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '';
        res.statusCode = 200;
        return res.end(JSON.stringify({ text }));
    } catch (e) {
        console.error('ai proxy error', e);
        res.statusCode = 500;
        return res.end(JSON.stringify({ error: 'internal_error' }));
    }
};
