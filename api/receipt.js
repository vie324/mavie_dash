// /api/receipt — 出納帳の領収書・レシートの写真
//   POST ?shop=<id>&month=YYYY-MM&type=image/jpeg[&name=...]   本文に画像（application/octet-stream・4MBまで）
//        → { id, path, type, size, name }  このあと entry.add / entry.edit の receipts に渡す
//   GET  ?shop=<id>&path=<保存先パス>                           → 署名付きURL（10分）へリダイレクト
// 権限は出納帳と同じ（オーナー/マネージャー = 全店舗、店長 = 自店舗、スタッフ = 不可）。
// 保存先: Supabase Storage の非公開バケット receipts/<店舗>/<月>/<id>.<拡張子>

'use strict';

const crypto = require('crypto');
const { getSession } = require('./_lib/auth');
const { canAccessShop, MONTH_RE } = require('./_lib/cashbook');
const S = require('./_lib/storage');

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.end(JSON.stringify({ error, ...extra }));
}

// 店舗の領収書パスか（他店舗のパスを指定されても見えないようにする）
function pathForShop(path, shopId) {
    const m = S.RECEIPT_PATH_RE.exec(String(path || ''));
    return !!m && m[1] === String(shopId);
}

function readRawBody(req, limit) {
    return new Promise((resolve, reject) => {
        if (req.body !== undefined && req.body !== null) {
            if (Buffer.isBuffer(req.body)) return resolve(req.body);
            if (typeof req.body === 'string') return resolve(Buffer.from(req.body, 'binary'));
            return reject(new Error('unexpected body'));
        }
        const chunks = [];
        let size = 0;
        req.on('data', c => {
            size += c.length;
            if (size > limit) { reject(new Error('too_large')); req.destroy(); return; }
            chunks.push(c);
        });
        req.on('end', () => resolve(Buffer.concat(chunks)));
        req.on('error', reject);
    });
}

// 画像の中身で種類を判定（Content-Type の申告だけを信用しない）
function sniff(buf) {
    if (buf.length > 3 && buf[0] === 0xff && buf[1] === 0xd8 && buf[2] === 0xff) return 'image/jpeg';
    if (buf.length > 8 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4e && buf[3] === 0x47) return 'image/png';
    if (buf.length > 12 && buf.slice(0, 4).toString('ascii') === 'RIFF' && buf.slice(8, 12).toString('ascii') === 'WEBP') return 'image/webp';
    return null;
}

module.exports = async (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    const session = getSession(req);
    if (!session) return bad(res, 401, 'auth_required');
    if (session.role === 'staff') return bad(res, 403, 'forbidden');
    if (!S.storageAvailable()) return bad(res, 501, 'storage_unconfigured', { detail: '領収書の写真の保存には Supabase が必要です' });

    const url = new URL(req.url, 'http://localhost');
    const shopId = url.searchParams.get('shop') || '';
    if (!/^\d+$/.test(shopId)) return bad(res, 400, 'invalid_request', { fields: ['shop'] });
    if (!canAccessShop(session, shopId)) return bad(res, 403, 'forbidden');

    try {
        if (req.method === 'GET') {
            const path = url.searchParams.get('path') || '';
            if (!pathForShop(path, shopId)) return bad(res, 400, 'invalid_request', { fields: ['path'] });
            const signed = await S.signedUrl(path, 600);
            if (!signed) return bad(res, 404, 'not_found');
            res.statusCode = 302;
            res.setHeader('Location', signed);
            return res.end();
        }
        if (req.method !== 'POST') return bad(res, 405, 'method_not_allowed');

        const month = url.searchParams.get('month') || '';
        if (!MONTH_RE.test(month)) return bad(res, 400, 'invalid_request', { fields: ['month'] });
        const declared = (req.headers['content-length'] && Number(req.headers['content-length'])) || 0;
        if (declared > S.MAX_BYTES) return bad(res, 413, 'too_large', { detail: '画像は4MBまでです' });
        let buf;
        try {
            buf = await readRawBody(req, S.MAX_BYTES);
        } catch (e) {
            if (e.message === 'too_large') return bad(res, 413, 'too_large', { detail: '画像は4MBまでです' });
            throw e;
        }
        if (!buf.length) return bad(res, 400, 'invalid_request', { detail: '画像が空です' });
        const type = sniff(buf);
        if (!type || !S.ALLOWED[type]) return bad(res, 415, 'unsupported_type', { detail: 'JPEG / PNG / WebP の画像のみ保存できます' });
        const id = 'r_' + crypto.randomBytes(8).toString('hex');
        const path = `receipts/${shopId}/${month}/${id}.${S.ALLOWED[type]}`;
        await S.putObject(path, buf, type);
        const name = String(url.searchParams.get('name') || '').replace(/[\u0000-\u001f\u007f]/g, '').slice(0, 80);
        res.statusCode = 200;
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        return res.end(JSON.stringify({ ok: true, id, path, type, size: buf.length, name }));
    } catch (e) {
        console.error('receipt api error', e);
        return bad(res, 500, 'internal_error');
    }
};
