// GET /api/manifest?store=..&staff=..&mode=..
// PWA マニフェストを閲覧コンテキスト付きで返す。
// 静的な manifest.webmanifest は start_url が固定（/index.html）のため、スタッフ・店長の専用URL
// （?store=..&staff=..）を「ホーム画面に追加」するとパラメータが落ちてオーナーログインになってしまう。
// index.html の <head> で、URLにパラメータがあるときだけ <link rel="manifest"> をこのAPIに差し替える。

'use strict';

const { readFileSync } = require('fs');
const { join } = require('path');

let base = null;
function baseManifest() {
    if (!base) base = JSON.parse(readFileSync(join(__dirname, '..', 'manifest.webmanifest'), 'utf8'));
    return base;
}

// 値は 64 文字までの英数字・ハイフン・アンダースコア・日本語のみ（それ以外は無視して既定の start_url にする）
const SAFE = /^[\w\-぀-ヿ一-鿿]{1,64}$/;

function contextParams(query) {
    const params = new URLSearchParams();
    for (const key of ['store', 'shop_id', 'staff', 'staff_id', 'mode']) {
        const v = query[key];
        if (typeof v === 'string' && SAFE.test(v)) params.set(key, v);
    }
    return params;
}

module.exports = (req, res) => {
    res.setHeader('Content-Type', 'application/manifest+json; charset=utf-8');
    res.setHeader('Cache-Control', 'public, max-age=3600');
    if (req.method !== 'GET') {
        res.statusCode = 405;
        return res.end('{}');
    }
    const url = new URL(req.url, 'http://localhost');
    const query = Object.fromEntries(url.searchParams.entries());
    const params = contextParams(query);
    const qs = params.toString();
    const startUrl = qs ? `/?${qs}` : '/';
    const manifest = {
        ...baseManifest(),
        // 絶対パスにする（/api/ 配下から返すため、相対だと /api/ 基準になってしまう）
        id: startUrl,
        start_url: startUrl,
        scope: '/',
        icons: (baseManifest().icons || []).map(i => ({ ...i, src: '/' + String(i.src).replace(/^\.?\//, '') })),
    };
    res.statusCode = 200;
    res.end(JSON.stringify(manifest));
};
