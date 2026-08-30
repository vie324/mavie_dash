// GET /api/data-proxy?path=<エンドポイント> — SalonOneプロキシの実体
// フロントは /api/data/<path> を呼び、vercel.json の rewrite でここへ届く。
// （Vercelのキャッチオールが複数セグメントのパスにマッチしない問題の回避策）

'use strict';

module.exports = require('./_lib/data-handler');
