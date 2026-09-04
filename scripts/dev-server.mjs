// ローカル開発サーバー: Vercelの「静的ファイル + api/ 関数」構成を模倣する
// 使い方: node scripts/dev-server.mjs [port]
// SALONONE_API_KEY 未設定ならデモデータで動作する。

import http from 'node:http';
import { createRequire } from 'node:module';
import { readFile } from 'node:fs/promises';
import { extname, join, normalize } from 'node:path';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const root = join(fileURLToPath(import.meta.url), '..', '..');
process.env.DEV_SERVER = '1';

const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.mjs': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json',
    '.png': 'image/png',
    '.svg': 'image/svg+xml',
    '.webmanifest': 'application/manifest+json',
    '.ico': 'image/x-icon',
};

const routes = [
    { re: /^\/api\/auth\/session$/, handler: () => require(join(root, 'api/auth/session.js')) },
    { re: /^\/api\/auth\/login$/, handler: () => require(join(root, 'api/auth/login.js')) },
    { re: /^\/api\/auth\/logout$/, handler: () => require(join(root, 'api/auth/logout.js')) },
    { re: /^\/api\/ai$/, handler: () => require(join(root, 'api/ai.js')) },
    { re: /^\/api\/manual$/, handler: () => require(join(root, 'api/manual.js')) },
    { re: /^\/api\/shift$/, handler: () => require(join(root, 'api/shift.js')) },
    { re: /^\/api\/accounts$/, handler: () => require(join(root, 'api/accounts.js')) },
    { re: /^\/api\/goals$/, handler: () => require(join(root, 'api/goals.js')) },
    { re: /^\/api\/data(-proxy|\/.+)$/, handler: () => require(join(root, 'api/_lib/data-handler.js')) },
];

const server = http.createServer(async (req, res) => {
    const url = new URL(req.url, 'http://localhost');
    try {
        for (const r of routes) {
            if (r.re.test(url.pathname)) {
                res.setHeader('Content-Type', 'application/json; charset=utf-8');
                await r.handler()(req, res);
                return;
            }
        }
        // 静的ファイル
        let path = url.pathname === '/' ? '/index.html' : url.pathname;
        path = normalize(path).replace(/^(\.\.[/\\])+/, '');
        const file = join(root, path);
        if (!file.startsWith(root)) { res.statusCode = 403; return res.end(); }
        try {
            const data = await readFile(file);
            res.setHeader('Content-Type', MIME[extname(file)] || 'application/octet-stream');
            res.end(data);
        } catch {
            res.statusCode = 404;
            res.end('not found');
        }
    } catch (e) {
        console.error('server error', e);
        res.statusCode = 500;
        res.end(JSON.stringify({ error: 'internal_error' }));
    }
});

const port = Number(process.argv[2]) || 3000;
server.listen(port, () => console.log(`dev server: http://localhost:${port}`));
