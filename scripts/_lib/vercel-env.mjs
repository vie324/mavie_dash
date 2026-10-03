// Vercel 環境変数の設定ヘルパー（setup-env.mjs / setup-supabase.mjs から利用）
// 値は Vercel CLI に stdin で渡し、コマンドライン引数やログに出さない。
//
// 認証: `vercel login` 済み、または環境変数 VERCEL_TOKEN。
// プロジェクトの特定: `vercel link` 済み（.vercel/project.json）、または VERCEL_ORG_ID + VERCEL_PROJECT_ID。

import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

const VERCEL_CMD = process.platform === 'win32' ? 'npx.cmd' : 'npx';
const VERCEL_ARGS = ['--yes', 'vercel@latest'];

function baseArgs() {
    const args = [...VERCEL_ARGS];
    if (process.env.VERCEL_TOKEN) args.push('--token', process.env.VERCEL_TOKEN);
    if (process.env.VERCEL_SCOPE) args.push('--scope', process.env.VERCEL_SCOPE);
    return args;
}

export function vercel(args, { input, quiet } = {}) {
    const r = spawnSync(VERCEL_CMD, [...baseArgs(), ...args], {
        input,
        encoding: 'utf8',
        env: { ...process.env, VERCEL_TELEMETRY_DISABLED: '1' },
        stdio: input === undefined ? ['inherit', 'pipe', 'pipe'] : ['pipe', 'pipe', 'pipe'],
    });
    const out = (r.stdout || '') + (r.stderr || '');
    if (!quiet && r.status !== 0) process.stderr.write(out);
    return { ok: r.status === 0, out };
}

export function ensureVercelReady(root) {
    const who = vercel(['whoami'], { quiet: true });
    if (!who.ok) {
        throw new Error('Vercel にログインしていません。`npx vercel login` を実行するか、環境変数 VERCEL_TOKEN を設定してください。');
    }
    const linked = existsSync(join(root, '.vercel', 'project.json'))
        || (process.env.VERCEL_ORG_ID && process.env.VERCEL_PROJECT_ID);
    if (!linked) {
        process.stdout.write('Vercel プロジェクトをリンクします（`vercel link`）...\n');
        const link = vercel(['link', '--yes']);
        if (!link.ok) throw new Error('`vercel link` に失敗しました。VERCEL_ORG_ID / VERCEL_PROJECT_ID を設定するか、対話で `npx vercel link` を実行してください。');
    }
    return who.out.trim().split('\n').pop();
}

// 設定済みの環境変数名（production）を返す
export function listEnvNames(target = 'production') {
    const r = vercel(['env', 'ls', target], { quiet: true });
    if (!r.ok) return new Set();
    const names = new Set();
    for (const line of r.out.split('\n')) {
        const m = line.match(/^\s*([A-Z][A-Z0-9_]+)\s/);
        if (m) names.add(m[1]);
    }
    return names;
}

// 1件設定。既存なら force=true のときだけ上書き。戻り値: 'added' | 'updated' | 'skipped' | 'failed'
export function setEnv(name, value, { targets = ['production'], force = false, existing } = {}) {
    const exists = existing ? existing.has(name) : false;
    if (exists && !force) return 'skipped';
    const args = ['env', 'add', name, targets.join(','), '--sensitive'];
    if (exists) args.push('--force');
    const r = vercel(args, { input: String(value), quiet: true });
    if (!r.ok) {
        // 既に存在する場合の典型メッセージ（listEnvNames が取れなかったとき）
        if (/already exist/i.test(r.out) && !force) return 'skipped';
        if (/already exist/i.test(r.out) && force) {
            const r2 = vercel(['env', 'add', name, targets.join(','), '--sensitive', '--force'], { input: String(value), quiet: true });
            if (r2.ok) return 'updated';
        }
        process.stderr.write(r.out);
        return 'failed';
    }
    return exists ? 'updated' : 'added';
}

export function deployProduction(root) {
    process.stdout.write('本番を再デプロイします（`vercel --prod`）...\n');
    const r = vercel(['--prod', '--yes', '--cwd', root]);
    return r.ok;
}
