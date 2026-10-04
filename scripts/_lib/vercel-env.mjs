// Vercel 環境変数の設定ヘルパー（setup-env.mjs / setup-supabase.mjs から利用）
// 値は Vercel CLI に stdin で渡す（または REST API のリクエスト本文に入れる）。コマンドライン引数やログに出さない。
//
// 認証と動作モード:
//   - VERCEL_TOKEN + VERCEL_ORG_ID + VERCEL_PROJECT_ID が揃っていれば REST API を直接使う（API モード）。
//     チーム限定スコープのトークンは CLI の `vercel whoami` が通らない（User not found）ため、クラウドセッションではこちらを使う。
//   - それ以外は Vercel CLI（`vercel login` 済み、または VERCEL_TOKEN のみ）。プロジェクト未リンクなら `vercel link` を自動実行。
//
// すべて async。戻り値の仕様は両モードで共通。

import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { join } from 'node:path';

const VERCEL_CMD = process.platform === 'win32' ? 'npx.cmd' : 'npx';
const VERCEL_ARGS = ['--yes', 'vercel@latest'];
const API = 'https://api.vercel.com';

const useApi = () => !!(process.env.VERCEL_TOKEN && process.env.VERCEL_ORG_ID && process.env.VERCEL_PROJECT_ID);

// ---------- REST API モード ----------
async function api(path, { method = 'GET', body } = {}) {
    const sep = path.includes('?') ? '&' : '?';
    const res = await fetch(`${API}${path}${sep}teamId=${encodeURIComponent(process.env.VERCEL_ORG_ID)}`, {
        method,
        headers: { Authorization: `Bearer ${process.env.VERCEL_TOKEN}`, 'Content-Type': 'application/json' },
        body: body === undefined ? undefined : JSON.stringify(body),
    });
    const text = await res.text();
    let json = null;
    try { json = text ? JSON.parse(text) : null; } catch (_) { /* テキストのまま */ }
    if (!res.ok) {
        const err = new Error(`Vercel API ${method} ${path} → ${res.status}: ${json?.error?.message || text.slice(0, 300)}`);
        err.status = res.status;
        err.code = json?.error?.code;
        throw err;
    }
    return json;
}

const projectPath = () => `/v9/projects/${encodeURIComponent(process.env.VERCEL_PROJECT_ID)}`;

async function apiEnsureReady() {
    let project;
    try {
        project = await api(projectPath());
    } catch (e) {
        if (e.status === 401 || e.status === 403 || e.status === 404) {
            throw new Error(`Vercel のトークンでプロジェクトを取得できません（${e.status}）。VERCEL_TOKEN のスコープ・期限と VERCEL_ORG_ID / VERCEL_PROJECT_ID を確認してください`);
        }
        throw e;
    }
    return `${project.name} (API, team ${process.env.VERCEL_ORG_ID})`;
}

async function apiListEnvNames(target) {
    const r = await api(`${projectPath()}/env`);
    const names = new Set();
    for (const e of r?.envs || []) {
        if (!target || (e.target || []).includes(target)) names.add(e.key);
    }
    return names;
}

async function apiSetEnv(name, value, { targets, force, exists }) {
    try {
        await api(`/v10/projects/${encodeURIComponent(process.env.VERCEL_PROJECT_ID)}/env${force ? '?upsert=true' : ''}`, {
            method: 'POST',
            body: { key: name, value: String(value), type: 'sensitive', target: targets },
        });
        return exists ? 'updated' : 'added';
    } catch (e) {
        if (e.code === 'ENV_ALREADY_EXISTS' || /already exist/i.test(e.message)) {
            if (!force) return 'skipped';
        }
        process.stderr.write(`${e.message}\n`);
        return 'failed';
    }
}

async function apiDeployProduction() {
    process.stdout.write('本番を再デプロイします（最新の本番デプロイを API で再デプロイ）...\n');
    const list = await api(`/v6/deployments?projectId=${encodeURIComponent(process.env.VERCEL_PROJECT_ID)}&target=production&limit=1`);
    const latest = list?.deployments?.[0];
    if (!latest) { process.stderr.write('本番デプロイがまだありません。先に一度デプロイしてください\n'); return false; }
    let dep;
    try {
        dep = await api('/v13/deployments?forceNew=1', { method: 'POST', body: { name: latest.name, deploymentId: latest.uid, target: 'production' } });
    } catch (e) {
        process.stderr.write(`${e.message}\n`);
        return false;
    }
    process.stdout.write(`  デプロイ開始: https://${dep.url}\n`);
    // 完了まで待つ（最大 10 分）
    const started = Date.now();
    while (Date.now() - started < 10 * 60 * 1000) {
        await new Promise(r => setTimeout(r, 5000));
        const d = await api(`/v13/deployments/${encodeURIComponent(dep.id)}`);
        const state = d.readyState || d.state;
        if (state === 'READY') { process.stdout.write(`  デプロイ完了: https://${d.url}\n`); return true; }
        if (state === 'ERROR' || state === 'CANCELED') { process.stderr.write(`  デプロイが ${state} で終了しました: https://${d.url}\n`); return false; }
    }
    process.stderr.write('  デプロイの完了待ちがタイムアウトしました（Vercel の画面で状態を確認してください）\n');
    return false;
}

// ---------- CLI モード ----------
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

function cliEnsureReady(root) {
    const who = vercel(['whoami'], { quiet: true });
    if (!who.ok) {
        throw new Error('Vercel にログインしていません。`npx vercel login` を実行するか、環境変数 VERCEL_TOKEN（＋ VERCEL_ORG_ID / VERCEL_PROJECT_ID）を設定してください。');
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

function cliListEnvNames(target) {
    const r = vercel(['env', 'ls', target], { quiet: true });
    if (!r.ok) return new Set();
    const names = new Set();
    for (const line of r.out.split('\n')) {
        const m = line.match(/^\s*([A-Z][A-Z0-9_]+)\s/);
        if (m) names.add(m[1]);
    }
    return names;
}

function cliSetEnv(name, value, { targets, force, exists }) {
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

function cliDeployProduction(root) {
    process.stdout.write('本番を再デプロイします（`vercel --prod`）...\n');
    const r = vercel(['--prod', '--yes', '--cwd', root]);
    return r.ok;
}

// ---------- 公開 API（両モード共通） ----------

// ログイン・プロジェクト特定の確認。戻り値は表示用の文字列（ユーザー名やプロジェクト名）
export async function ensureVercelReady(root) {
    return useApi() ? apiEnsureReady() : cliEnsureReady(root);
}

// 設定済みの環境変数名（production）を返す
export async function listEnvNames(target = 'production') {
    return useApi() ? apiListEnvNames(target) : cliListEnvNames(target);
}

// 1件設定。既存なら force=true のときだけ上書き。戻り値: 'added' | 'updated' | 'skipped' | 'failed'
export async function setEnv(name, value, { targets = ['production'], force = false, existing } = {}) {
    const exists = existing ? existing.has(name) : false;
    if (exists && !force) return 'skipped';
    return useApi() ? apiSetEnv(name, value, { targets, force, exists }) : cliSetEnv(name, value, { targets, force, exists });
}

// 本番を再デプロイ。成功なら true
export async function deployProduction(root) {
    return useApi() ? apiDeployProduction() : cliDeployProduction(root);
}
