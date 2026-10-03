#!/usr/bin/env node
// Supabase（サーバー保存先）のセットアップをコマンドラインで行うスクリプト。
// SQL Editor を開かずに supabase/schema.sql を適用し、必要なら Vercel の環境変数まで設定します。
//
// 前提: Supabase の Personal Access Token を環境変数 SUPABASE_ACCESS_TOKEN に設定
//       （https://supabase.com/dashboard/account/tokens で発行。チャットには貼らない）
//
// 使い方:
//   node scripts/setup-supabase.mjs orgs                                  … 組織一覧（create 用）
//   node scripts/setup-supabase.mjs create --name vie-dash [--org <id>]   … プロジェクト作成（東京リージョン）
//   node scripts/setup-supabase.mjs schema --ref <project-ref>            … schema.sql を適用（何度でも安全）
//   node scripts/setup-supabase.mjs verify --ref <project-ref>            … テーブルと service_role キーの疎通確認
//   node scripts/setup-supabase.mjs all --ref <project-ref> --write-vercel … schema + verify + Vercel に SUPABASE_URL / KEY を設定して再デプロイ
//
// --ref は SUPABASE_URL（https://<ref>.supabase.co）または環境変数 SUPABASE_PROJECT_REF でも指定できます。
// --write-vercel は `npx vercel login` 済み（または VERCEL_TOKEN）が必要。キーは画面に表示せず Vercel に直接渡します。

import { readFile } from 'node:fs/promises';
import { randomBytes } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const root = join(fileURLToPath(import.meta.url), '..', '..');
const API = 'https://api.supabase.com/v1';

const argv = process.argv.slice(2);
const cmd = argv.find(a => !a.startsWith('--')) || 'help';
const flags = {};
for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith('--')) continue;
    const eq = a.indexOf('=');
    if (eq > 0) { flags[a.slice(2, eq)] = a.slice(eq + 1); continue; }
    const next = argv[i + 1];
    if (next !== undefined && !next.startsWith('--')) { flags[a.slice(2)] = next; i++; } else flags[a.slice(2)] = true;
}

function token() {
    const t = process.env.SUPABASE_ACCESS_TOKEN;
    if (!t) throw new Error('環境変数 SUPABASE_ACCESS_TOKEN が未設定です（https://supabase.com/dashboard/account/tokens で発行）');
    return t;
}

async function api(path, { method = 'GET', body } = {}) {
    const res = await fetch(API + path, {
        method,
        headers: { Authorization: `Bearer ${token()}`, 'Content-Type': 'application/json', Accept: 'application/json' },
        body: body === undefined ? undefined : JSON.stringify(body),
    });
    const text = await res.text();
    let json = null;
    try { json = text ? JSON.parse(text) : null; } catch (_) { /* テキストのまま */ }
    if (!res.ok) throw new Error(`Supabase API ${method} ${path} → ${res.status}: ${json?.message || text.slice(0, 300)}`);
    return json;
}

function projectRef() {
    const direct = typeof flags.ref === 'string' ? flags.ref : process.env.SUPABASE_PROJECT_REF;
    if (direct) return direct.replace(/^https?:\/\//, '').split('.')[0];
    const url = process.env.SUPABASE_URL || process.env.VIE_SUPABASE_URL;
    const m = url && url.match(/^https:\/\/([a-z0-9-]+)\.supabase\./);
    if (m) return m[1];
    throw new Error('プロジェクトを特定できません。--ref <project-ref>（または SUPABASE_URL）を指定してください');
}

async function serviceRoleKey(ref) {
    const keys = await api(`/projects/${ref}/api-keys?reveal=true`);
    const list = Array.isArray(keys) ? keys : [];
    const sr = list.find(k => k.name === 'service_role') || list.find(k => k.type === 'secret' && /service/i.test(k.name || ''));
    if (!sr?.api_key) throw new Error('service_role キーを取得できませんでした（Project Settings → API からコピーして setup-env.mjs で設定してください）');
    return sr.api_key;
}

async function cmdOrgs() {
    const orgs = await api('/organizations');
    for (const o of orgs) process.stdout.write(`  ${o.id}  ${o.name}\n`);
    if (!orgs.length) process.stdout.write('  組織がありません。https://supabase.com/dashboard で作成してください\n');
}

async function cmdCreate() {
    const name = typeof flags.name === 'string' ? flags.name : 'vie-dash';
    let org = typeof flags.org === 'string' ? flags.org : '';
    if (!org) {
        const orgs = await api('/organizations');
        if (orgs.length !== 1) throw new Error(`組織が ${orgs.length} 件あります。--org <id> で指定してください（一覧: orgs）`);
        org = orgs[0].id;
    }
    const dbPass = randomBytes(24).toString('base64url');
    const p = await api('/projects', {
        method: 'POST',
        body: { name, organization_id: org, region: flags.region || 'ap-northeast-1', db_pass: dbPass, plan: 'free' },
    });
    process.stdout.write(`作成しました: ${p.name}  ref=${p.id}  region=${p.region}\n`);
    process.stdout.write('  DBパスワード（SQL Editor や psql で直接つなぐとき以外は不要。Supabase の画面でいつでも再設定できます）: ' + dbPass + '\n');
    process.stdout.write('  プロビジョニングに1〜2分かかります。その後 `schema --ref ' + p.id + '` を実行してください\n');
}

async function waitActive(ref) {
    for (let i = 0; i < 30; i++) {
        const p = await api(`/projects/${ref}`);
        if (p.status === 'ACTIVE_HEALTHY') return p;
        process.stdout.write(`  プロジェクト状態: ${p.status} … 待機中\n`);
        await new Promise(r => setTimeout(r, 10000));
    }
    throw new Error('プロジェクトが起動しません（Supabase の画面で状態を確認してください）');
}

async function cmdSchema(ref) {
    await waitActive(ref);
    const sql = await readFile(join(root, 'supabase', 'schema.sql'), 'utf8');
    await api(`/projects/${ref}/database/query`, { method: 'POST', body: { query: sql } });
    process.stdout.write(`schema.sql を適用しました（ref=${ref}）\n`);
}

async function cmdVerify(ref) {
    const url = `https://${ref}.supabase.co`;
    const key = await serviceRoleKey(ref);
    const res = await fetch(`${url}/rest/v1/vie_kv?select=key&limit=1`, { headers: { apikey: key, Authorization: `Bearer ${key}`, Accept: 'application/json' } });
    if (!res.ok) throw new Error(`vie_kv テーブルにアクセスできません（${res.status}）。schema を先に実行してください`);
    // anon からは見えないこと（RLS）の確認
    const keys = await api(`/projects/${ref}/api-keys?reveal=true`);
    const anon = (Array.isArray(keys) ? keys : []).find(k => k.name === 'anon');
    if (anon?.api_key) {
        const r2 = await fetch(`${url}/rest/v1/vie_kv?select=key&limit=1`, { headers: { apikey: anon.api_key, Authorization: `Bearer ${anon.api_key}` } });
        if (r2.ok) process.stdout.write('  ⚠ anon キーで vie_kv が読めています。schema.sql（RLS 有効化）を再実行してください\n');
    }
    process.stdout.write(`OK: ${url} の vie_kv に service_role で接続できます\n`);
    return { url, key };
}

async function cmdWriteVercel(ref) {
    const { url, key } = await cmdVerify(ref);
    const { ensureVercelReady, listEnvNames, setEnv, deployProduction } = await import('./_lib/vercel-env.mjs');
    ensureVercelReady(root);
    const existing = listEnvNames('production');
    const force = !!flags.force;
    for (const [name, value] of [['SUPABASE_URL', url], ['SUPABASE_SERVICE_ROLE_KEY', key]]) {
        const r = setEnv(name, value, { force, existing });
        process.stdout.write(`  Vercel ${name.padEnd(26)} ${r}\n`);
        if (r === 'failed') throw new Error(`${name} の設定に失敗しました`);
        if (r === 'skipped') process.stdout.write('    （既存の値を残しました。上書きは --force）\n');
    }
    if (flags['no-deploy']) return;
    if (!deployProduction(root)) throw new Error('再デプロイに失敗しました');
    process.stdout.write('完了。設定 → 連携状態 の「日報・目標・シフトの保存」が「Supabaseに保存」になれば成功です\n');
}

async function main() {
    switch (cmd) {
        case 'orgs': return cmdOrgs();
        case 'create': return cmdCreate();
        case 'schema': return cmdSchema(projectRef());
        case 'verify': { await cmdVerify(projectRef()); return; }
        case 'all': {
            const ref = projectRef();
            await cmdSchema(ref);
            if (flags['write-vercel']) return cmdWriteVercel(ref);
            await cmdVerify(ref);
            process.stdout.write('次: `node scripts/setup-supabase.mjs all --ref ' + ref + ' --write-vercel` で Vercel に設定できます\n');
            return;
        }
        default:
            process.stdout.write(`使い方:
  node scripts/setup-supabase.mjs orgs
  node scripts/setup-supabase.mjs create --name vie-dash [--org <id>] [--region ap-northeast-1]
  node scripts/setup-supabase.mjs schema --ref <project-ref>
  node scripts/setup-supabase.mjs verify --ref <project-ref>
  node scripts/setup-supabase.mjs all --ref <project-ref> [--write-vercel] [--force] [--no-deploy]
環境変数: SUPABASE_ACCESS_TOKEN（必須）, VERCEL_TOKEN（--write-vercel で未ログインのとき）
`);
    }
}

main().catch(e => { process.stderr.write(`\nエラー: ${e.message}\n`); process.exit(1); });
