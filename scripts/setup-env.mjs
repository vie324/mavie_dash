#!/usr/bin/env node
// Vercel の環境変数（パスワード・AUTH_SECRET・Gemini・Supabase）をまとめて設定して再デプロイするスクリプト。
// 設定タブ「連携状態」の「未設定」を Vercel の画面を開かずに解消するためのもの。
//
// 使い方（対話）:
//   node scripts/setup-env.mjs
// 使い方（非対話。Claude Code などに実行させるとき）:
//   node scripts/setup-env.mjs --yes --generate-passwords [--gemini-key-env GEMINI_API_KEY] [--no-deploy]
//   値はフラグ（--admin-password=...）より、環境変数（VIE_ADMIN_PASSWORD など）で渡す方が履歴に残らず安全です。
//
// 対応する変数と値の渡し方（優先順: フラグ > 環境変数 VIE_<NAME> > 対話/自動生成）
//   AUTH_SECRET                 … 常に自動生成（既に設定済みならスキップ）
//   ADMIN_PASSWORD              … --admin-password / VIE_ADMIN_PASSWORD / --generate-passwords で自動生成
//   MANAGER_PASSWORD            … --manager-password / VIE_MANAGER_PASSWORD / --generate-passwords で自動生成
//   STORE_PASSWORDS             … --store-passwords（JSON） / VIE_STORE_PASSWORDS。省略可（画面の「スタッフアカウントの発行」で店舗の行から発行する方が簡単）
//   GEMINI_API_KEY              … --gemini-key / VIE_GEMINI_API_KEY / --gemini-key-env <別の環境変数名>
//   SUPABASE_URL                … --supabase-url / VIE_SUPABASE_URL（setup-supabase.mjs --write-vercel でも設定できる）
//   SUPABASE_SERVICE_ROLE_KEY   … --supabase-key / VIE_SUPABASE_SERVICE_ROLE_KEY（anon / publishable キーは拒否）
//
// その他: --force（既存の値を上書き）, --preview（Preview 環境にも設定）, --no-deploy（再デプロイしない）, --dry-run
//
// 前提: `npx vercel login` 済み、または VERCEL_TOKEN + VERCEL_ORG_ID + VERCEL_PROJECT_ID（REST API を直接使用。チーム限定トークン可）。
//       CLI 利用時にプロジェクト未リンクなら `vercel link` を自動実行します。
// 自動生成したパスワードはこの実行の標準出力に一度だけ表示します（どこにも保存しません）。

import { randomBytes, randomInt } from 'node:crypto';
import { createInterface } from 'node:readline/promises';
import { stdin, stdout } from 'node:process';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';
import { ensureVercelReady, listEnvNames, setEnv, deployProduction } from './_lib/vercel-env.mjs';

const root = join(fileURLToPath(import.meta.url), '..', '..');

// ---- 引数 ----
const argv = process.argv.slice(2);
const flags = {};
for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (!a.startsWith('--')) continue;
    const eq = a.indexOf('=');
    if (eq > 0) { flags[a.slice(2, eq)] = a.slice(eq + 1); continue; }
    const next = argv[i + 1];
    if (next !== undefined && !next.startsWith('--')) { flags[a.slice(2)] = next; i++; }
    else flags[a.slice(2)] = true;
}
const yes = !!flags.yes;
const dryRun = !!flags['dry-run'];
const force = !!flags.force;
const targets = flags.preview ? ['production', 'preview'] : ['production'];

const PASS_CHARS = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789';
const genPassword = (n = 12) => Array.from({ length: n }, () => PASS_CHARS[randomInt(PASS_CHARS.length)]).join('');

const rl = yes ? null : createInterface({ input: stdin, output: stdout });
async function ask(q) {
    if (!rl) return '';
    const answer = await rl.question(q);
    return answer.trim();
}

function fromFlagOrEnv(flag, envName) {
    if (typeof flags[flag] === 'string' && flags[flag] !== '') return flags[flag];
    if (process.env['VIE_' + envName]) return process.env['VIE_' + envName];
    return '';
}

// api/_lib/kv.js の supabaseKeyProblem と同じ判定
function supabaseKeyProblem(key) {
    if (/^sb_publishable_/.test(key)) return 'publishable キー（ブラウザ用）です。service_role（secret）キーを使ってください';
    const parts = key.split('.');
    if (parts.length === 3) {
        try {
            const payload = JSON.parse(Buffer.from(parts[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8'));
            if (payload?.role && payload.role !== 'service_role') return `${payload.role} キーです。service_role（secret）キーを使ってください`;
        } catch (_) { /* JWTでなければ判定しない */ }
    }
    return null;
}

async function main() {
    let user = '(dry-run)';
    let existing = new Set();
    if (!dryRun) {
        user = await ensureVercelReady(root);
        existing = await listEnvNames('production');
    } else {
        // dry-run でもログイン済みなら実際の設定状況を表示する（設定後の確認に使う）。未ログインなら失敗させない
        try {
            user = `${await ensureVercelReady(root)} (dry-run)`;
            existing = await listEnvNames('production');
        } catch (_) {
            stdout.write('（Vercel 未ログインのため設定状況を取得できません。以下はすべて「未設定」として表示します）\n');
        }
    }
    stdout.write(`Vercel: ${user}\n`);
    const show = (name) => existing.has(name) ? '設定済み' : '未設定';
    stdout.write(['AUTH_SECRET', 'ADMIN_PASSWORD', 'MANAGER_PASSWORD', 'STORE_PASSWORDS', 'GEMINI_API_KEY', 'SUPABASE_URL', 'SUPABASE_SERVICE_ROLE_KEY']
        .map(n => `  ${n.padEnd(26)} ${show(n)}`).join('\n') + '\n\n');

    const plan = [];   // { name, value, note }
    const generated = []; // 画面に一度だけ出すもの

    // AUTH_SECRET: 常に自動生成
    if (!existing.has('AUTH_SECRET') || force) plan.push({ name: 'AUTH_SECRET', value: randomBytes(32).toString('hex'), note: '自動生成' });

    // パスワード
    for (const [name, label, flag] of [
        ['ADMIN_PASSWORD', 'オーナー（/ のパスワード）', 'admin-password'],
        ['MANAGER_PASSWORD', 'マネージャー（?mode=manager）', 'manager-password'],
    ]) {
        if (existing.has(name) && !force) continue;
        let v = fromFlagOrEnv(flag, name);
        if (!v && flags['generate-passwords']) { v = genPassword(); generated.push([label, v]); }
        if (!v && rl) {
            v = await ask(`${label} のパスワード（空Enterで自動生成、"-" でスキップ）: `);
            if (v === '-') v = '';
            else if (!v) { v = genPassword(); generated.push([label, v]); }
        }
        if (v) plan.push({ name, value: v, note: label });
    }

    // 店長パスワード（JSON）
    if (!existing.has('STORE_PASSWORDS') || force) {
        let v = fromFlagOrEnv('store-passwords', 'STORE_PASSWORDS');
        if (!v && rl) v = await ask('店長パスワード STORE_PASSWORDS（JSON。例 {"chiba":"pass1"}。空Enterでスキップ＝画面から発行）: ');
        if (v) {
            try { JSON.parse(v); } catch (_) { throw new Error('STORE_PASSWORDS が JSON ではありません'); }
            plan.push({ name: 'STORE_PASSWORDS', value: v, note: '店長' });
        }
    }

    // Gemini
    if (!existing.has('GEMINI_API_KEY') || force) {
        let v = fromFlagOrEnv('gemini-key', 'GEMINI_API_KEY');
        if (!v && typeof flags['gemini-key-env'] === 'string') v = process.env[flags['gemini-key-env']] || '';
        if (!v && rl) v = await ask('GEMINI_API_KEY（Google AI Studio で発行。空Enterでスキップ）: ');
        if (v) plan.push({ name: 'GEMINI_API_KEY', value: v, note: 'AIアドバイス' });
    }

    // Supabase
    if (!existing.has('SUPABASE_URL') || !existing.has('SUPABASE_SERVICE_ROLE_KEY') || force) {
        let url = fromFlagOrEnv('supabase-url', 'SUPABASE_URL');
        let key = fromFlagOrEnv('supabase-key', 'SUPABASE_SERVICE_ROLE_KEY');
        if (!url && rl) url = await ask('SUPABASE_URL（Project Settings → API の Project URL。空Enterでスキップ）: ');
        if (url && !key && rl) key = await ask('SUPABASE_SERVICE_ROLE_KEY（service_role / secret キー）: ');
        if (url && key) {
            if (!/^https:\/\/[a-z0-9-]+\.supabase\.(co|in)$/.test(url.replace(/\/+$/, ''))) throw new Error(`SUPABASE_URL の形式が不正です: ${url}`);
            const problem = supabaseKeyProblem(key);
            if (problem) throw new Error(`SUPABASE_SERVICE_ROLE_KEY: ${problem}`);
            plan.push({ name: 'SUPABASE_URL', value: url.replace(/\/+$/, ''), note: 'サーバー保存' });
            plan.push({ name: 'SUPABASE_SERVICE_ROLE_KEY', value: key, note: 'サーバー保存' });
        } else if (url || key) {
            stdout.write('  ⚠ SUPABASE_URL と SUPABASE_SERVICE_ROLE_KEY は両方必要です。Supabase の設定はスキップします\n');
        }
    }
    rl?.close();

    if (plan.length === 0) {
        stdout.write('設定する変数はありません（すべて設定済み。上書きは --force）\n');
        return;
    }
    stdout.write('設定する変数:\n' + plan.map(p => `  ${p.name.padEnd(26)} ${p.note}`).join('\n') + '\n');
    if (dryRun) { stdout.write('(dry-run のため何もしません)\n'); return; }

    let failed = 0;
    for (const p of plan) {
        const r = await setEnv(p.name, p.value, { targets, force, existing });
        stdout.write(`  ${p.name.padEnd(26)} ${r}\n`);
        if (r === 'failed') failed++;
    }
    if (generated.length) {
        stdout.write('\n==== 自動生成したパスワード（この画面にしか出ません。今すぐ控えてください）====\n');
        for (const [label, v] of generated) stdout.write(`  ${label}: ${v}\n`);
        stdout.write('===========================================================================\n');
    }
    if (failed) throw new Error(`${failed} 件の設定に失敗しました`);

    if (flags['no-deploy']) {
        stdout.write('\n環境変数は次のデプロイから有効になります（--no-deploy のため再デプロイしていません）\n');
        return;
    }
    if (!(await deployProduction(root))) throw new Error('再デプロイに失敗しました。Vercel の画面から Redeploy してください');
    stdout.write('\n完了。ダッシュボードの 設定 → 連携状態 で「設定済み」になっていることを確認してください\n');
}

main().catch(e => { process.stderr.write(`\nエラー: ${e.message}\n`); process.exit(1); });
