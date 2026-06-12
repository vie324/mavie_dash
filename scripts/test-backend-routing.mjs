#!/usr/bin/env node
/**
 * backend.js の apiFetch ルーティング単体テスト（ブラウザAPIをスタブしてNodeで実行）
 *   node scripts/test-backend-routing.mjs
 * GASの各actionが正しい Supabase RPC + 引数に変換されること、
 * GASモードで素通しされることを検証する。
 */
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

// ---- ブラウザAPIスタブ ----
const store = new Map();
globalThis.localStorage = {
    getItem: k => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => store.set(k, String(v)),
    removeItem: k => store.delete(k),
};
globalThis.document = { addEventListener: () => {}, getElementById: () => null, querySelectorAll: () => [], querySelector: () => null, createElement: () => ({ style: {} }), head: { appendChild: () => {} } };
globalThis.window = globalThis;
globalThis.location = { href: 'https://example.com/index.html' };
// crypto.randomUUID は Node 19+ で標準搭載（getter専用globalのため代入しない）

const calls = [];
const realFetch = globalThis.fetch;
globalThis.fetch = async (url, options = {}) => {
    calls.push({ url: String(url), options });
    return new Response(JSON.stringify([{ mocked: true }]), { status: 200, headers: { 'Content-Type': 'application/json' } });
};

// ---- backend.js をロード ----
const src = readFileSync(join(root, 'assets/js/backend.js'), 'utf8');
eval(src);

let failed = 0;
function expect(name, cond, detail = '') {
    if (cond) console.log(`  OK   ${name}`);
    else { console.error(`  FAIL ${name} ${detail}`); failed++; }
}
const lastCall = () => calls[calls.length - 1];
const lastRpc = () => {
    const c = lastCall();
    return { fn: c.url.split('/rest/v1/rpc/')[1], args: c.options.body ? JSON.parse(c.options.body) : null, headers: c.options.headers };
};

const GAS = 'https://script.google.com/macros/s/XXX/exec';

(async () => {
    console.log('— GASモード: 素通し —');
    localStorage.setItem('mavie_backend_mode', 'gas');
    await apiFetch(`${GAS}?action=load_goals`);
    expect('GASモードはそのままfetch', lastCall().url.includes('script.google.com'));

    console.log('— Supabaseモード: RPCルーティング —');
    localStorage.setItem('mavie_backend_mode', 'supabase');
    localStorage.setItem('mavie_supabase_url', 'https://proj.supabase.co');
    localStorage.setItem('mavie_supabase_anon_key', 'anonkey');

    let res = await apiFetch(GAS); // action無し = get_data
    expect('get_data → api_get_sales', lastRpc().fn === 'api_get_sales');
    expect('anon keyヘッダー付与', lastRpc().headers.apikey === 'anonkey');
    expect('Responseとして配列が返る', Array.isArray(await res.json()));

    await apiFetch(`${GAS}?action=get_customers_by_store&store=chiba`);
    expect('by_store → p_store=chiba', lastRpc().fn === 'api_get_customers' && lastRpc().args.p_store === 'chiba');

    await apiFetch(`${GAS}?action=get_customers_today`);
    expect('today → p_today=true', lastRpc().args.p_today === true);

    await apiFetch(GAS, { method: 'POST', body: JSON.stringify({ action: 'save_goals', goals: { a: 1 }, salaries: { b: 2 } }) });
    expect('save_goals → api_save_goals(p_goals,p_salaries)', lastRpc().fn === 'api_save_goals' && lastRpc().args.p_goals.a === 1 && lastRpc().args.p_salaries.b === 2);

    await apiFetch(GAS, { method: 'POST', body: JSON.stringify({ action: 'add_record', record: { date: '2026/6/3' } }) });
    expect('add_record → api_add_record(p)', lastRpc().fn === 'api_add_record' && lastRpc().args.p.date === '2026/6/3');

    await apiFetch(GAS, { method: 'POST', body: JSON.stringify({ action: 'update', rows: [{ id: 1 }] }) });
    expect('update → api_update_rows(p)', lastRpc().fn === 'api_update_rows' && lastRpc().args.p[0].id === 1);

    await apiFetch(GAS, { method: 'POST', body: JSON.stringify({ action: 'save_settings', settings: { staffRoster: {} } }) });
    expect('save_settings → api_save_settings(p)', lastRpc().fn === 'api_save_settings');

    await apiFetch(GAS, { method: 'POST', body: JSON.stringify({ action: 'save_passwords', passwords: { chiba_kiki: 'x' } }) });
    expect('save_passwords → api_save_passwords(p)', lastRpc().fn === 'api_save_passwords' && lastRpc().args.p.chiba_kiki === 'x');

    // verify_password: ok=true を返すRPCをモック
    globalThis.fetch = async (url, options = {}) => {
        calls.push({ url: String(url), options });
        return new Response(JSON.stringify({ ok: true }), { status: 200 });
    };
    res = await apiFetch(`${GAS}?action=verify_password&page_type=admin&store=&staff=&password=secret`);
    let body = await res.json();
    expect('verify_password 成功 → sessionToken発行', body.status === 'success' && !!body.sessionToken);
    expect('RPC引数にパスワードが渡る', lastRpc().args.p_password === 'secret');

    res = await apiFetch(`${GAS}?action=verify_session&session_token=${body.sessionToken}&page_type=admin`);
    body = await res.json();
    expect('発行済みセッションの検証成功', body.status === 'success');

    res = await apiFetch(`${GAS}?action=verify_session&session_token=bogus&page_type=admin`);
    body = await res.json();
    expect('不正セッションは拒否', body.status === 'error');

    // load_settings 形状
    globalThis.fetch = async (url, options = {}) => {
        calls.push({ url: String(url), options });
        return new Response(JSON.stringify({ staffRoster: { chiba: ['kiki'] }, adCosts: null, monthlyClose: null, geminiApiKey: null }), { status: 200 });
    };
    res = await apiFetch(`${GAS}?action=load_settings`);
    body = await res.json();
    expect('load_settings → {status, settings.staffRoster}', body.status === 'success' && body.settings.staffRoster.chiba[0] === 'kiki');

    // RPCエラー時もGAS互換 {status:'error'} で返る（fetch例外にしない）
    globalThis.fetch = async () => new Response(JSON.stringify({ message: 'relation does not exist' }), { status: 404 });
    res = await apiFetch(`${GAS}?action=load_goals`);
    body = await res.json();
    expect('RPCエラー → status:error に変換', body.status === 'error' && body.message.includes('does not exist'));

    globalThis.fetch = realFetch;
    console.log(failed ? `\n❌ ${failed} 件失敗` : '\n✅ 全テスト合格');
    process.exit(failed ? 1 : 0);
})();
