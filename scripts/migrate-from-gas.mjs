#!/usr/bin/env node
/**
 * GAS(スプレッドシート) → Supabase データ移行スクリプト
 *
 * 使い方:
 *   GAS_URL="https://script.google.com/macros/s/xxx/exec" \
 *   SUPABASE_URL="https://xxxx.supabase.co" \
 *   SUPABASE_SERVICE_KEY="eyJ..." \
 *   node scripts/migrate-from-gas.mjs [--wipe] [--dry-run]
 *
 *   --wipe    : 移行前に daily_reports / customers を全削除（再実行時の重複防止）
 *   --dry-run : 取得・変換のみ行い、書き込みしない
 *
 * 前提: supabase/migrations/0001_init.sql 適用済み。
 * SERVICE_KEY は Supabase Dashboard > Settings > API の service_role（絶対に公開しない）。
 */

const GAS_URL = process.env.GAS_URL;
const SB_URL = (process.env.SUPABASE_URL || '').replace(/\/+$/, '');
const SB_KEY = process.env.SUPABASE_SERVICE_KEY;
const WIPE = process.argv.includes('--wipe');
const DRY = process.argv.includes('--dry-run');

if (!GAS_URL || !SB_URL || !SB_KEY) {
    console.error('環境変数 GAS_URL / SUPABASE_URL / SUPABASE_SERVICE_KEY を設定してください');
    process.exit(1);
}

const sbHeaders = {
    apikey: SB_KEY,
    Authorization: `Bearer ${SB_KEY}`,
    'Content-Type': 'application/json',
};

async function gasGet(action) {
    const sep = GAS_URL.includes('?') ? '&' : '?';
    const url = action ? `${GAS_URL}${sep}action=${action}&nocache=true` : `${GAS_URL}${sep}nocache=true`;
    const res = await fetch(url, { redirect: 'follow' });
    if (!res.ok) throw new Error(`GAS ${action || 'get_data'}: HTTP ${res.status}`);
    return res.json();
}

async function sbRest(method, path, body) {
    const res = await fetch(`${SB_URL}${path}`, {
        method,
        headers: { ...sbHeaders, Prefer: 'return=minimal' },
        body: body === undefined ? undefined : JSON.stringify(body),
    });
    if (!res.ok) throw new Error(`Supabase ${method} ${path}: HTTP ${res.status} ${await res.text()}`);
}

async function sbRpc(fn, args) {
    const res = await fetch(`${SB_URL}/rest/v1/rpc/${fn}`, {
        method: 'POST',
        headers: sbHeaders,
        body: JSON.stringify(args || {}),
    });
    if (!res.ok) throw new Error(`Supabase rpc/${fn}: HTTP ${res.status} ${await res.text()}`);
    return res.json().catch(() => null);
}

const toInt = v => { const n = parseInt(v, 10); return Number.isFinite(n) ? n : 0; };

// 'YYYY/M/D' → 'YYYY-MM-DD'
function toIsoDate(s) {
    if (!s) return null;
    const m = String(s).match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (!m) return null;
    return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
}

function mapReport(d) {
    const s = d.sales || {}, dc = d.discounts || {}, c = d.customers || {}, nr = d.nextRes || {};
    const date = toIsoDate(d.date);
    if (!date || !d.store) return null;
    return {
        report_date: date,
        store_id: String(d.store),
        staff_name: String(d.staff || ''),
        sales_cash: toInt(s.cash), sales_credit: toInt(s.credit),
        sales_qr: toInt(s.qr), sales_product: toInt(s.product),
        discount_hpb_points: toInt(dc.hpbPoints), discount_hpb_gift: toInt(dc.hpbGift),
        discount_other: toInt(dc.other), refund: toInt(dc.refund),
        new_hpb: toInt(c.newHPB), new_mininai: toInt(c.newMiniNai),
        existing_count: toInt(c.existing), acquaintance: toInt(c.acquaintance),
        next_res_new_hpb: toInt(nr.newHPB), next_res_new_mininai: toInt(nr.newMiniNai),
        next_res_existing: toInt(nr.existing),
        reviews_5star: toInt(d.reviews5Star), blog_updates: toInt(d.blogUpdates),
        sns_updates: toInt(d.snsUpdates),
    };
}

function mapCustomer(c) {
    const { id, store, timestamp, name, nameKana, ...rest } = c || {};
    let answeredAt = null;
    if (timestamp) {
        const t = typeof timestamp === 'number' ? new Date(timestamp) : new Date(String(timestamp));
        if (!isNaN(t)) answeredAt = t.toISOString();
    }
    return {
        store_id: store || null,
        answered_at: answeredAt,
        name: name || null,
        name_kana: nameKana || null,
        details: rest,
    };
}

async function insertBatched(table, rows, size = 400) {
    for (let i = 0; i < rows.length; i += size) {
        await sbRest('POST', `/rest/v1/${table}`, rows.slice(i, i + size));
        process.stdout.write(`\r  ${table}: ${Math.min(i + size, rows.length)}/${rows.length} 件`);
    }
    console.log();
}

(async () => {
    console.log('=== GAS → Supabase 移行 ===');
    console.log(`GAS: ${GAS_URL.slice(0, 60)}…`);
    console.log(`Supabase: ${SB_URL}${DRY ? '（dry-run: 書き込みなし）' : ''}\n`);

    // 1. 売上日報
    console.log('[1/5] 売上日報を取得中…');
    let sales = await gasGet('');
    if (!Array.isArray(sales)) sales = sales?.data || [];
    const reports = sales.map(mapReport).filter(Boolean);
    console.log(`  取得 ${sales.length} 件 → 変換 ${reports.length} 件`);

    // 2. 顧客（カウンセリング）
    console.log('[2/5] 顧客データを取得中…');
    let customersRes = await gasGet('get_customers');
    const customers = (customersRes?.data || (Array.isArray(customersRes) ? customersRes : [])).map(mapCustomer);
    console.log(`  取得 ${customers.length} 件`);

    // 3. 目標・基本給
    console.log('[3/5] 目標・基本給を取得中…');
    const goalsRes = await gasGet('load_goals');
    const goals = goalsRes?.goals || {};
    const salaries = goalsRes?.salaries || {};
    console.log(`  目標 ${Object.keys(goals).length} ヶ月分 / 基本給 ${Object.keys(salaries).length} 店舗分`);

    // 4. 設定（スタッフ名簿・広告費・月締め）
    console.log('[4/5] 設定を取得中…');
    const settingsRes = await gasGet('load_settings');
    const settings = settingsRes?.settings || {};

    // 5. スタッフパスワード（平文 → 移行先でbcryptハッシュ化される）
    console.log('[5/5] パスワードを取得中…');
    const pwRes = await gasGet('load_passwords');
    const rawPw = pwRes?.passwords || {};
    // ネスト形式 {store:{staff:pass}} とフラット形式 {store_staff:pass} の両方に対応
    const flatPw = {};
    Object.entries(rawPw).forEach(([k, v]) => {
        if (v && typeof v === 'object') {
            Object.entries(v).forEach(([staff, pass]) => { flatPw[`${k}_${staff}`] = pass ?? ''; });
        } else {
            flatPw[k] = v ?? '';
        }
    });

    if (DRY) {
        console.log('\n--dry-run のため書き込みをスキップしました。');
        return;
    }

    if (WIPE) {
        console.log('\n--wipe: 既存の daily_reports / customers を削除…');
        await sbRest('DELETE', '/rest/v1/daily_reports?id=gt.0');
        await sbRest('DELETE', '/rest/v1/customers?id=gt.0');
    }

    console.log('\n書き込み中…');
    // 名簿を先に同期（store FK を確実に）
    await sbRpc('api_save_settings', { p: {
        staffRoster: settings.staffRoster || undefined,
        adCosts: settings.adCosts || undefined,
        monthlyClose: settings.monthlyClose || undefined,
    }});
    console.log('  設定(名簿/広告費/月締め): OK');

    await insertBatched('daily_reports', reports);
    await insertBatched('customers', customers);

    await sbRpc('api_save_goals', { p_goals: goals, p_salaries: salaries });
    console.log('  目標・基本給: OK');

    if (Object.keys(flatPw).length) {
        await sbRpc('api_save_passwords', { p: flatPw });
        console.log(`  パスワード: ${Object.keys(flatPw).length} 件をハッシュ化して保存`);
    }

    const health = await sbRpc('api_health');
    console.log(`\n✅ 移行完了！ Supabase上: 日報 ${health.reports} 件 / 顧客 ${health.customers} 件`);
    console.log('ダッシュボードの 設定タブ > バックエンド設定 で Supabase に切り替えてください。');
})().catch(e => { console.error('\n❌ 移行エラー:', e.message); process.exit(1); });
