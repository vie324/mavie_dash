// ============================================================
// 既存データ移行スクリプト（一度きり実行）
//   現行 GAS の JSON 出力を読み取り、Supabase へ投入する。
//   GAS が店舗名・スタッフ名・数値を正規化済みなので、それを再利用する。
//
// 使い方:
//   1) npm install
//   2) 環境変数を設定（.env.example 参照）
//        SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY / GAS_URL
//   3) node scripts/migrate.mjs           （確認のみ。投入はしない）
//      node scripts/migrate.mjs --apply   （実際に投入）
//      node scripts/migrate.mjs --apply --truncate （既存を消してから投入）
// ============================================================
import { createClient } from '@supabase/supabase-js';

const { SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, GAS_URL } = process.env;
if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY || !GAS_URL) {
  console.error('環境変数 SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY / GAS_URL を設定してください');
  process.exit(1);
}
const APPLY = process.argv.includes('--apply');
const TRUNCATE = process.argv.includes('--truncate');
const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, { auth: { persistSession: false } });

const toISO = (d) => {                 // '2026/1/5' → '2026-01-05'
  if (!d) return null;
  const [y, m, day] = String(d).split('/');
  if (!y || !m || !day) return null;
  return `${y}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
};
const get = async (action) => {
  const url = action ? `${GAS_URL}?action=${action}&nocache=true` : `${GAS_URL}?nocache=true`;
  const r = await fetch(url);
  if (!r.ok) throw new Error(`GAS ${action} HTTP ${r.status}`);
  return r.json();
};

async function main() {
  console.log(APPLY ? '=== 本番投入モード ===' : '=== 確認モード（--apply で投入）===');

  // ---- 売上 ----
  const sales = await get('get_data');
  const salesRows = (Array.isArray(sales) ? sales : []).map(r => ({
    report_date: toISO(r.date), store: r.store, staff: r.staff,
    sales_cash: r.sales.cash, sales_credit: r.sales.credit, sales_qr: r.sales.qr, sales_product: r.sales.product,
    discount_hpb_points: r.discounts.hpbPoints, discount_hpb_gift: r.discounts.hpbGift,
    discount_other: r.discounts.other, discount_refund: r.discounts.refund,
    cust_new_hpb: r.customers.newHPB, cust_new_mininai: r.customers.newMiniNai,
    cust_existing: r.customers.existing, cust_acquaintance: r.customers.acquaintance,
    nextres_new_hpb: r.nextRes.newHPB, nextres_new_mininai: r.nextRes.newMiniNai, nextres_existing: r.nextRes.existing,
    reviews_5star: r.reviews5Star, blog_updates: r.blogUpdates, sns_updates: r.snsUpdates
  })).filter(r => r.report_date);

  // ---- 顧客 ----
  const cust = await get('get_customers');
  const custRows = ((cust && cust.data) || []).map(c => ({
    store: c.store, submitted_at: c.timestamp ? new Date(c.timestamp).toISOString() : null,
    name: c.name, name_kana: c.nameKana, address: c.address, phone: c.phone,
    birthday: toISO(c.birthday), job: c.job, sns_ok: c.snsOk, visit_reason: c.visitReason,
    from_other_salon: c.fromOtherSalon, dissatisfaction: c.dissatisfaction, allergy: c.allergy,
    eyebrow_frequency: c.eyebrowFrequency, eyebrow_last_care: c.eyebrowLastCare, eyebrow_concern: c.eyebrowConcern,
    eyebrow_design: c.eyebrowDesign, eyebrow_design_image: c.eyebrowDesignImage,
    eyebrow_impression: c.eyebrowImpression, eyebrow_trouble: c.eyebrowTrouble,
    lash_frequency: c.lashFrequency, lash_design: c.lashDesign, lash_design_image: c.lashDesignImage,
    lash_eye_look: c.lashEyeLook, lash_contact: c.lashContact, lash_trouble: c.lashTrouble,
    agreement: c.agreement
  }));

  // ---- 設定・目標 ----
  const goals = await get('load_goals');
  const settings = await get('load_settings');

  console.log(`売上: ${salesRows.length} 件 / 顧客: ${custRows.length} 件 を移行対象として検出`);
  if (!APPLY) {
    console.log('サンプル(売上1件):', JSON.stringify(salesRows[0], null, 2));
    console.log('サンプル(顧客1件):', JSON.stringify(custRows[0], null, 2));
    console.log('→ 問題なければ --apply を付けて再実行してください');
    return;
  }

  if (TRUNCATE) {
    console.log('既存データを削除中...');
    await supabase.from('sales_reports').delete().neq('id', 0);
    await supabase.from('customers').delete().neq('id', 0);
  }

  // バッチ投入（500件ずつ）
  const insertAll = async (table, rows) => {
    for (let i = 0; i < rows.length; i += 500) {
      const { error } = await supabase.from(table).insert(rows.slice(i, i + 500));
      if (error) throw error;
    }
  };
  await insertAll('sales_reports', salesRows);
  await insertAll('customers', custRows);

  await supabase.from('app_config').upsert([
    { key: 'goals_data', value: goals.goals || {} },
    { key: 'salaries_data', value: goals.salaries || {} },
    { key: 'staff_roster', value: settings.settings?.staffRoster || null }
  ], { onConflict: 'key' });

  // 件数照合
  const { count: salesCount } = await supabase.from('sales_reports').select('*', { count: 'exact', head: true });
  const { count: custCount } = await supabase.from('customers').select('*', { count: 'exact', head: true });
  console.log(`投入完了 → sales_reports: ${salesCount} 件 / customers: ${custCount} 件`);
}

main().catch(e => { console.error(e); process.exit(1); });
