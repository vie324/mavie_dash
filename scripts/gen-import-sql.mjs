// ============================================================
// 既存GASデータ → Supabase投入用SQL生成（貼り付け移行用）
//   現行GASのJSON出力を取得し、Supabase SQL Editor に貼るだけの
//   INSERT文（.sqlファイル）を ./sql-out/ に生成する。
//   ※ 生成物は個人情報を含むため Git にコミットしないこと（.gitignore 済み）。
//
// 使い方:
//   GAS_URL=https://script.google.com/macros/s/XXX/exec node scripts/gen-import-sql.mjs
//   （GAS_URL 未指定時は下記デフォルトを使用）
//
// 出力:
//   sql-out/import_1_sales_config.sql        … 売上 + 目標/基本給/名簿
//   sql-out/import_2_customers_partN.sql ... … 顧客（1ファイル1500行に分割）
//
// 投入順:
//   1) supabase/schema.sql を実行
//   2) import_1 → import_2 → ... の順に SQL Editor で実行
//   3) 件数照合:  select count(*) from sales_reports;  /  customers;
// ============================================================
import fs from 'fs';

const GAS = process.env.GAS_URL ||
  'https://script.google.com/macros/s/AKfycbwrz7LgQb2uH9VTtmalZxIEcJnHc-Ae53UNMnDi0MM5eLdP7XmZKOlDPTaOL5pmsFwf/exec';
const OUT = 'sql-out';

const getJSON = async (qs) => { const r = await fetch(GAS + qs, { redirect: 'follow' }); return r.json(); };

// ---- SQLリテラル ヘルパー ----
const q = (v) => {
  if (v === null || v === undefined) return 'NULL';
  // NUL/制御文字を除去（Postgres はテキストに U+0000 を許可しない）
  const s = String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim();
  return s === '' ? 'NULL' : "'" + s.replace(/'/g, "''") + "'";
};
const n = (v) => { const x = parseInt(v, 10); return isNaN(x) ? 0 : x; };
const STORES = ['chiba', 'honatsugi', 'yamato'];
const storeLit = (v) => STORES.includes(v) ? "'" + v + "'" : null;
const dq = (v) => {
  if (!v) return 'NULL';
  const m = String(v).trim().match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (!m) return 'NULL';
  const Y = +m[1], M = +m[2], D = +m[3];
  if (M < 1 || M > 12 || D < 1 || D > 31) return 'NULL';
  const dt = new Date(Y, M - 1, D);
  if (dt.getFullYear() !== Y || dt.getMonth() !== M - 1 || dt.getDate() !== D) return 'NULL';
  return "'" + m[1] + '-' + String(M).padStart(2, '0') + '-' + String(D).padStart(2, '0') + "'";
};
const tq = (epoch) => {
  if (!epoch || typeof epoch !== 'number') return 'NULL';
  const dt = new Date(epoch);
  return isNaN(dt.getTime()) ? 'NULL' : "'" + dt.toISOString() + "'";
};
const jq = (obj) => "'" + JSON.stringify(obj ?? null).replace(/'/g, "''") + "'::jsonb";

const stats = { salesIn: 0, salesSkip: 0, custIn: 0, custSkip: 0, bdayNull: 0 };

const sales = await getJSON('?action=get_data&nocache=true');
const custResp = await getJSON('?action=get_customers&nocache=true');
const customers = custResp.data || [];
const goalsResp = await getJSON('?action=load_goals');
const settingsResp = await getJSON('?action=load_settings');

const salesCols = '(report_date, store, staff, sales_cash, sales_credit, sales_qr, sales_product, discount_hpb_points, discount_hpb_gift, discount_other, discount_refund, cust_new_hpb, cust_new_mininai, cust_existing, cust_acquaintance, nextres_new_hpb, nextres_new_mininai, nextres_existing, reviews_5star, blog_updates, sns_updates)';
const salesTuples = [];
for (const r of sales) {
  const d = dq(r.date), st = storeLit(r.store), stf = q(r.staff);
  if (d === 'NULL' || !st || stf === 'NULL') { stats.salesSkip++; continue; }
  const s = r.sales || {}, dc = r.discounts || {}, c = r.customers || {}, nx = r.nextRes || {};
  salesTuples.push(`(${d},${st},${stf},${n(s.cash)},${n(s.credit)},${n(s.qr)},${n(s.product)},${n(dc.hpbPoints)},${n(dc.hpbGift)},${n(dc.other)},${n(dc.refund)},${n(c.newHPB)},${n(c.newMiniNai)},${n(c.existing)},${n(c.acquaintance)},${n(nx.newHPB)},${n(nx.newMiniNai)},${n(nx.existing)},${n(r.reviews5Star)},${n(r.blogUpdates)},${n(r.snsUpdates)})`);
  stats.salesIn++;
}

const custCols = '(store, submitted_at, name, name_kana, address, phone, birthday, job, sns_ok, visit_reason, from_other_salon, dissatisfaction, allergy, eyebrow_frequency, eyebrow_last_care, eyebrow_concern, eyebrow_design, eyebrow_design_image, eyebrow_impression, eyebrow_trouble, lash_frequency, lash_design, lash_design_image, lash_eye_look, lash_contact, lash_trouble, agreement)';
const custTuples = [];
for (const c of customers) {
  const st = storeLit(c.store);
  if (!st) { stats.custSkip++; continue; }
  const bd = dq(c.birthday); if (c.birthday && bd === 'NULL') stats.bdayNull++;
  custTuples.push(`(${st},${tq(c.timestamp)},${q(c.name)},${q(c.nameKana)},${q(c.address)},${q(c.phone)},${bd},${q(c.job)},${q(c.snsOk)},${q(c.visitReason)},${q(c.fromOtherSalon)},${q(c.dissatisfaction)},${q(c.allergy)},${q(c.eyebrowFrequency)},${q(c.eyebrowLastCare)},${q(c.eyebrowConcern)},${q(c.eyebrowDesign)},${q(c.eyebrowDesignImage)},${q(c.eyebrowImpression)},${q(c.eyebrowTrouble)},${q(c.lashFrequency)},${q(c.lashDesign)},${q(c.lashDesignImage)},${q(c.lashEyeLook)},${q(c.lashContact)},${q(c.lashTrouble)},${q(c.agreement)})`);
  stats.custIn++;
}

const batchInsert = (table, cols, tuples, size = 500) => {
  let sql = '';
  for (let i = 0; i < tuples.length; i += size) {
    sql += `insert into ${table}\n${cols} values\n` + tuples.slice(i, i + size).join(',\n') + ';\n\n';
  }
  return sql;
};

fs.mkdirSync(OUT, { recursive: true });

let f1 = '-- Mavie データ移行 (1): 売上 ' + stats.salesIn + '件 + 目標/基本給/名簿\n';
f1 += '-- 前提: 先に supabase/schema.sql を実行。再投入時: truncate table public.sales_reports restart identity;\n\n';
f1 += batchInsert('public.sales_reports', salesCols, salesTuples);
f1 += 'insert into public.app_config (key, value) values\n';
f1 += `('goals_data', ${jq(goalsResp.goals || {})}),\n`;
f1 += `('salaries_data', ${jq(goalsResp.salaries || {})}),\n`;
f1 += `('staff_roster', ${jq(settingsResp.settings ? settingsResp.settings.staffRoster : null)})\n`;
f1 += 'on conflict (key) do update set value = excluded.value, updated_at = now();\n';
fs.writeFileSync(`${OUT}/import_1_sales_config.sql`, f1);

const perFile = 1500;
let part = 0;
for (let i = 0; i < custTuples.length; i += perFile) {
  part++;
  const slice = custTuples.slice(i, i + perFile);
  let head = `-- Mavie データ移行 顧客 part ${part}: ${slice.length}件 [全${stats.custIn}件]\n`;
  if (part === 1) head += '-- 再投入時: truncate table public.customers restart identity;\n';
  head += '\n';
  fs.writeFileSync(`${OUT}/import_${part + 1}_customers_part${part}.sql`, head + batchInsert('public.customers', custCols, slice));
}

console.log('生成完了:', JSON.stringify(stats));
console.log(`出力先: ${OUT}/  （schema.sql の後に import_1 → import_${part + 1} の順で実行）`);
