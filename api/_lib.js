// ============================================================
// 互換API 共通ライブラリ
// 現行 GAS(Code.gs) と同じレスポンス形を Supabase で再現する。
// ※ ファイル名先頭が "_" のため Vercel のルートにはならない（共有モジュール）。
// ============================================================
import { createClient } from '@supabase/supabase-js';
import crypto from 'crypto';

// service_role キーで動作（RLS をバイパス）。絶対にフロントに出さない。
export const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY,
  { auth: { persistSession: false } }
);

const SESSION_SECRET = process.env.SESSION_SECRET || 'change-me';
const SESSION_TTL_MS = 24 * 60 * 60 * 1000; // 24時間（GAS と同じ）

// ---------- レスポンス補助 ----------
export function cors(res) {
  res.setHeader('Access-Control-Allow-Origin', '*'); // 本番では自ドメインに限定推奨
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Accept');
}
export function json(res, body, status = 200) {
  res.status(status).setHeader('Content-Type', 'application/json; charset=utf-8');
  res.end(JSON.stringify(body));
}

// ---------- リクエストボディの取得 ----------
// フロントは Content-Type: text/plain で JSON 文字列を送ってくる（GAS の名残）。
export async function readBody(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  if (typeof req.body === 'string' && req.body) {
    try { return JSON.parse(req.body); } catch { return {}; }
  }
  // フォールバック: 生ストリームを読む
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const raw = Buffer.concat(chunks).toString('utf8');
  try { return raw ? JSON.parse(raw) : {}; } catch { return {}; }
}

// ---------- 日付整形（GAS は 'yyyy/M/d' / Asia/Tokyo） ----------
export function dateOnlyToSlash(v) {           // 'YYYY-MM-DD' → 'yyyy/M/d'
  if (!v) return '';
  const s = String(v).slice(0, 10);
  const [y, m, d] = s.split('-');
  if (!y || !m || !d) return String(v);
  return `${Number(y)}/${Number(m)}/${Number(d)}`;
}
export function tsToSlash(v) {                  // timestamptz/epoch → 'yyyy/M/d'(JST)
  if (!v) return '';
  const date = typeof v === 'number' ? new Date(v) : new Date(v);
  if (isNaN(date.getTime())) return '';
  return date.toLocaleDateString('ja-JP', { timeZone: 'Asia/Tokyo' }); // → '2026/1/5'
}
export function storeName(store) {
  return store === 'chiba' ? '千葉店'
    : store === 'honatsugi' ? '本厚木店'
    : store === 'yamato' ? '大和店' : store;
}

// ---------- 売上: DB行 → API形（現行フロントが期待する形） ----------
export function salesRowToApi(r) {
  return {
    id: r.id,
    date: dateOnlyToSlash(r.report_date),
    store: r.store,
    storeName: storeName(r.store),
    staff: r.staff,
    sales: { cash: r.sales_cash, credit: r.sales_credit, qr: r.sales_qr, product: r.sales_product },
    discounts: { hpbPoints: r.discount_hpb_points, hpbGift: r.discount_hpb_gift, other: r.discount_other, refund: r.discount_refund },
    customers: { newHPB: r.cust_new_hpb, newMiniNai: r.cust_new_mininai, existing: r.cust_existing, acquaintance: r.cust_acquaintance },
    nextRes: { newHPB: r.nextres_new_hpb, newMiniNai: r.nextres_new_mininai, existing: r.nextres_existing },
    reviews5Star: r.reviews_5star, blogUpdates: r.blog_updates, snsUpdates: r.sns_updates
  };
}

// ---------- 売上: API形 → DB更新列（POST update 用） ----------
export function salesApiToCols(row) {
  const s = row.sales || {}, d = row.discounts || {}, c = row.customers || {}, n = row.nextRes || {};
  return {
    sales_cash: s.cash || 0, sales_credit: s.credit || 0, sales_qr: s.qr || 0, sales_product: s.product || 0,
    discount_hpb_points: d.hpbPoints || 0, discount_hpb_gift: d.hpbGift || 0, discount_other: d.other || 0, discount_refund: d.refund || 0,
    cust_new_hpb: c.newHPB || 0, cust_new_mininai: c.newMiniNai || 0, cust_existing: c.existing || 0, cust_acquaintance: c.acquaintance || 0,
    nextres_new_hpb: n.newHPB || 0, nextres_new_mininai: n.newMiniNai || 0, nextres_existing: n.existing || 0,
    reviews_5star: row.reviews5Star || 0, blog_updates: row.blogUpdates || 0, sns_updates: row.snsUpdates || 0
  };
}

// ---------- 顧客: DB行 → API形 ----------
export function customerRowToApi(c) {
  const ts = c.submitted_at ? new Date(c.submitted_at).getTime() : 0;
  return {
    id: `${c.store}_${c.id}`,
    store: c.store,
    storeName: storeName(c.store),
    date: c.submitted_at ? tsToSlash(c.submitted_at) : '',
    timestamp: ts,
    name: c.name || '', nameKana: c.name_kana || '', address: c.address || '', phone: c.phone || '',
    birthday: c.birthday ? dateOnlyToSlash(c.birthday) : '', job: c.job || '', snsOk: c.sns_ok || '',
    visitReason: c.visit_reason || '', fromOtherSalon: c.from_other_salon || '', dissatisfaction: c.dissatisfaction || '', allergy: c.allergy || '',
    eyebrowFrequency: c.eyebrow_frequency || '', eyebrowLastCare: c.eyebrow_last_care || '', eyebrowConcern: c.eyebrow_concern || '',
    eyebrowDesign: c.eyebrow_design || '', eyebrowDesignImage: c.eyebrow_design_image || '', eyebrowImpression: c.eyebrow_impression || '', eyebrowTrouble: c.eyebrow_trouble || '',
    lashFrequency: c.lash_frequency || '', lashDesign: c.lash_design || '', lashDesignImage: c.lash_design_image || '',
    lashEyeLook: c.lash_eye_look || '', lashContact: c.lash_contact || '', lashTrouble: c.lash_trouble || '', agreement: c.agreement || ''
  };
}

// ---------- app_config 読み書き ----------
export async function getConfig(keys) {
  const { data, error } = await supabase.from('app_config').select('key,value').in('key', keys);
  if (error) throw error;
  const m = {};
  (data || []).forEach(r => { m[r.key] = r.value; });
  return m;
}
export async function setConfig(key, value) {
  const { error } = await supabase.from('app_config')
    .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: 'key' });
  if (error) throw error;
}

// ---------- セッショントークン（ステートレス／GAS 挙動を再現） ----------
export function signToken(payload) {
  const body = Buffer.from(JSON.stringify({ ...payload, exp: Date.now() + SESSION_TTL_MS })).toString('base64url');
  const sig = crypto.createHmac('sha256', SESSION_SECRET).update(body).digest('base64url');
  return `${body}.${sig}`;
}
export function verifyTokenStr(token) {
  if (!token || !token.includes('.')) return null;
  const [body, sig] = token.split('.');
  const expect = crypto.createHmac('sha256', SESSION_SECRET).update(body).digest('base64url');
  if (sig !== expect) return null;
  try {
    const payload = JSON.parse(Buffer.from(body, 'base64url').toString('utf8'));
    if (payload.exp && Date.now() > payload.exp) return null;
    return payload;
  } catch { return null; }
}
