-- ============================================================
-- Mavie Dashboard - Supabase スキーマ
-- 「フォームは残してデータだけ移行」用
-- Supabase の SQL Editor に貼り付けて実行してください。
-- ============================================================

-- ---------- 売上日報 ----------
create table if not exists public.sales_reports (
  id                   bigint generated always as identity primary key,
  report_date          date    not null,                 -- 出勤日（旧B列）
  store                text    not null check (store in ('chiba','honatsugi','yamato')),
  staff                text    not null,                  -- 小文字英字（kiki 等）
  sales_cash           integer not null default 0,
  sales_credit         integer not null default 0,
  sales_qr             integer not null default 0,
  sales_product        integer not null default 0,        -- 物販（内数）
  discount_hpb_points  integer not null default 0,
  discount_hpb_gift    integer not null default 0,
  discount_other       integer not null default 0,
  discount_refund      integer not null default 0,
  cust_new_hpb         integer not null default 0,
  cust_new_mininai     integer not null default 0,
  cust_existing        integer not null default 0,
  cust_acquaintance    integer not null default 0,
  nextres_new_hpb      integer not null default 0,
  nextres_new_mininai  integer not null default 0,
  nextres_existing     integer not null default 0,
  reviews_5star        integer not null default 0,
  blog_updates         integer not null default 0,
  sns_updates          integer not null default 0,
  created_at           timestamptz not null default now()
);
create index if not exists idx_sales_reports_date  on public.sales_reports (report_date);
create index if not exists idx_sales_reports_store on public.sales_reports (store, staff);

-- ---------- 顧客カウンセリング ----------
create table if not exists public.customers (
  id                    bigint generated always as identity primary key,
  store                 text not null check (store in ('chiba','honatsugi','yamato')),
  submitted_at          timestamptz,            -- フォーム回答日時
  name                  text,
  name_kana             text,
  address               text,
  phone                 text,
  birthday              date,
  job                   text,
  sns_ok                text,
  visit_reason          text,
  from_other_salon      text,
  dissatisfaction       text,
  allergy               text,
  eyebrow_frequency     text,
  eyebrow_last_care     text,
  eyebrow_concern       text,
  eyebrow_design        text,
  eyebrow_design_image  text,
  eyebrow_impression    text,
  eyebrow_trouble       text,
  lash_frequency        text,
  lash_design           text,
  lash_design_image     text,
  lash_eye_look         text,
  lash_contact          text,
  lash_trouble          text,
  agreement             text,
  created_at            timestamptz not null default now()
);
create index if not exists idx_customers_store on public.customers (store);
create index if not exists idx_customers_ts    on public.customers (submitted_at);

-- ---------- 設定・目標・基本給・パスワード（キーバリュー） ----------
-- 既存は JSON 文字列で保持しているため、まずは JSON のまま移すのが最短。
--   goals_data    : 目標（月別・店舗別・スタッフ別）
--   salaries_data : 基本給
--   staff_roster  : スタッフ名簿
--   admin_password: 管理ページのパスワード（文字列）
--   gemini_api_key: Gemini キー（※将来サーバー側へ移行推奨）
--   passwords_data: スタッフログインパスワード { "chiba_kiki": "xxxx" }
create table if not exists public.app_config (
  key        text primary key,
  value      jsonb,
  updated_at timestamptz not null default now()
);

-- ============================================================
-- RLS（行レベルセキュリティ）
-- 互換API は service_role キーで動作し RLS を **バイパス** します。
-- RLS を有効化し公開ポリシーを作らないことで、anon キー
-- （ブラウザ）からの直接アクセスを既定で拒否します。
-- ============================================================
alter table public.sales_reports enable row level security;
alter table public.customers     enable row level security;
alter table public.app_config    enable row level security;
-- ※ ポリシーは意図的に作成しません（= service_role 経由のみ許可）。
--   将来 supabase-js 直結 / Supabase Auth に進む際にポリシーを追加します。
