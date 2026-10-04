-- vie ダッシュボード: サーバー保存用のキーバリューテーブル
-- Supabase の SQL Editor でこのファイルの内容をそのまま実行してください（何度実行しても安全）。
--
-- 保存されるもの（すべてサーバー側の関数だけが service_role キーで読み書き）:
--   vie:manual:<YYYY-MM>  手入力データ（次回予約数〈媒体別 × 新規/2回目以降〉・ブログ/SNS/口コミ・物販・広告費・入金突合）
--   vie:shift:<YYYY-MM>   シフト希望休の申請・割当・承認
--   vie:shiftconfig       シフトルール
--   vie:accounts          店長/スタッフのパスワード（scrypt ハッシュ。発行はダッシュボードの設定タブから。SQLでは作らない）
--
-- 確認・集計用の読み取り専用SQL（アカウント一覧・スタッフ別/媒体別の次回予約）は supabase/queries.sql にあります。

create table if not exists public.vie_kv (
    key        text primary key,
    value      jsonb not null,
    updated_at timestamptz not null default now()
);

comment on table public.vie_kv is 'vie dashboard: 手入力・シフト・アカウントのサーバー保存（service_role 専用）';

-- ブラウザ用キー（anon / authenticated）からは一切アクセスできないようにする。
-- RLS を有効にしてポリシーを作らない = 権限なし。service_role は RLS をバイパスする。
alter table public.vie_kv enable row level security;
revoke all on table public.vie_kv from anon, authenticated;

-- 補足: 保存時の競合防止のため "lock:<キー>" という行が一時的に作られます（数秒で自動削除・期限切れは上書き）。
-- Table Editor で見かけても消す必要はありません。

-- 領収書・レシートの写真（出納帳）: 非公開バケット。ダッシュボードの API が service_role で保存し、
-- 閲覧は API が発行する短時間の署名付きURL経由のみ。初回保存時に API が自動作成するので、ここでは無くても動きます。
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('vie-receipts', 'vie-receipts', false, 4194304, array['image/jpeg', 'image/png', 'image/webp'])
on conflict (id) do nothing;
