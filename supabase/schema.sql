-- vie ダッシュボード: サーバー保存用のキーバリューテーブル
-- Supabase の SQL Editor でこのファイルの内容をそのまま実行してください（何度実行しても安全）。
--
-- 保存されるもの（すべてサーバー側の関数だけが service_role キーで読み書き）:
--   vie:manual:<YYYY-MM>  手入力データ（次回予約数・ブログ/SNS/口コミ・物販・広告費・入金突合）
--   vie:shift:<YYYY-MM>   シフト希望休の申請・割当・承認
--   vie:shiftconfig       シフトルール
--   vie:accounts          店長/スタッフのパスワード（scrypt ハッシュ）

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
