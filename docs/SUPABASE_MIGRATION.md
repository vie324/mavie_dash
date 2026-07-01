# GAS → Supabase 完全移行ガイド

このダッシュボードは **バックエンド抽象化レイヤー**(`assets/js/backend.js`)を備えており、
設定タブのスイッチひとつで GAS(スプレッドシート) と Supabase を切り替えられます。
移行は以下の4ステップ。**切替後も GAS に戻せばいつでもロールバック可能**です
(移行後に GAS 側へ入力した分は再移行が必要)。

```
┌─────────┐   ①SQL適用    ┌──────────┐   ③スクリプトで   ┌──────────┐
│ Supabase │ ◀──────────── │ このリポジトリ │   データ移行      │   GAS    │
│ (Postgres)│               │ supabase/ 配下 │ ◀──────────────── │(スプレッドシート)│
└─────────┘               └──────────┘                    └──────────┘
        ▲ ④設定タブで切替
┌─────────┐
│ ダッシュボード │ … コード変更なしで両対応(apiFetchゲートウェイ)
└─────────┘
```

---

## ステップ1: Supabase プロジェクト作成

1. https://supabase.com → **New Project**(無料プランでOK)
2. リージョンは **Northeast Asia (Tokyo)** を推奨
3. 作成後、**Settings > API** で以下を控える
   - `Project URL`(例: `https://abcd1234.supabase.co`)
   - `anon` `public` キー(フロント用)
   - `service_role` キー(移行スクリプト専用。**絶対に公開しない・フロントに入れない**)

## ステップ2: スキーマ適用（★1回で完了）

> **一番かんたん**: [`supabase/setup.sql`](../supabase/setup.sql) を **SQL Editor に貼り付けて Run するだけ**。
> スキーマ・API・面談機能・初期データ・権限・Realtime を1回で全部セットアップします（冪等）。

（旧: 個別に `migrations/0001_init.sql` → `migrations/0002_reviews.sql` → `seed.sql` を順に適用してもOK。CLI派は `supabase db push`。）

<details><summary>ステップ2（詳細）</summary>

1. Supabase Dashboard → **SQL Editor** → New query
2. [`supabase/migrations/0001_init.sql`](../supabase/migrations/0001_init.sql) の内容を貼り付けて **Run**
3. 続けて [`supabase/seed.sql`](../supabase/seed.sql) を実行(店舗・既定スタッフの登録)

> Supabase CLI 利用者は `supabase db push` でも可。

これで以下が作成されます:

| テーブル | 内容 |
|---------|------|
| `stores` / `staff` | 店舗・スタッフ(基本給、**bcryptハッシュ化された**パスワード) |
| `daily_reports` | 売上日報(現GASのA〜X列を正規化) |
| `goals` | 月別×店舗×スタッフの目標(jsonb) |
| `customers` | カウンセリング回答(可変項目はjsonb) |
| `reviews` | 月次面談・振り返り(自己評価/振り返り/AI評価) |
| `app_settings` | 広告費 / 月締め確定 / 管理パスワードハッシュ |

API は `api_*` という Postgres 関数(RPC)として公開され、**GASのactionと1:1対応・レスポンス形状も互換**です
(`get_data`→`api_get_sales`、`save_goals`→`api_save_goals`、`save_review`→`api_save_review` など)。

</details>

## ステップ3: データ移行

ローカルPC(Node.js 18+)で:

```bash
GAS_URL="https://script.google.com/macros/s/xxxx/exec" \
SUPABASE_URL="https://abcd1234.supabase.co" \
SUPABASE_SERVICE_KEY="(service_roleキー)" \
node scripts/migrate-from-gas.mjs
```

- 移行対象: 売上日報 / 顧客(カウンセリング) / 目標・基本給 / スタッフ名簿 / 広告費 / 月締め / スタッフパスワード(**平文→bcryptハッシュ化**)
- 再実行する場合は `--wipe` を付けると日報・顧客を入れ替え
- `--dry-run` で取得・変換のみ確認可能

## ステップ4: ダッシュボードを切り替え

1. ダッシュボード → **設定タブ → バックエンド設定**
2. `Supabase Project URL` と `anon key` を入力 → **Supabase接続テスト**(件数が表示されればOK)
3. 「**Supabase**」を選択 → **保存して切り替え**(自動で再読み込み)

以上で完了。全端末で同じ設定が必要です(スタッフのスマホ含む)。
Vercel で `config.js` を注入した場合は、各端末での入力は不要になります(下記)。

---

## Vercel デプロイと環境変数

このアプリは**静的サイト**なので、Vercel には「Framework Preset: Other / そのままデプロイ」でOKです。
接続先は次の2通りで設定できます。

- **(推奨) config.js を環境変数から自動生成**: `vercel.json` の `buildCommand`(`node scripts/gen-config.mjs`)が、Vercelの環境変数から `config.js` を生成し、全端末の既定接続先になります。スタッフは何も設定せず使えます。
- **UIで各自入力**: 環境変数を設定しない場合、従来どおり設定タブで各端末が入力(localStorage)。

### 必要な環境変数 一覧

| 変数名 | 設定場所 | 必須 | 用途 / 備考 |
|--------|---------|:---:|------|
| `SUPABASE_URL` | Vercel (Build) | ✔ | `https://xxxx.supabase.co`。config.js に注入 |
| `SUPABASE_ANON_KEY` | Vercel (Build) | ✔ | anon(public)キー。**公開可**(RLS前提)。config.js に注入 |
| `BACKEND_MODE` | Vercel (Build) | – | `supabase`(既定) / `gas`。省略時、URL・keyがあれば supabase |
| `GEMINI_API_KEY` | Supabase Secrets | ✔※ | AIコーチ・AI面談評価。`supabase secrets set GEMINI_API_KEY=...` |
| `GEMINI_MODEL` | Supabase Secrets | – | 既定 `gemini-2.0-flash` |
| `GAS_URL` | ローカル(移行時) | ✔※ | 移行スクリプトのGAS取得元。`scripts/migrate-from-gas.mjs` |
| `SUPABASE_SERVICE_KEY` | ローカル(移行時) | ✔※ | service_roleキー。**絶対に公開・コミットしない** |

※ `GEMINI_*` はAI機能を使う場合のみ。`GAS_URL`/`SUPABASE_SERVICE_KEY` は移行スクリプト実行時のみ。
サンプルは [`.env.example`](../.env.example) / フロント既定は [`config.sample.js`](../config.sample.js)。

> **キーの取り扱い**: `anon key` は公開されても良いキー(RLSが守る)。`service_role` と `GEMINI_API_KEY` は**秘匿**。前者は移行スクリプトのローカル実行のみ、後者はSupabase Secrets(Edge Function)に置き、ブラウザには出しません。

### Vercel 設定手順(要約)

1. GitHub リポジトリを Vercel にインポート(Framework: Other)
2. **Settings → Environment Variables** に上表の Vercel(Build) 3変数を登録
3. Deploy(`buildCommand` が `config.js` を生成)
4. 生成された `config.js` は `.gitignore` 済み(リポジトリには入りません)

---

## (任意) Gemini AI を Edge Function 中継にする

現在 Gemini API キーはブラウザの localStorage に保存されています。
Supabase 移行後は Edge Function 中継にすることで**キーを完全に秘匿**できます。

```bash
supabase functions deploy gemini-advice
supabase secrets set GEMINI_API_KEY=AIza....
```

デプロイ済みなら、Supabaseモードの AIアドバイス/AIコーチは自動で中継を使います
(未デプロイ時はローカルキーにフォールバック)。

## (任意) Realtime 即時反映

スキーマ適用時に `daily_reports` が Realtime 対象になっています。
Supabaseモードでは日報の追加・修正が**数秒で全端末に反映**されます
(従来の5分ポーリングも併用)。
Dashboard → Database → Replication で `supabase_realtime` に
`daily_reports` が含まれているか確認できます。

---

## GASとの対応表

| GAS action | Supabase RPC | 備考 |
|-----------|--------------|------|
| `get_data` | `api_get_sales()` | 素の配列を返す(互換) |
| `add_record` | `api_add_record(p)` | 日報入力 |
| `update` | `api_update_rows(p)` | データ編集タブ |
| `get_customers(_today/_by_store)` | `api_get_customers(p_store,p_today)` | |
| `load_goals` / `save_goals` | `api_load_goals()` / `api_save_goals(...)` | 基本給は staff テーブルへ |
| `load_settings` / `save_settings` | `api_load_settings()` / `api_save_settings(p)` | 名簿は stores/staff へ同期 |
| `load_passwords` / `save_passwords` | `api_load_passwords()` / `api_save_passwords(p)` | 平文は返らない(`__SET__`) |
| `verify_password` / `verify_session` | `api_verify_password(...)` + クライアントセッション | |
| `get_reviews` | `api_get_reviews()` | 面談: ネスト構造 {ym:{store:{staff}}} |
| `save_review` | `api_save_review(p)` | スタッフの振り返り保存(AI評価は保持) |
| `save_review_ai` | `api_save_review_ai(...)` | 管理者のAI評価保存(未提出でも可) |

## セキュリティ上の改善点と既知の制約

**改善:**
- パスワードが bcrypt ハッシュ保存になり、**平文がどこにも残らない**(GASはシート・localStorageに平文)
- スタッフ認証はサーバー側照合(`api_verify_password`)
- Gemini APIキーを Edge Function Secret に移せる
- テーブルは RLS で直接アクセス遮断(書き込みは全て RPC 経由)

**制約(GASと同等レベル、Phase 4 で改善予定):**
- anon キーを知っていれば RPC を呼べる(GASのURLを知っていれば呼べるのと同等)
- 本格的なユーザー認証は Supabase Auth への移行が必要(ロードマップ I8)

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| 接続テストで「migrations未適用の可能性」 | ステップ2のSQLを実行したか確認 |
| 移行スクリプトが 401/403 | `SUPABASE_SERVICE_KEY` に service_role を指定しているか確認 |
| 切替後にデータが空 | ステップ3の移行を実行したか / `api_health` の件数を確認 |
| スタッフページにログインできない | パスワード移行済みか確認。`api_save_passwords` で再設定可能(設定タブからも可) |
| 元に戻したい | 設定タブで「GAS」を選んで保存(データは消えません) |
