# 未設定項目の解消チェックリスト（設定 → 連携状態）

設定タブ「SalonOne 連携状態」で **未設定** になっている項目について、
「今できていないこと」「やること」「誰がやるか（人の手 / Claude Code）」をまとめたものです。

SalonOne 本体（接続状態・ブランド・スキーマ・個人情報を含まない）は設定済みなので、
売上・来店・マーケの閲覧は今のまま使えます。未設定なのは **保存・AI・パスワード** の3系統です。

## 1. 現状と影響

| 項目 | 現状 | できていないこと | 緊急度 |
|---|---|---|---|
| オーナーパスワード `ADMIN_PASSWORD` | 未設定 | **URLを知っていれば誰でも全店舗の売上・給与・設定を閲覧できる**。スタッフにURLを配り始める前に必須 | ◎ 最優先 |
| 日報・目標・シフトの保存 `SUPABASE_URL` / `SUPABASE_SERVICE_ROLE_KEY` | 未設定（この端末のみ） | 日報・月次目標・基本給・入金突合・広告費が **入力した端末のブラウザにしか残らない**（別端末・スタッフと共有されない、端末を変えると消える）。**出納帳・シフト・スタッフアカウントの発行は使えない** | ◎ |
| スタッフパスワード | 未設定 | スタッフ専用URLがパスワードなしで開ける。画面の「スタッフアカウントの発行」は **Supabase が先に必要**（上の項目が未設定の間はボタンを押しても保存できない） | ○（Supabase の後） |
| 店長パスワード `STORE_PASSWORDS` | 未設定 | 店長URL（`?store=◯`）がパスワードなしで開ける。環境変数でも、画面の「スタッフアカウントの発行」の店舗の行からでも設定できる（画面の方が簡単・Supabase が必要） | ○ |
| マネージャーパスワード `MANAGER_PASSWORD` | 未設定 | マネージャーURL（`?mode=manager`）がパスワードなしで開ける。マネージャー役を使わないなら不要 | △ |
| AIアドバイス `GEMINI_API_KEY` | 未設定 | マイダッシュボードのAIコーチ（スタッフ向けアドバイス生成）が出ない。それ以外に影響なし | △ |
| セッション署名 `AUTH_SECRET`（画面には出ない） | 未設定 | SalonOne APIキーから導出して動作中。APIキーを差し替えると全員ログアウトになる。推奨設定 | △ |

## 2. やること（順番どおり）

| # | やること | 人の手でしかできない部分 | Claude Code でできる部分 |
|---|---|---|---|
| 1 | Vercel への接続 | ローカル: `npx vercel login`（ブラウザ認証）。クラウドセッション: Vercel のトークンを環境の設定に登録 | 以降の環境変数設定・再デプロイをすべて実行 |
| 2 | オーナー / マネージャーパスワードと `AUTH_SECRET` | パスワードを決める（または自動生成を控える） | `node scripts/setup-env.mjs --yes --generate-passwords` で設定＋再デプロイ |
| 3 | Supabase プロジェクト | Supabase にサインアップ。アクセストークンを発行して環境に登録（`SUPABASE_ACCESS_TOKEN`） | プロジェクト作成（東京）・`schema.sql` 適用・疎通確認・Vercel への `SUPABASE_URL` / `SUPABASE_SERVICE_ROLE_KEY` 設定・再デプロイ（`scripts/setup-supabase.mjs`） |
| 4 | スタッフ・店長アカウント | 発行されたパスワードを **LINEで各スタッフに送る** | （画面操作のため Claude Code 対象外）設定 → スタッフアカウントの発行 → 「◯名に一括発行」。手順と文面は [STAFF_ROLLOUT.md](STAFF_ROLLOUT.md) |
| 5 | AIアドバイス（任意） | [Google AI Studio](https://aistudio.google.com/) でAPIキーを発行し、環境に `GEMINI_API_KEY` として登録 | `node scripts/setup-env.mjs --yes --gemini-key-env GEMINI_API_KEY` で Vercel に設定 |

人の手が要るのは **ログイン・サインアップ・トークン発行・LINE送付** だけです。それ以外はスクリプト化してあります。

## 3. Claude Code にやらせる手順

### A. 自分のPCの Claude Code で（推奨・最短）

リポジトリのフォルダで Claude Code を開き、先に一度だけ `npx vercel login` を済ませます。あとは次のように頼みます。

```
docs/SETUP_CHECKLIST.md の手順で Vercel の環境変数を設定して。
scripts/setup-env.mjs を --yes --generate-passwords で実行し、
自動生成されたパスワードを教えて。終わったら再デプロイまでやって。
```

Supabase もやる場合は、https://supabase.com/dashboard/account/tokens でトークンを発行し、
ターミナルで `export SUPABASE_ACCESS_TOKEN=...` してから Claude Code を起動して、次のように頼みます。

```
scripts/setup-supabase.mjs で Supabase を用意して。
プロジェクトがまだ無ければ create（東京）、その後 all --write-vercel まで実行して。
```

Gemini は https://aistudio.google.com/ でキーを発行し `export GEMINI_API_KEY=...` してから:

```
scripts/setup-env.mjs --yes --gemini-key-env GEMINI_API_KEY で AIアドバイスを有効にして。
```

### B. claude.ai のクラウドセッション（このセッションのような環境）で

クラウド環境には Vercel / Supabase の認証がないため、環境の設定（セッションのタイトルバーの環境メニュー → Edit）に
次の変数を登録してから新しいセッションを開いてください。**トークンをチャットに貼らないでください。**

| 変数 | 取得先 | 用途 |
|---|---|---|
| `VERCEL_TOKEN` | Vercel → Account Settings → Tokens | 環境変数の設定・再デプロイ |
| `VERCEL_ORG_ID` / `VERCEL_PROJECT_ID` | Vercel プロジェクトの Settings → General（`vercel link` の代わり） | プロジェクトの特定 |
| `SUPABASE_ACCESS_TOKEN` | Supabase → Account → Access Tokens | プロジェクト作成・SQL 適用・キー取得 |
| `GEMINI_API_KEY` | Google AI Studio | AIアドバイス（任意） |

登録後は A と同じ文面で頼めば、クラウドセッションからそのまま実行できます。

## 4. スクリプトの説明

- `scripts/setup-env.mjs` … Vercel の環境変数をまとめて設定して再デプロイ。既存の値は上書きしない（`--force` で上書き）。
  自動生成したパスワードは実行時の出力に一度だけ表示し、どこにも保存しない。`--dry-run` で設定内容だけ確認できる
- `scripts/setup-supabase.mjs` … Supabase Management API でプロジェクト作成（`create`）・`schema.sql` 適用（`schema`）・
  service_role での疎通と anon から読めないこと（RLS）の確認（`verify`）・Vercel への設定（`all --write-vercel`）。
  service_role キーは画面に表示せず Vercel に直接渡す
- どちらも `vercel env add --sensitive` を使うため、設定後は Vercel の画面でも値を読み返せません（上書きのみ）

## 5. 完了の確認

1. ダッシュボードを再読み込みし、オーナーパスワードを求められること
2. 設定 → 連携状態: 「日報・目標・シフトの保存」= **Supabaseに保存（全端末で共有）**、各パスワード = **設定済み**
3. 日報を1件保存し、別端末（スマホ）で同じ値が見えること
4. 旧バージョンで端末に保存していた目標・日報は、サーバーが空のとき初回起動時に自動で引き継がれます（念のため 設定 → エクスポート で控えを取ってから）
