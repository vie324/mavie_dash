# 実行手順書：フォームは残してデータだけ Supabase へ移行（+ Vercel ホスティング）

この手順書は、**Google フォームでの入力はそのまま継続**しつつ、データの保管先と読み書きを **Supabase** に移し、フロントを **Vercel** で本番公開するための具体的な作業ログです。スタッフの入力操作は変わりません。

> 関連ファイル（このリポジトリに同梱）
> - `supabase/schema.sql` … Supabase テーブル定義
> - `scripts/migrate.mjs` … 既存データの初回移行
> - `gas-sync/FormSync.gs` … フォーム回答を Supabase に同期
> - `api/index.js`, `api/_lib.js` … 現行 GAS 互換API（Vercel Functions）
> - `vercel.json`, `package.json`, `.env.example` … デプロイ設定

---

## 全体像

```
[スタッフ] → Google フォーム → スプレッドシート（従来どおり）
                                   │  ← onFormSubmit トリガー
                                   ▼
                              Supabase(PostgreSQL)   ← 本番DB
                                   ▲
                                   │ 互換API（/api, Vercel）
                              ダッシュボード（Vercel 静的ホスティング）
```

ポイント：**入力経路（フォーム）は二重化された状態**になります。スプレッドシートは当面バックアップとして残し、Supabase が本番データになります。

---

## ステップ 1. Supabase プロジェクト作成 & テーブル作成

1. [supabase.com](https://supabase.com/) でプロジェクト作成（リージョン: Tokyo 推奨）。
2. 左メニュー **SQL Editor** を開き、`supabase/schema.sql` の内容を貼り付けて **Run**。
3. **Project Settings → API** で以下を控える：
   - `Project URL`（= `SUPABASE_URL`）
   - `service_role` キー（= `SUPABASE_SERVICE_ROLE_KEY`、**非公開**）

## ステップ 2. 既存データの初回移行

ローカル（PC）で実行します。Node.js 18+ が必要です。

```bash
# リポジトリ直下で
npm install

# 環境変数を用意（.env.example をコピーして実値を記入）
cp .env.example .env
#  → SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY / GAS_URL を埋める

# まず確認（DBには書き込まない）
node scripts/migrate.mjs            # 件数とサンプルを表示

# 問題なければ投入
node scripts/migrate.mjs --apply
```

- 完了時に `sales_reports: NNN 件 / customers: NNN 件` が表示されます。
- 現行 GAS の件数（`?action=get_data` の配列長、`?action=get_customers` の `data.length`）と一致するか確認。
- やり直す場合は `node scripts/migrate.mjs --apply --truncate`（既存を消してから再投入）。

## ステップ 3. フォーム → Supabase 同期トリガーの設置

回答が増えても Supabase に自動で入るようにします。

1. **売上フォーム**に紐づくスプレッドシート → 拡張機能 → Apps Script。
2. `gas-sync/FormSync.gs` を貼り付け、`SUPABASE_URL` を自分のURLに変更。
3. `setSupabaseKey()` の中に `service_role` キーを記入し、**1回だけ実行**（スクリプトプロパティに保存。実行後はコード上のキー文字列を削除）。
4. トリガー（時計アイコン）→ トリガー追加：
   - 関数 `onSalesFormSubmit` / イベント「フォーム送信時」。
5. **顧客フォーム**（千葉・本厚木・大和の各スプレッドシート）でも同様に貼り付け、`onCustomerFormSubmit` をフォーム送信トリガーで登録。
6. テスト：フォームを1件送信し、Supabase の Table Editor に行が増えることを確認。

> 補足：売上フォームの列順は `gas/Code.gs` の `SALES_COLUMNS` と一致している前提です。顧客フォームは項目名のキーワードで自動判定します（`gas/Code.gs:751` の `findCol` と同じ方針）。項目名が大きく異なる場合は `FormSync.gs` の `find([...])` を調整してください。

## ステップ 4. Vercel へデプロイ（フロント + 互換API）

1. [vercel.com](https://vercel.com/) → **Add New → Project** → `vie324/mavie_dash` をインポート。
2. Framework Preset は **Other**、Build Command は空（ビルド不要の静的サイト）。`api/` 配下は自動的に Serverless Function として認識されます。
3. **Settings → Environment Variables** に登録：
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_ROLE_KEY`
   - `SESSION_SECRET`（`openssl rand -hex 32` などで生成）
4. **Deploy**。発行URL（例 `https://mavie-dash.vercel.app`）を控える。
5. 動作確認：
   - `https://<your>.vercel.app/api?action=health` → `{"status":"ok",...}`
   - `https://<your>.vercel.app/api?action=get_data` → 売上の配列
   - `https://<your>.vercel.app/api?action=get_customers` → `{"status":"success","data":[...]}`

## ステップ 5. ダッシュボードの接続先を切り替え

`assets/js/dashboard.js` の接続先を Vercel の互換API に向けます。

```js
// assets/js/dashboard.js : 5 行目
// 変更前
const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
// 変更後
const DEFAULT_API_URL = 'https://<your>.vercel.app/api';
```

> 注意：ブラウザの `localStorage`（キー `mavie_spreadsheet_api_url`）に**旧URLが残っている**と、そちらが優先されます。
> 各端末で「設定」タブからAPI URLを新URLに保存し直すか、検証時はブラウザのアプリデータを一度クリアしてください。

変更を push すると Vercel が自動で再デプロイします。

## ステップ 6. 検証（並行運用）

旧環境（GitHub + GAS）はそのまま生かしつつ、Vercel のURLで以下を確認：

- [ ] サマリー / 売上詳細 / カレンダー / KPI・日報 / 顧客分析 / マーケ 各タブが表示
- [ ] 店舗・スタッフ専用URL（`?store=chiba&staff=kiki` 等）の絞り込み
- [ ] 売上目標設定・基本給・インセンティブの数値が旧環境と一致
- [ ] 「データ編集」タブで更新 → Supabase に反映（`update`）
- [ ] 管理ページ／スタッフページの認証（`verify_password` / `verify_session`）
- [ ] フォーム送信 → 当日中にダッシュボードへ反映（`onFormSubmit` 同期）
- [ ] 合計売上・件数が移行前後で一致

## ステップ 7. 本番切替（カットオーバー）

1. Vercel に**独自ドメイン**を割り当て（Settings → Domains、HTTPS 自動）。
2. 関係者の利用URLを Vercel のドメインへ変更。
3. 問題があればドメインを旧環境へ戻すだけでロールバック可能。
4. 安定後、旧 GitHub ホスティングは停止。スプレッドシート／GAS はバックアップとして当面保持（フォーム同期は継続）。

---

## 認証について（現状維持 → 後日強化）

本互換APIは、**現行のパスワード方式の挙動をそのまま再現**しています（管理者は `app_config.admin_password`、スタッフは `app_config.passwords_data` で照合、24時間有効のセッショントークン）。
- 管理者パスワードを設定するには、ダッシュボードの設定から保存（`save_settings` の `adminPassword`）。
- スタッフパスワードは従来どおり設定画面から保存（`save_passwords`）。

> セキュリティ強化（推奨・後日）：パスワードをURLクエリで送る現行方式から、**Supabase Auth（メール+パスワード, JWT）** へ移行すると、より安全になります。データ移行が安定してから着手して問題ありません。

## Gemini（AI アドバイス）について

現状はキーがブラウザに露出しています。データ移行が落ち着いたら、`api/ai-advice.js` のようなサーバー関数を追加し、`GEMINI_API_KEY` を Vercel 環境変数に置いてフロントから秘匿することを推奨します（本移行スコープ外・別途対応）。

---

## トラブルシューティング

| 症状 | 対処 |
|---|---|
| `/api?action=get_data` が500 | Vercel の環境変数（URL/キー）未設定。Logs を確認。 |
| ダッシュボードが旧データのまま | `localStorage` の旧API URL。設定タブで新URLを保存し直す。 |
| 認証が必ず失敗 | `SESSION_SECRET` 未設定、または `admin_password`/`passwords_data` 未登録。 |
| フォーム回答がSupabaseに来ない | トリガー未登録 / `service_role` キー未保存 / 項目名の不一致（`FormSync.gs` の `find` を調整）。 |
| 日付がずれる | タイムゾーン。売上は日付のみ、顧客は JST 表示で整形済み。元データの形式を確認。 |
