# 本番移行手順書：GitHub + GAS → Vercel + Supabase

Mavie ダッシュボード（アイラッシュサロン）を、テスト環境（GitHub ホスティング + Google Apps Script + スプレッドシート）から、本番環境（Vercel + Supabase）へ移行するための手順をまとめたものです。

---

## 0. 現状（As-Is）と移行後（To-Be）

### 現状の構成

| レイヤー | 現状 |
|---|---|
| フロントエンド | `index.html` + `assets/js/dashboard.js`（約6,200行）+ `assets/css/dashboard.css` の**ビルド不要の静的サイト**。GitHub でホスト。 |
| バックエンドAPI | Google Apps Script（`gas/Code.gs`）の Web アプリ。`?action=...` 形式の REST もどき。 |
| データベース | Google スプレッドシート（シートを表として利用）。 |
| データ入力 | **Google フォーム** → スプレッドシート（売上日報・顧客カウンセリングの両方）。 |
| 認証 | 独自実装（パスワード照合 + セッショントークンを `PropertiesService` に保存）。 |
| AI 機能 | Gemini API を**ブラウザから直接**呼び出し（APIキーがフロントに露出）。 |
| 外部CDN | Tailwind / Chart.js / Lucide / Google Fonts（CDN読み込み）。 |

### 移行後の構成

| レイヤー | 移行後 |
|---|---|
| フロントエンド | **Vercel**（静的ホスティング + 自動デプロイ + 独自ドメイン + HTTPS）。 |
| バックエンドAPI | **Supabase**（PostgREST 自動API）+ 必要に応じて **Edge Functions / Vercel Functions**。 |
| データベース | **Supabase（PostgreSQL）**。 |
| データ入力 | 後述（Phase 3）。当面は Google フォーム継続 → 最終的にアプリ内フォーム。 |
| 認証 | **Supabase Auth**（推奨）or 現行方式をテーブル化して踏襲。 |
| AI 機能 | サーバー関数経由に変更し、**Gemini キーを秘匿**。 |

---

## 1. 移行前に決める3つの方針（重要）

実装に入る前に、以下3点の方針を決めてください。本手順書では **推奨案** を前提に進めます。

### 方針① データ入力（Google フォーム）をどうするか — ★最重要

現状、売上日報・顧客カウンセリングは **Google フォームで入力 → スプレッドシートに蓄積** されています。Supabase に移すと、この「入力経路」をどう置き換えるかが最大の論点です。

| 案 | 内容 | 工数 | リスク |
|---|---|---|---|
| **A（推奨・段階移行）** | Google フォームは当面そのまま。フォーム送信時に Apps Script トリガー（`onFormSubmit`）で **Supabase にも1行コピー**。入力体験を変えずにDBだけ移行。 | 小 | 低 |
| B（最終形） | Vercel 上に**自前の入力フォーム**（売上日報入力・顧客カウンセリング入力）を構築し、直接 Supabase へ。Google フォーム廃止。 | 大 | 中 |

→ **まず A で移行し、運用が安定してから B に進む** ルートを推奨します。

### 方針② バックエンドAPIの作り方

| 方式 | 内容 | フロント改修量 |
|---|---|---|
| **方式1（推奨・最小改修）** | **互換API**を Vercel Functions（or Supabase Edge Functions）で実装。現状GASと同じ `?action=...` の契約を保つ。フロントは接続先URLを差し替えるだけ。 | 極小 |
| 方式2（本格刷新） | フロントに `@supabase/supabase-js` を導入し、`fetch` を `supabase.from(...)` に全面置換。RLS + Supabase Auth で保護。 | 大 |

→ 本番化を急ぐなら **方式1（互換API）** で素早くカットオーバーし、その後 方式2 へ段階リファクタが現実的です。

### 方針③ 認証

- **推奨**：Supabase Auth（メール+パスワード）に置き換え。管理者/店舗/スタッフの権限は RLS とユーザーメタデータで制御。
- **簡易**：現行のパスワード方式をそのまま `app_config` テーブルに移して踏襲（短期的にはこれでも可）。

---

## 2. 全体フェーズ（チェックリスト）

- [ ] **Phase 0** 事前準備（アカウント・ツール）
- [ ] **Phase 1** Supabase データベース設計（テーブル作成）
- [ ] **Phase 2** 既存データの移行（スプレッドシート → Supabase）
- [ ] **Phase 3** データ入力経路の再設計（Google フォーム対応）
- [ ] **Phase 4** バックエンドAPI（互換API）の構築
- [ ] **Phase 5** フロントエンド改修（接続先の差し替え）
- [ ] **Phase 6** 認証の移行
- [ ] **Phase 7** Gemini APIキーの秘匿化
- [ ] **Phase 8** Vercel へのデプロイ
- [ ] **Phase 9** テスト・並行運用・本番切替（カットオーバー）
- [ ] **Phase 10** 運用・バックアップ・コスト管理

---

## Phase 0. 事前準備

1. **アカウント作成**
   - [Supabase](https://supabase.com/)（GitHub ログイン可）でプロジェクト作成。リージョンは `Northeast Asia (Tokyo)` を推奨。
   - [Vercel](https://vercel.com/)（GitHub ログイン可）。
2. **GitHub リポジトリ連携**：Vercel から `vie324/mavie_dash` をインポート。
3. **ローカル開発ツール**（任意・互換API実装に必要）
   - Node.js 18+ / npm
   - Supabase CLI（`npm i -g supabase`）
   - Vercel CLI（`npm i -g vercel`）
4. **控えておく値**
   - Supabase: `Project URL`、`anon key`、`service_role key`（**service_role は絶対にフロントに置かない**）。
   - 現行 GAS の URL（データ移行のエクスポート元として使用）：
     `https://script.google.com/macros/s/AKfycbwrz7LgQb2uH9VTtmalZxIEcJnHc-Ae53UNMnDi0MM5eLdP7XmZKOlDPTaOL5pmsFwf/exec`

---

## Phase 1. Supabase データベース設計

Supabase の **SQL Editor** で以下を実行してテーブルを作成します。現行のデータモデル（`gas/Code.gs` の `SALES_COLUMNS` と顧客シートのマッピング）に合わせています。

### 1-1. 売上日報テーブル

```sql
create table public.sales_reports (
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
create index on public.sales_reports (report_date);
create index on public.sales_reports (store, staff);
```

### 1-2. 顧客カウンセリングテーブル

```sql
create table public.customers (
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
create index on public.customers (store);
create index on public.customers (submitted_at);
```

### 1-3. 設定・目標・基本給（キーバリュー方式）

現状は「目標」「基本給」「設定」「パスワード」を JSON 文字列で保持しているため、まずは JSON のまま移すのが最短です。

```sql
create table public.app_config (
  key        text primary key,   -- 'goals_data' / 'salaries_data' / 'staff_roster' / 'admin_password' / 'gemini_api_key'
  value      jsonb,
  updated_at timestamptz not null default now()
);
```

> 補足：将来的には目標を `monthly_goals(month, store, staff, target_*)` のように正規化するとレポートが書きやすくなります。移行初期は JSON のままで構いません。

### 1-4. RLS（行レベルセキュリティ）

- **方式1（互換API）を選ぶ場合**：APIサーバー側で `service_role` キーを使うため、まずは RLS を有効化しつつ、サーバー経由でのみアクセスする運用にします（anon キーからは原則拒否）。
  ```sql
  alter table public.sales_reports enable row level security;
  alter table public.customers     enable row level security;
  alter table public.app_config    enable row level security;
  -- ポリシーを作らなければ anon からは読めない＝サーバー(service_role)経由のみ許可
  ```
- **方式2（supabase-js 直結）を選ぶ場合**：Supabase Auth と組み合わせ、`auth.uid()` やユーザーの店舗メタデータに基づく `select/insert/update` ポリシーを定義します（Phase 6 参照）。

---

## Phase 2. 既存データの移行（スプレッドシート → Supabase）

**ポイント**：現行 GAS の API がすでにデータを正規化（店舗名・スタッフ名・数値変換）してくれているので、**GAS のJSON出力をそのまま移行元として再利用**するのが最も確実です。

### 2-1. 移行スクリプト例（Node.js / 一度きり実行）

```js
// migrate.mjs   実行: node migrate.mjs
import { createClient } from '@supabase/supabase-js';

const GAS = 'https://script.google.com/macros/s/AKfycbwrz7LgQb2uH9VTtmalZxIEcJnHc-Ae53UNMnDi0MM5eLdP7XmZKOlDPTaOL5pmsFwf/exec';
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);

const toISO = (d) => {                    // '2026/1/5' → '2026-01-05'
  if (!d) return null;
  const [y, m, day] = d.split('/');
  return `${y}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
};

// ---- 売上日報 ----
const sales = await (await fetch(`${GAS}?action=get_data&nocache=true`)).json();
const salesRows = sales.map(r => ({
  report_date: toISO(r.date), store: r.store, staff: r.staff,
  sales_cash: r.sales.cash, sales_credit: r.sales.credit, sales_qr: r.sales.qr, sales_product: r.sales.product,
  discount_hpb_points: r.discounts.hpbPoints, discount_hpb_gift: r.discounts.hpbGift,
  discount_other: r.discounts.other, discount_refund: r.discounts.refund,
  cust_new_hpb: r.customers.newHPB, cust_new_mininai: r.customers.newMiniNai,
  cust_existing: r.customers.existing, cust_acquaintance: r.customers.acquaintance,
  nextres_new_hpb: r.nextRes.newHPB, nextres_new_mininai: r.nextRes.newMiniNai, nextres_existing: r.nextRes.existing,
  reviews_5star: r.reviews5Star, blog_updates: r.blogUpdates, sns_updates: r.snsUpdates,
}));
await supabase.from('sales_reports').insert(salesRows);

// ---- 顧客 ----
const cust = await (await fetch(`${GAS}?action=get_customers&nocache=true`)).json();
const custRows = (cust.data || []).map(c => ({
  store: c.store, submitted_at: c.timestamp ? new Date(c.timestamp).toISOString() : null,
  name: c.name, name_kana: c.nameKana, address: c.address, phone: c.phone,
  birthday: toISO(c.birthday), job: c.job, sns_ok: c.snsOk, visit_reason: c.visitReason,
  from_other_salon: c.fromOtherSalon, dissatisfaction: c.dissatisfaction, allergy: c.allergy,
  eyebrow_frequency: c.eyebrowFrequency, eyebrow_last_care: c.eyebrowLastCare, eyebrow_concern: c.eyebrowConcern,
  eyebrow_design: c.eyebrowDesign, eyebrow_design_image: c.eyebrowDesignImage,
  eyebrow_impression: c.eyebrowImpression, eyebrow_trouble: c.eyebrowTrouble,
  lash_frequency: c.lashFrequency, lash_design: c.lashDesign, lash_design_image: c.lashDesignImage,
  lash_eye_look: c.lashEyeLook, lash_contact: c.lashContact, lash_trouble: c.lashTrouble,
  agreement: c.agreement,
}));
await supabase.from('customers').insert(custRows);

// ---- 目標・基本給・設定 ----
const goals = await (await fetch(`${GAS}?action=load_goals`)).json();
const settings = await (await fetch(`${GAS}?action=load_settings`)).json();
await supabase.from('app_config').upsert([
  { key: 'goals_data',    value: goals.goals    || {} },
  { key: 'salaries_data', value: goals.salaries || {} },
  { key: 'staff_roster',  value: settings.settings?.staffRoster || null },
]);

console.log(`売上 ${salesRows.length} 件 / 顧客 ${custRows.length} 件を移行しました`);
```

### 2-2. 注意点

- **日付形式**：GAS は `yyyy/M/d`。PostgreSQL の `date` 型には `YYYY-MM-DD` で渡す（上記 `toISO` で変換）。
- **重複防止**：再実行時の二重登録を避けるため、移行は1回で完了させるか、`(report_date, store, staff)` に一意制約を付けて `upsert` にする。
- **件数照合**：移行後、Supabase の件数と GAS の件数（`get_data` の配列長・`get_customers` の `data.length`）が一致するか必ず確認。

---

## Phase 3. データ入力経路の再設計（Google フォーム対応）

### 案A（推奨）：Google フォームを残したまま Supabase に同期

フォームに紐づくスプレッドシートの Apps Script に、送信トリガー（`onFormSubmit`）を追加し、1件ずつ Supabase REST に POST します。

```js
// フォーム連携スプレッドシートの Apps Script に追加
function onFormSubmit(e) {
  const SUPABASE_URL = 'https://xxxx.supabase.co';
  const SERVICE_KEY  = '※スクリプトプロパティに保存（コードに直書きしない）';
  const v = e.namedValues; // フォーム項目名 → 値

  const payload = {            // ↓ 顧客フォームの例。項目名は実フォームに合わせる
    store: 'chiba',
    submitted_at: new Date().toISOString(),
    name: (v['お名前'] || [''])[0],
    phone: (v['電話番号'] || [''])[0],
    // … 既存の findCol マッピング（gas/Code.gs:751 付近）に対応させる
  };

  UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/customers`, {
    method: 'post',
    contentType: 'application/json',
    headers: { apikey: SERVICE_KEY, Authorization: 'Bearer ' + SERVICE_KEY, Prefer: 'return=minimal' },
    payload: JSON.stringify(payload),
  });
}
```

→ これで「入力は今まで通り Google フォーム、データは Supabase にも蓄積」となり、スタッフの操作を変えずに移行できます。

### 案B（最終形）：アプリ内フォーム

Vercel 上に売上日報入力ページ・顧客カウンセリング入力ページを新設し、Supabase へ直接 `insert`。Google フォーム/スプレッドシートを完全に廃止。落ち着いてから着手で構いません。

---

## Phase 4. バックエンドAPI（互換API）の構築 ＝ 方式1

フロント改修を最小化するため、現行 GAS と**同じ契約**のエンドポイントを Vercel Functions で用意します。`api/index.js` 1ファイルで `?action=` を振り分ける例：

```js
// api/index.js  （Vercel Serverless Function / Node ランタイム）
import { createClient } from '@supabase/supabase-js';
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_ROLE_KEY);

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*'); // 必要に応じてドメイン限定
  const action = (req.query.action) || (req.body && req.body.action) || 'get_data';

  if (req.method === 'GET') {
    switch (action) {
      case 'health':   return res.json({ status: 'ok' });
      case 'get_data': {
        const { data } = await supabase.from('sales_reports').select('*').order('report_date');
        return res.json(data.map(toSalesShape));      // ← 現行フロントが期待する形に整形
      }
      case 'get_customers': {
        const { data } = await supabase.from('customers').select('*');
        return res.json({ status: 'success', data: data.map(toCustomerShape) });
      }
      case 'load_goals': {
        const { data } = await supabase.from('app_config').select('*')
          .in('key', ['goals_data','salaries_data']);
        const m = Object.fromEntries(data.map(r => [r.key, r.value]));
        return res.json({ status: 'success', goals: m.goals_data || {}, salaries: m.salaries_data || {} });
      }
      // load_settings / verify_password / verify_session / get_all なども同様に
    }
  }

  if (req.method === 'POST') {
    switch (action) {
      case 'add_record': /* insert */ break;
      case 'update':     /* update */ break;
      case 'save_goals': /* app_config upsert */ break;
      // …
    }
  }
}
```

> `toSalesShape` / `toCustomerShape` は、Phase 2 の移行マッピングの**逆変換**（DB列 → 現行フロントが読む `r.sales.cash` などのネスト構造）です。これを用意すればフロントはほぼ無改修で動きます。
>
> キャッシュ層（GAS の `CacheService`）は PostgreSQL 直結なら基本不要。必要なら Vercel のエッジキャッシュや簡易メモリキャッシュで代替します。

### 現行 API ↔ Supabase 対応表（付録A も参照）

| 現行 GAS action | 役割 | 互換API実装の中身 |
|---|---|---|
| `get_data` / `get_all` | 売上取得 | `sales_reports` を select → 整形 |
| `get_customers` / `_today` / `_by_store` | 顧客取得 | `customers` を select（日付・店舗で絞り込み） |
| `load_goals` | 目標・基本給 | `app_config` の `goals_data` / `salaries_data` |
| `load_settings` | 設定 | `app_config` の `staff_roster` 等 |
| `load_passwords` / `verify_password` / `verify_session` | 認証 | Supabase Auth へ置換（Phase 6） |
| `update` | 売上更新 | `sales_reports` を update |
| `add_record` | 売上追加 | `sales_reports` に insert |
| `save_goals` / `save_settings` | 保存 | `app_config` を upsert |
| `clear_cache` / `health` | 補助 | 不要 or 簡易実装 |

---

## Phase 5. フロントエンド改修

方式1（互換API）なら改修は最小です。

1. **接続先URLの差し替え**：`assets/js/dashboard.js:5`
   ```js
   // 変更前
   const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
   // 変更後（Vercel の API へ）
   const DEFAULT_API_URL = 'https://<your-app>.vercel.app/api';
   ```
   `localStorage` に旧URLが残るため、設定画面から再保存するか、初回に強制上書きするコードを入れておくと安全です。
2. **CORS**：互換API側で `Access-Control-Allow-Origin` を返す（上記 Phase 4 参照）。
3. **POST の Content-Type**：GAS は `text/plain` を使う場面がありますが、Vercel Functions では `application/json` で受けられるよう実装を合わせます。
4. （任意）CDN 依存（Tailwind 等）は本番では**ビルド版に差し替え**ると安定・高速になります。最初は CDN のままでも動作します。

> 方式2（supabase-js 直結）を選ぶ場合は、`fetch(API_URL...)` を `supabase.from(...).select()/insert()/update()` に置換し、`@supabase/supabase-js` を読み込みます。改修量は大きいですが、最終形としては最もシンプルになります。

---

## Phase 6. 認証の移行

現状：`verify_password` でパスワード照合 →`PropertiesService` にセッショントークン保存、URLパラメータ `?store=&staff=` でスタッフ専用ビュー（`dashboard.js:2028` の `checkUrlParams`）。

### 推奨：Supabase Auth へ

1. 管理者・各スタッフを Supabase Auth のユーザーとして登録（メール+パスワード）。
2. ユーザーメタデータに `role`（admin/staff）、`store`、`staff` を持たせる。
3. フロントのログインを `supabase.auth.signInWithPassword()` に置換。セッションは Supabase が JWT で管理（独自トークン不要）。
4. スタッフ専用ビューの絞り込みは、URLパラメータではなく**ログインユーザーのメタデータ**で制御。RLS で他店舗データへのアクセスを遮断。

### 簡易：現行方式の踏襲

短期的には、`app_config` にパスワードを移し、互換APIで `verify_password` 相当を実装するだけでも動きます（セキュリティ強度は現状と同等）。本番ではできるだけ Supabase Auth への移行を推奨します。

---

## Phase 7. Gemini APIキーの秘匿化

現状（`dashboard.js:5566` 付近）は、Gemini キーをブラウザから直接 `generativelanguage.googleapis.com` に送っており、**キーが第三者に見える**状態です。

- キーを **Vercel の環境変数**（`GEMINI_API_KEY`）に移す。
- `api/ai-advice.js` のようなサーバー関数を作り、フロントはそこを呼ぶだけにする（キーはサーバー内に留める）。
- 設定画面の「Gemini APIキー入力欄」はサーバー設定に置き換え、フロントからは削除。

---

## Phase 8. Vercel へのデプロイ

1. **プロジェクト構成**（静的 + API 関数）
   ```
   /                      … 静的ファイル（index.html, assets/）をそのまま配信
   /api/index.js          … 互換API（Phase 4）
   /api/ai-advice.js      … Gemini プロキシ（Phase 7）
   vercel.json            … 配信・ルーティング設定
   ```
2. **`vercel.json` 例**
   ```json
   {
     "cleanUrls": true,
     "rewrites": [{ "source": "/api", "destination": "/api/index" }]
   }
   ```
3. **環境変数**（Vercel ダッシュボード → Settings → Environment Variables）
   - `SUPABASE_URL`、`SUPABASE_SERVICE_ROLE_KEY`、`GEMINI_API_KEY`
   - `service_role` は**サーバー関数からのみ**参照。フロント用に公開する場合は `anon key` のみ。
4. **デプロイ**：GitHub に push → Vercel が自動ビルド/デプロイ。`main` を本番、フィーチャーブランチをプレビューに割り当て。
5. **独自ドメイン**：Vercel にカスタムドメインを追加し、DNS を設定（HTTPS は自動発行）。

---

## Phase 9. テスト・並行運用・本番切替（カットオーバー）

1. **並行運用**：GAS/スプレッドシートを止めずに、Vercel 版を別URL（プレビュー）で動かして検証。
2. **動作確認チェックリスト**
   - [ ] サマリー・売上詳細・カレンダー・KPI・顧客分析・マーケの各タブが表示される
   - [ ] 店舗別/スタッフ別 専用URL（`?store=&staff=`）が正しく絞り込まれる
   - [ ] 目標設定・基本給・インセンティブ計算が一致する
   - [ ] データ編集（更新）が Supabase に反映される
   - [ ] 認証（管理者・スタッフ）が機能する
   - [ ] Gemini アドバイスが動く（キーは露出しない）
   - [ ] 件数・合計金額が移行前後で一致する
3. **データ同期の停止**：B案へ進む場合、カットオーバー時点で Google フォーム入力を停止し、アプリ内入力へ切替。A案継続なら同期を維持。
4. **本番切替**：独自ドメインを Vercel 版に向ける。問題があれば DNS を旧環境に戻すだけでロールバック可能。
5. **旧環境の停止**：安定後、GAS デプロイをアーカイブ（スプレッドシートはバックアップとして当面保持）。

---

## Phase 10. 運用・バックアップ・コスト

- **バックアップ**：Supabase の自動バックアップ（Pro 以上で日次）＋ 定期的に CSV エクスポート。
- **コスト目安**（2026年時点の一般的なプラン。最新は各公式で要確認）
  - Supabase：Free（500MB DB・1GBストレージ）でも開始可能だが、**Free は一定期間アクセスが無いとプロジェクトが一時停止**するため、本番は **Pro（約$25/月）** を推奨。
  - Vercel：Hobby は**非商用向け**。サロン業務（商用）では **Pro（約$20/月）** が適切。
  - Gemini API：従量課金（利用量次第）。
- **監視**：Vercel の Analytics / Logs、Supabase の Logs でエラー監視。
- **セキュリティ**：`service_role` キーは必ずサーバー側のみ。RLS を本番で必ず有効化。CORS は本番ドメインに限定。

---

## 付録A. 現行コードの参照ポイント

| 項目 | 場所 |
|---|---|
| 売上カラム定義 | `gas/Code.gs:70`（`SALES_COLUMNS`） |
| 売上の正規化・出力形 | `gas/Code.gs:414`（`getSalesData`） |
| 顧客シートの列マッピング | `gas/Code.gs:751`（`findCol`） |
| API ルーティング | `gas/Code.gs:259`（`doGet`）/ `gas/Code.gs:360`（`doPost`） |
| 認証 | `gas/Code.gs:1232`（`verifyPassword`）/ `1329`（`verifySession`） |
| フロントの接続先 | `assets/js/dashboard.js:5`（`DEFAULT_API_URL`） |
| スタッフ名簿 | `assets/js/dashboard.js:11`（`STAFF_ROSTER`） |
| スタッフ専用URL判定 | `assets/js/dashboard.js:2028`（`checkUrlParams`） |
| Gemini 呼び出し | `assets/js/dashboard.js:5566` 付近（`getAIAdvice`） |

## 付録B. 最短ルート（要点だけ）

本番化を最優先する場合の最小手順：

1. Supabase プロジェクト作成 → Phase 1 の SQL でテーブル作成。
2. Phase 2 のスクリプトで既存データを移行。
3. Phase 4 の互換APIを Vercel Functions で実装（`?action=` 契約を踏襲）。
4. `dashboard.js:5` の `DEFAULT_API_URL` を Vercel の `/api` に差し替え。
5. Phase 3 案A（`onFormSubmit` 同期）で入力経路を維持。
6. Vercel にデプロイ → 独自ドメイン割当 → 検証 → カットオーバー。
7. 落ち着いてから、認証の Supabase Auth 化（Phase 6）・Gemini 秘匿化（Phase 7）・アプリ内フォーム化（案B）を順次。
