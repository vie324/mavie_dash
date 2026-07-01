-- =====================================================================
-- vie Eyelash Salon Dashboard : Supabase 初期スキーマ + API (RPC)
--
-- 設計方針:
--   * フロントエンドは GAS と同じ「action API」感覚で使えるよう、
--     api_* という SECURITY DEFINER 関数群を PostgREST RPC として公開する。
--     レスポンス形状は GAS (gas/Code.gs) と互換。
--   * テーブルへの直接アクセスは RLS で遮断（daily_reports の SELECT のみ
--     Realtime 用に許可）。書き込みはすべて RPC 経由。
--   * パスワードは bcrypt(pgcrypto crypt) でハッシュ保存。平文は保持しない。
--
-- 適用方法: Supabase Dashboard > SQL Editor に貼り付けて実行
--           （または supabase db push / migration up）
-- =====================================================================

create extension if not exists pgcrypto with schema extensions;

-- ---------------------------------------------------------------------
-- 1. テーブル
-- ---------------------------------------------------------------------

create table if not exists stores (
    id          text primary key,              -- 'chiba' / 'honatsugi' / 'yamato'
    name        text not null,                 -- '千葉店' など表示名
    is_active   boolean not null default true,
    sort_order  integer not null default 0,
    created_at  timestamptz not null default now()
);

create table if not exists staff (
    id            bigint generated always as identity primary key,
    store_id      text not null references stores(id) on delete cascade,
    name          text not null,               -- 'kiki' などローマ字小文字（既存運用に合わせる）
    base_salary   integer not null default 220000,
    password_hash text,                        -- crypt() ハッシュ。null/'' = パスワード不要
    is_active     boolean not null default true,
    sort_order    integer not null default 0,
    created_at    timestamptz not null default now(),
    unique (store_id, name)
);

create table if not exists daily_reports (
    id                    bigint generated always as identity primary key,
    report_date           date not null,
    store_id              text not null references stores(id),
    staff_name            text not null,
    sales_cash            integer not null default 0,
    sales_credit          integer not null default 0,
    sales_qr              integer not null default 0,
    sales_product         integer not null default 0,
    discount_hpb_points   integer not null default 0,
    discount_hpb_gift     integer not null default 0,
    discount_other        integer not null default 0,
    refund                integer not null default 0,
    new_hpb               integer not null default 0,
    new_mininai           integer not null default 0,
    existing_count        integer not null default 0,
    acquaintance          integer not null default 0,
    next_res_new_hpb      integer not null default 0,
    next_res_new_mininai  integer not null default 0,
    next_res_existing     integer not null default 0,
    reviews_5star         integer not null default 0,
    blog_updates          integer not null default 0,
    sns_updates           integer not null default 0,
    created_at            timestamptz not null default now(),
    updated_at            timestamptz not null default now()
);
create index if not exists idx_daily_reports_date  on daily_reports (report_date);
create index if not exists idx_daily_reports_store on daily_reports (store_id, report_date);

-- 月別×店舗×スタッフの目標（構造変化に強いよう jsonb 保存）
create table if not exists goals (
    year_month  text not null,                 -- 'YYYY/M'（既存フォーマット踏襲）
    store_id    text not null,
    staff_name  text not null,
    data        jsonb not null default '{}',
    updated_at  timestamptz not null default now(),
    primary key (year_month, store_id, staff_name)
);

-- カウンセリング回答（Googleフォーム由来。可変項目は details jsonb に）
create table if not exists customers (
    id          bigint generated always as identity primary key,
    store_id    text,
    answered_at timestamptz,
    name        text,
    name_kana   text,
    details     jsonb not null default '{}',
    created_at  timestamptz not null default now()
);
create index if not exists idx_customers_store_date on customers (store_id, answered_at desc);

-- アプリ設定 (ad_costs / monthly_close / admin_password_hash など)
create table if not exists app_settings (
    key        text primary key,
    value      jsonb,
    updated_at timestamptz not null default now()
);

-- updated_at 自動更新
create or replace function set_updated_at() returns trigger
language plpgsql as $$
begin
    new.updated_at = now();
    return new;
end $$;

drop trigger if exists trg_daily_reports_updated on daily_reports;
create trigger trg_daily_reports_updated before update on daily_reports
for each row execute function set_updated_at();

drop trigger if exists trg_goals_updated on goals;
create trigger trg_goals_updated before update on goals
for each row execute function set_updated_at();

-- ---------------------------------------------------------------------
-- 2. RLS（直接アクセス遮断。書き込みは RPC 経由のみ）
-- ---------------------------------------------------------------------
alter table stores        enable row level security;
alter table staff         enable row level security;
alter table daily_reports enable row level security;
alter table goals         enable row level security;
alter table customers     enable row level security;
alter table app_settings  enable row level security;

-- Realtime(postgres_changes) 用に daily_reports の SELECT のみ匿名許可
drop policy if exists "daily_reports_read" on daily_reports;
create policy "daily_reports_read" on daily_reports for select using (true);

-- Realtime パブリケーションへ追加（ローカル環境等で無い場合は無視）
do $$ begin
    alter publication supabase_realtime add table daily_reports;
exception when others then null;
end $$;

-- ---------------------------------------------------------------------
-- 3. ヘルパー
-- ---------------------------------------------------------------------

-- 'YYYY/M/D'（ゼロ埋めなし・既存GAS互換）の日付文字列
create or replace function jp_date_str(d date) returns text
language sql immutable as $$
    select extract(year from d)::int || '/' || extract(month from d)::int || '/' || extract(day from d)::int
$$;

-- ---------------------------------------------------------------------
-- 4. API (RPC) — GAS互換レスポンス
-- ---------------------------------------------------------------------

-- ?action=health 相当
create or replace function api_health() returns jsonb
language sql security definer set search_path = public, extensions as $$
    select jsonb_build_object(
        'status', 'success',
        'version', '0001',
        'timestamp', now(),
        'reports', (select count(*) from daily_reports),
        'customers', (select count(*) from customers)
    )
$$;

-- ?action=get_data 相当（素の配列を返す）
create or replace function api_get_sales() returns jsonb
language sql security definer set search_path = public, extensions as $$
    select coalesce(jsonb_agg(jsonb_build_object(
        'id', r.id,
        'date', jp_date_str(r.report_date),
        'store', r.store_id,
        'staff', r.staff_name,
        'sales', jsonb_build_object(
            'cash', r.sales_cash, 'credit', r.sales_credit,
            'qr', r.sales_qr, 'product', r.sales_product),
        'discounts', jsonb_build_object(
            'hpbPoints', r.discount_hpb_points, 'hpbGift', r.discount_hpb_gift,
            'other', r.discount_other, 'refund', r.refund),
        'customers', jsonb_build_object(
            'newHPB', r.new_hpb, 'newMiniNai', r.new_mininai,
            'existing', r.existing_count, 'acquaintance', r.acquaintance),
        'nextRes', jsonb_build_object(
            'newHPB', r.next_res_new_hpb, 'newMiniNai', r.next_res_new_mininai,
            'existing', r.next_res_existing),
        'reviews5Star', r.reviews_5star,
        'blogUpdates', r.blog_updates,
        'snsUpdates', r.sns_updates
    ) order by r.report_date, r.id), '[]'::jsonb)
    from daily_reports r
$$;

-- ?action=add_record 相当
create or replace function api_add_record(p jsonb) returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
begin
    -- 店舗が未登録なら自動作成（GASは自由入力のため互換動作）
    insert into stores (id, name) values (p->>'store', p->>'store')
    on conflict (id) do nothing;

    insert into daily_reports (
        report_date, store_id, staff_name,
        sales_cash, sales_credit, sales_qr, sales_product,
        discount_hpb_points, discount_hpb_gift, discount_other, refund,
        new_hpb, new_mininai, existing_count, acquaintance,
        next_res_new_hpb, next_res_new_mininai, next_res_existing,
        reviews_5star, blog_updates, sns_updates
    ) values (
        to_date(p->>'date', 'YYYY/MM/DD'),
        p->>'store',
        p->>'staff',
        coalesce((p#>>'{sales,cash}')::int, 0),
        coalesce((p#>>'{sales,credit}')::int, 0),
        coalesce((p#>>'{sales,qr}')::int, 0),
        coalesce((p#>>'{sales,product}')::int, 0),
        coalesce((p#>>'{discounts,hpbPoints}')::int, 0),
        coalesce((p#>>'{discounts,hpbGift}')::int, 0),
        coalesce((p#>>'{discounts,other}')::int, 0),
        coalesce((p#>>'{discounts,refund}')::int, 0),
        coalesce((p#>>'{customers,newHPB}')::int, 0),
        coalesce((p#>>'{customers,newMiniNai}')::int, 0),
        coalesce((p#>>'{customers,existing}')::int, 0),
        coalesce((p#>>'{customers,acquaintance}')::int, 0),
        coalesce((p#>>'{nextRes,newHPB}')::int, 0),
        coalesce((p#>>'{nextRes,newMiniNai}')::int, 0),
        coalesce((p#>>'{nextRes,existing}')::int, 0),
        coalesce((p->>'reviews5Star')::int, 0),
        coalesce((p->>'blogUpdates')::int, 0),
        coalesce((p->>'snsUpdates')::int, 0)
    );
    return jsonb_build_object('status', 'success');
end $$;

-- ?action=update 相当（データ編集タブの一括保存）
create or replace function api_update_rows(p jsonb) returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
declare
    r jsonb;
    n integer := 0;
begin
    for r in select * from jsonb_array_elements(coalesce(p, '[]'::jsonb)) loop
        update daily_reports set
            sales_cash           = coalesce((r#>>'{sales,cash}')::int, sales_cash),
            sales_credit         = coalesce((r#>>'{sales,credit}')::int, sales_credit),
            sales_qr             = coalesce((r#>>'{sales,qr}')::int, sales_qr),
            sales_product        = coalesce((r#>>'{sales,product}')::int, sales_product),
            discount_hpb_points  = coalesce((r#>>'{discounts,hpbPoints}')::int, discount_hpb_points),
            discount_hpb_gift    = coalesce((r#>>'{discounts,hpbGift}')::int, discount_hpb_gift),
            discount_other       = coalesce((r#>>'{discounts,other}')::int, discount_other),
            refund               = coalesce((r#>>'{discounts,refund}')::int, refund),
            new_hpb              = coalesce((r#>>'{customers,newHPB}')::int, new_hpb),
            new_mininai          = coalesce((r#>>'{customers,newMiniNai}')::int, new_mininai),
            existing_count       = coalesce((r#>>'{customers,existing}')::int, existing_count),
            acquaintance         = coalesce((r#>>'{customers,acquaintance}')::int, acquaintance),
            next_res_new_hpb     = coalesce((r#>>'{nextRes,newHPB}')::int, next_res_new_hpb),
            next_res_new_mininai = coalesce((r#>>'{nextRes,newMiniNai}')::int, next_res_new_mininai),
            next_res_existing    = coalesce((r#>>'{nextRes,existing}')::int, next_res_existing),
            reviews_5star        = coalesce((r->>'reviews5Star')::int, reviews_5star),
            blog_updates         = coalesce((r->>'blogUpdates')::int, blog_updates),
            sns_updates          = coalesce((r->>'snsUpdates')::int, sns_updates)
        where id = (r->>'id')::bigint;
        if found then n := n + 1; end if;
    end loop;
    return jsonb_build_object('status', 'success', 'updated', n);
end $$;

-- ?action=get_customers / get_customers_today / get_customers_by_store 相当
create or replace function api_get_customers(p_store text default null, p_today boolean default false)
returns jsonb
language sql security definer set search_path = public, extensions as $$
    select coalesce(jsonb_agg(t.obj order by t.answered_at desc nulls last), '[]'::jsonb)
    from (
        select c.answered_at,
               c.details || jsonb_build_object(
                   'id', c.id,
                   'store', c.store_id,
                   'timestamp', (extract(epoch from c.answered_at) * 1000)::bigint,
                   'name', c.name,
                   'nameKana', c.name_kana
               ) as obj
        from customers c
        where (p_store is null or c.store_id = p_store)
          and (not p_today
               or (c.answered_at at time zone 'Asia/Tokyo')::date
                  = (now() at time zone 'Asia/Tokyo')::date)
    ) t
$$;

-- ?action=load_goals 相当
create or replace function api_load_goals() returns jsonb
language sql security definer set search_path = public, extensions as $$
    select jsonb_build_object(
        'goals', coalesce((
            select jsonb_object_agg(ym, stores_obj) from (
                select year_month as ym, jsonb_object_agg(store_id, staff_obj) as stores_obj
                from (
                    select year_month, store_id, jsonb_object_agg(staff_name, data) as staff_obj
                    from goals group by year_month, store_id
                ) a group by year_month
            ) b
        ), '{}'::jsonb),
        'salaries', coalesce((
            select jsonb_object_agg(store_id, staff_obj) from (
                select store_id, jsonb_object_agg(name, base_salary) as staff_obj
                from staff where is_active group by store_id
            ) s
        ), '{}'::jsonb)
    )
$$;

-- ?action=save_goals 相当
create or replace function api_save_goals(p_goals jsonb default '{}', p_salaries jsonb default '{}')
returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
declare
    ym record; st record; sf record;
begin
    for ym in select * from jsonb_each(coalesce(p_goals, '{}'::jsonb)) loop
        for st in select * from jsonb_each(ym.value) loop
            insert into stores (id, name) values (st.key, st.key) on conflict (id) do nothing;
            for sf in select * from jsonb_each(st.value) loop
                insert into goals (year_month, store_id, staff_name, data)
                values (ym.key, st.key, sf.key, sf.value)
                on conflict (year_month, store_id, staff_name)
                do update set data = excluded.data, updated_at = now();
            end loop;
        end loop;
    end loop;

    for st in select * from jsonb_each(coalesce(p_salaries, '{}'::jsonb)) loop
        insert into stores (id, name) values (st.key, st.key) on conflict (id) do nothing;
        for sf in select * from jsonb_each_text(st.value) loop
            insert into staff (store_id, name, base_salary)
            values (st.key, sf.key, coalesce(sf.value::int, 220000))
            on conflict (store_id, name)
            do update set base_salary = excluded.base_salary;
        end loop;
    end loop;
    return jsonb_build_object('status', 'success');
end $$;

-- ?action=load_settings 相当
-- 注: anthropicApiKey はサーバーに保存しない（Edge Function の Secret を使用）
create or replace function api_load_settings() returns jsonb
language sql security definer set search_path = public, extensions as $$
    select jsonb_build_object(
        'staffRoster', (
            select jsonb_object_agg(s.id, names order by s.sort_order, s.id) from (
                select st.id, st.sort_order,
                       (select coalesce(jsonb_agg(f.name order by f.sort_order, f.created_at), '[]'::jsonb)
                        from staff f where f.store_id = st.id and f.is_active) as names
                from stores st where st.is_active
            ) s
        ),
        'anthropicApiKey', null,
        'adCosts', (select value from app_settings where key = 'ad_costs'),
        'monthlyClose', (select value from app_settings where key = 'monthly_close')
    )
$$;

-- ?action=save_settings 相当
create or replace function api_save_settings(p jsonb) returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
declare
    st record; nm text;
begin
    -- スタッフ名簿の同期（追加=upsert / 削除=非アクティブ化。履歴は消さない）
    if p ? 'staffRoster' and jsonb_typeof(p->'staffRoster') = 'object' then
        update stores set is_active = (p->'staffRoster' ? id);
        for st in select * from jsonb_each(p->'staffRoster') loop
            insert into stores (id, name, is_active) values (st.key, st.key, true)
            on conflict (id) do update set is_active = true;
            update staff set is_active = false
            where store_id = st.key
              and not (name in (select jsonb_array_elements_text(st.value)));
            for nm in select jsonb_array_elements_text(st.value) loop
                insert into staff (store_id, name) values (st.key, nm)
                on conflict (store_id, name) do update set is_active = true;
            end loop;
        end loop;
    end if;

    -- 管理パスワード（ハッシュ化して保存。空文字 = 無効化）
    if p ? 'adminPassword' then
        insert into app_settings (key, value)
        values ('admin_password_hash',
                case when coalesce(p->>'adminPassword', '') = ''
                     then 'null'::jsonb
                     else to_jsonb(crypt(p->>'adminPassword', gen_salt('bf'))) end)
        on conflict (key) do update set value = excluded.value, updated_at = now();
    end if;

    -- 広告費 / 月締め確定
    if p ? 'adCosts' then
        insert into app_settings (key, value) values ('ad_costs', p->'adCosts')
        on conflict (key) do update set value = excluded.value, updated_at = now();
    end if;
    if p ? 'monthlyClose' then
        insert into app_settings (key, value) values ('monthly_close', p->'monthlyClose')
        on conflict (key) do update set value = excluded.value, updated_at = now();
    end if;

    -- anthropicApiKey は意図的に保存しない（Edge Function Secret 推奨）
    return jsonb_build_object('status', 'success');
end $$;

-- ?action=load_passwords 相当
-- 平文は返さない。設定済みなら '__SET__'、未設定なら '' を返す。
create or replace function api_load_passwords() returns jsonb
language sql security definer set search_path = public, extensions as $$
    select coalesce((
        select jsonb_object_agg(store_id, staff_obj) from (
            select store_id,
                   jsonb_object_agg(name,
                       case when coalesce(password_hash, '') = '' then '' else '__SET__' end) as staff_obj
            from staff where is_active group by store_id
        ) t
    ), '{}'::jsonb)
$$;

-- ?action=save_passwords 相当
-- 入力はGAS互換のフラット形式 {"store_staff": "newpass" | "" | "__SET__"}。
-- '__SET__'(変更なしプレースホルダ)はスキップ。'' はパスワード無効化。
create or replace function api_save_passwords(p jsonb) returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
declare
    e record; v_store text; v_staff text;
begin
    for e in select * from jsonb_each_text(coalesce(p, '{}'::jsonb)) loop
        continue when e.value = '__SET__';
        v_store := split_part(e.key, '_', 1);
        v_staff := substr(e.key, length(v_store) + 2);
        continue when v_store = '' or v_staff = '';
        insert into stores (id, name) values (v_store, v_store) on conflict (id) do nothing;
        insert into staff (store_id, name, password_hash)
        values (v_store, v_staff,
                case when e.value = '' then null else crypt(e.value, gen_salt('bf')) end)
        on conflict (store_id, name) do update
        set password_hash = excluded.password_hash, is_active = true;
    end loop;
    return jsonb_build_object('status', 'success');
end $$;

-- ?action=verify_password 相当（GAS同様: パスワード未設定なら誰でも通過）
create or replace function api_verify_password(
    p_page_type text, p_store text default '', p_staff text default '', p_password text default ''
) returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
declare
    stored text;
    ok boolean := false;
begin
    if p_page_type = 'admin' then
        stored := (select value #>> '{}' from app_settings where key = 'admin_password_hash');
    else
        select password_hash into stored
        from staff
        where store_id = p_store and lower(name) = lower(p_staff) and is_active
        limit 1;
    end if;
    ok := stored is null or stored = '' or stored = crypt(coalesce(p_password, ''), stored);
    return jsonb_build_object('ok', ok);
end $$;

-- ---------------------------------------------------------------------
-- 5. 権限（匿名キーで RPC のみ実行可能に）
-- ---------------------------------------------------------------------
grant usage on schema public to anon, authenticated;
grant execute on function
    api_health(),
    api_get_sales(),
    api_add_record(jsonb),
    api_update_rows(jsonb),
    api_get_customers(text, boolean),
    api_load_goals(),
    api_save_goals(jsonb, jsonb),
    api_load_settings(),
    api_save_settings(jsonb),
    api_load_passwords(),
    api_save_passwords(jsonb),
    api_verify_password(text, text, text, text),
    jp_date_str(date)
to anon, authenticated;
