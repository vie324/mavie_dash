-- =====================================================================
-- vie Dashboard : 月次面談・振り返り (Reviews)
-- 0001_init.sql の後に適用。GAS action と互換の RPC を公開する。
-- （一括セットアップは supabase/setup.sql 参照）
-- =====================================================================

create table if not exists reviews (
    id               bigint generated always as identity primary key,
    year_month       text not null,              -- 'YYYY/M'
    store_id         text not null,
    staff_name       text not null,
    meeting_date     date,
    interviewer      text,
    service          jsonb not null default '{}', -- {rating, reflection, action}
    retail           jsonb not null default '{}',
    total            jsonb not null default '{}',
    metrics          jsonb not null default '{}', -- {service,retail,total:{target,actual,rate,gap}}
    ai_evaluation    jsonb,                       -- {text, ratedAt}
    ai_evaluated_at  timestamptz,
    status           text not null default 'submitted',
    submitted_at     timestamptz default now(),
    created_at       timestamptz not null default now(),
    updated_at       timestamptz not null default now(),
    unique (year_month, store_id, staff_name)
);
create index if not exists idx_reviews_ym on reviews (year_month);

alter table reviews enable row level security;

drop trigger if exists trg_reviews_updated on reviews;
create trigger trg_reviews_updated before update on reviews
for each row execute function set_updated_at();

-- ?action=get_reviews 相当（ネスト構造 {ym:{store:{staff:{...}}}}）
create or replace function api_get_reviews() returns jsonb
language sql security definer set search_path = public, extensions as $$
    select coalesce((
        select jsonb_object_agg(ym, stores_obj) from (
            select year_month as ym, jsonb_object_agg(store_id, staff_obj) as stores_obj from (
                select year_month, store_id, jsonb_object_agg(staff_name, rec) as staff_obj from (
                    select year_month, store_id, staff_name, jsonb_build_object(
                        'meetingDate', to_char(meeting_date, 'YYYY-MM-DD'),
                        'interviewer', interviewer,
                        'service', service, 'retail', retail, 'total', total,
                        'metrics', metrics,
                        'status', status,
                        'submittedAt', case when submitted_at is null then null else (extract(epoch from submitted_at) * 1000)::bigint end,
                        'updatedAt', (extract(epoch from updated_at) * 1000)::bigint,
                        'ai', ai_evaluation
                    ) as rec
                    from reviews
                ) r group by year_month, store_id
            ) a group by year_month
        ) b
    ), '{}'::jsonb)
$$;

-- ?action=save_review 相当（スタッフ入力の upsert。AI評価は保持）
create or replace function api_save_review(p jsonb) returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
begin
    insert into stores (id, name) values (p->>'store', p->>'store') on conflict (id) do nothing;
    insert into reviews (year_month, store_id, staff_name, meeting_date, interviewer,
                         service, retail, total, metrics, status, submitted_at, updated_at)
    values (
        p->>'yearMonth', p->>'store', p->>'staff',
        nullif(p->>'meetingDate', '')::date,
        p->>'interviewer',
        coalesce(p->'service', '{}'::jsonb),
        coalesce(p->'retail', '{}'::jsonb),
        coalesce(p->'total', '{}'::jsonb),
        coalesce(p->'metrics', '{}'::jsonb),
        coalesce(p->>'status', 'submitted'),
        now(), now()
    )
    on conflict (year_month, store_id, staff_name) do update set
        meeting_date = nullif(p->>'meetingDate', '')::date,
        interviewer  = p->>'interviewer',
        service      = coalesce(p->'service', '{}'::jsonb),
        retail       = coalesce(p->'retail', '{}'::jsonb),
        total        = coalesce(p->'total', '{}'::jsonb),
        metrics      = coalesce(p->'metrics', reviews.metrics),
        status       = coalesce(p->>'status', 'submitted'),
        submitted_at = coalesce(reviews.submitted_at, now()),
        updated_at   = now();
    return jsonb_build_object('status', 'success');
end $$;

-- ?action=save_review_ai 相当（管理者のAI評価保存。未提出でも保存可）
create or replace function api_save_review_ai(
    p_year_month text, p_store text, p_staff text, p_ai jsonb, p_metrics jsonb default null
) returns jsonb
language plpgsql security definer set search_path = public, extensions as $$
begin
    insert into stores (id, name) values (p_store, p_store) on conflict (id) do nothing;
    insert into reviews (year_month, store_id, staff_name, ai_evaluation, ai_evaluated_at, metrics, status, submitted_at)
    values (p_year_month, p_store, p_staff, p_ai, now(), coalesce(p_metrics, '{}'::jsonb), 'reviewed', null)
    on conflict (year_month, store_id, staff_name) do update set
        ai_evaluation   = p_ai,
        ai_evaluated_at = now(),
        metrics         = coalesce(p_metrics, reviews.metrics),
        updated_at      = now();
    return jsonb_build_object('status', 'success');
end $$;

grant execute on function
    api_get_reviews(),
    api_save_review(jsonb),
    api_save_review_ai(text, text, text, jsonb, jsonb)
to anon, authenticated;
