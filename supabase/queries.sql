-- vie ダッシュボード: 確認・集計用のSQL（Supabase の SQL Editor に1つずつ貼って実行）
--
-- ※ アカウント（パスワード）の発行は SQL ではなく、ダッシュボードの
--   「設定 → スタッフアカウントの発行」（一括発行ボタン）で行ってください。
--   パスワードは scrypt でハッシュ化して保存するため SQL では正しく作れず、
--   スタッフID（SalonOneのID）も画面なら自動で選べます。
--   退職したスタッフは SalonOne でスタッフを削除すると、そのURLは開けなくなります。
--
-- ここにあるのはすべて「読み取り専用」（データは変更しません）。

-- ① 発行済みアカウントの一覧（パスワードは表示しません）
--    account = staff:<SalonOneのスタッフID> / store:<SalonOneの店舗ID>（店長用）
select e.key as account,
       (e.value->>'updatedAt')::timestamptz at time zone 'Asia/Tokyo' as updated_at_jst
from public.vie_kv kv
cross join lateral jsonb_each(kv.value) as e(key, value)
where kv.key = 'vie:accounts'
order by e.key;

-- ② スタッフ別の次回予約（月ごと・日報の合計）
--    next_new = 新規の次回予約、next_repeat = 2回目以降の次回予約、report_days = 日報を入れた日数
select substr(kv.key, 12) as month,
       split_part(d.key, ':', 2) as staff_id,
       count(*) as report_days,
       sum(coalesce((d.value->>'nextNew')::int, 0)) as next_new,
       sum(coalesce((d.value->>'nextRepeat')::int, 0)) as next_repeat
from public.vie_kv kv
cross join lateral jsonb_each(kv.value->'daily') as d(key, value)
where kv.key like 'vie:manual:%'
group by 1, 2
order by 1 desc, 4 desc, 5 desc;

-- ③ 媒体別の次回予約（月ごと）
--    source = SalonOneの流入元ID（other = その他・不明）。媒体名はダッシュボードのマーケタブで確認できます
--    ※ 媒体別の入力を始める前の日報は内訳がないため、ここには含まれません
select substr(kv.key, 12) as month,
       s.key as source,
       sum(coalesce((s.value->>'n')::int, 0)) as next_new,
       sum(coalesce((s.value->>'r')::int, 0)) as next_repeat
from public.vie_kv kv
cross join lateral jsonb_each(kv.value->'daily') as d(key, value)
cross join lateral jsonb_each(coalesce(d.value->'src', '{}'::jsonb)) as s(key, value)
where kv.key like 'vie:manual:%'
group by 1, 2
order by 1 desc, sum(coalesce((s.value->>'n')::int, 0)) + sum(coalesce((s.value->>'r')::int, 0)) desc;
