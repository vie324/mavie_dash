-- 初期データ（店舗マスタ + 既定スタッフ名簿）
-- 実データは scripts/migrate-from-gas.mjs でGASから移行するため、
-- このseedは「空のプロジェクトを最低限動かす」ためのもの。

insert into stores (id, name, sort_order) values
    ('chiba',     '千葉店',   1),
    ('honatsugi', '本厚木店', 2),
    ('yamato',    '大和店',   3)
on conflict (id) do nothing;

insert into staff (store_id, name, sort_order) values
    ('honatsugi', 'haruka', 1),
    ('honatsugi', 'vienna', 2),
    ('chiba',     'kiki',   1),
    ('chiba',     'karin',  2),
    ('chiba',     'nanami', 3),
    ('chiba',     'kanon',  4),
    ('chiba',     'ayami',  5),
    ('chiba',     'miki',   6),
    ('yamato',    'amano',  1)
on conflict (store_id, name) do nothing;
