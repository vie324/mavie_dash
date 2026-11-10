// 給与・インセンティブ計算のテスト（実行: npm test）
// 仕様書のテストケースをそのまま確認する。金額は config.js の 2026年11月〜の設定を使う。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';

import { PAYROLL_RULES, NEW_RULE_START_MONTH, ruleForMonth } from '../assets/js/payroll/config.js';
import {
    calcIncentive, nextTier, nextTierLabel, calcOvertimeTopUp, calcRetailAllowance,
    calcMonthlyPay, settingForMonth, salesForPayroll, ceilYen, hasHours, effectiveTierRates, validCustomRates,
} from '../assets/js/payroll/calc.js';

const require = createRequire(import.meta.url);
const payrollApi = require('../api/_lib/payroll.js')._internal;

const RULE = ruleForMonth('2026-11');
const yes = true, no = false;

// ---- インセンティブ（仕様書のテストケース）----
// [パターン, 税込施術売上, 開店・締め, 率, インセンティブ, 総支給（残業精算前）]
// ※ full 980,000 なし の総支給は仕様書の 385,560 が誤記で、正しくは 385,860（依頼者確認済み）
const INCENTIVE_CASES = [
    ['full', 600000, yes, 0, 0, 240000],
    ['full', 980000, yes, 40, 149600, 389600],
    ['full', 980000, no, 39, 145860, 385860],
    ['full', 1000000, yes, 40, 157600, 397600],
    ['full', 1100000, yes, 42, 207480, 447480],
    ['short_atsugi', 540000, no, 36, 2880, 220580],
    ['short_atsugi', 760000, no, 37, 84360, 302060],
    ['short_atsugi', 850000, no, 39, 124020, 341720],
    ['short_atsugi', 850000, yes, 40, 127200, 344900],
    ['short_chiba', 514001, no, 36, 361, 210461],
    ['short_chiba', 850000, no, 39, 131430, 341530],
    ['short_chiba', 900000, no, 40, 154800, 364900],
];

for (const [pattern, sales, duty, rate, incentive, total] of INCENTIVE_CASES) {
    test(`インセンティブ: ${pattern} 売上${sales} 開店締め${duty ? 'あり' : 'なし'} → ${rate}% / ${incentive}円 / 総支給${total}円`, () => {
        const r = calcIncentive(RULE, pattern, sales, duty);
        assert.equal(r.rate, rate);
        assert.equal(r.incentive, incentive);
        const pay = calcMonthlyPay(RULE, { patternKey: pattern, storeDuty: duty, serviceSales: sales });
        assert.equal(pay.rate, rate);
        assert.equal(pay.incentive, incentive);
        assert.equal(pay.totalBeforeOvertime, total);
        assert.equal(pay.total, total, '勤怠未入力・物販0円なら総支給 = 残業精算前');
    });
}

// ---- 次の段階 ----
const NEXT_CASES = [
    ['full', 980000, yes, '1,000,000超で41%（あと20,000円）'],
    ['short_atsugi', 540000, no, '559,000超で37%（あと19,000円）'],
    ['full', 1100000, yes, '次の段階なし'],
];
for (const [pattern, sales, duty, label] of NEXT_CASES) {
    test(`次の段階: ${pattern} 売上${sales} → ${label}`, () => {
        assert.equal(nextTierLabel(nextTier(RULE, pattern, sales, duty)), label);
        assert.equal(calcMonthlyPay(RULE, { patternKey: pattern, storeDuty: duty, serviceSales: sales }).nextLabel, label);
    });
}

test('次の段階: 売上がちょうど段階の金額のときは前の段階の率のまま・「あと1円」', () => {
    assert.equal(calcIncentive(RULE, 'full', 1000000, yes).rate, 40);
    assert.deepEqual(nextTier(RULE, 'full', 1000000, yes), { threshold: 1000000, rate: 41, remaining: 1 });
});

test('次の段階: 開始ライン以下なら開始ラインが次の段階（開店・締めの調整込みの率）', () => {
    assert.equal(nextTierLabel(nextTier(RULE, 'full', 600000, yes)), '606,000超で38%（あと6,000円）');
    assert.equal(nextTierLabel(nextTier(RULE, 'full', 600000, no)), '606,000超で37%（あと6,000円）');
    assert.equal(nextTierLabel(nextTier(RULE, 'short_chiba', 0, yes)), '513,000超で37%（あと513,000円）');
});

test('インセンティブ: 開始ライン以下は開店・締めの調整があっても0%・0円', () => {
    assert.deepEqual(
        { rate: calcIncentive(RULE, 'short_atsugi', 532000, yes).rate, inc: calcIncentive(RULE, 'short_atsugi', 532000, yes).incentive },
        { rate: 0, inc: 0 },
    );
});

test('インセンティブ: 開店・締めの指定がなければパターンの標準（フル=あり、時短=なし）', () => {
    assert.equal(calcIncentive(RULE, 'full', 980000).rate, 40);
    assert.equal(calcIncentive(RULE, 'short_atsugi', 850000).rate, 39);
    assert.equal(calcIncentive(RULE, 'full', 980000, null).rate, 40);
});

// ---- 残業精算 ----
// [パターン, インセンティブ, 残業, 法定時間外, 総労働時間, 追加支給]
const OVERTIME_CASES = [
    ['full', 149600, 6, 6, 185.33, 1143],
    ['full', 149600, 5, 5, 184.33, 0],
    ['full', 0, 8, 8, 187.33, 3143],
    ['short_atsugi', 124020, 4, 0, 160.92, 1675],
];
for (const [pattern, incentive, ot, stat, total, topUp] of OVERTIME_CASES) {
    test(`残業精算: ${pattern} インセンティブ${incentive} 残業${ot}h 法定外${stat}h 総労働${total}h → ${topUp}円`, () => {
        assert.equal(calcOvertimeTopUp(RULE, pattern, incentive, ot, stat, total), topUp);
    });
}

test('残業精算: 総労働時間が0ならインセンティブ部分は0（0で割らない）', () => {
    assert.equal(calcOvertimeTopUp(RULE, 'full', 149600, 8, 8, 0), 3143);
});

test('総支給: 勤怠を入力すると残業の追加支給が足される', () => {
    const pay = calcMonthlyPay(RULE, {
        patternKey: 'full', storeDuty: yes, serviceSales: 980000,
        hours: { overtime: 6, statutoryOvertime: 6, totalHours: 185.33 },
    });
    assert.equal(pay.hoursEntered, true);
    assert.equal(pay.overtimeTopUp, 1143);
    assert.equal(pay.total, 389600 + 1143);
});

test('総支給: 勤怠が未入力なら残業の追加支給は0円（未入力扱い）', () => {
    const pay = calcMonthlyPay(RULE, { patternKey: 'full', serviceSales: 980000, hours: {} });
    assert.equal(pay.hoursEntered, false);
    assert.equal(pay.overtimeTopUp, 0);
    assert.equal(hasHours({ overtime: null, totalHours: '' }), false);
    assert.equal(hasHours({ totalHours: 0 }), true);
});

// ---- 物販手当（依頼者の追加仕様: 5%、10万円を超えたら全額に10%、税込のまま、1円未満切り上げ）----
test('物販手当: 0円・5万円・ちょうど10万円は5%、10万円を超えたら全額に10%', () => {
    assert.deepEqual([calcRetailAllowance(RULE, 0).rate, calcRetailAllowance(RULE, 0).allowance], [0, 0]);
    assert.deepEqual([calcRetailAllowance(RULE, 50000).rate, calcRetailAllowance(RULE, 50000).allowance], [5, 2500]);
    assert.deepEqual([calcRetailAllowance(RULE, 100000).rate, calcRetailAllowance(RULE, 100000).allowance], [5, 5000]);
    assert.deepEqual([calcRetailAllowance(RULE, 100001).rate, calcRetailAllowance(RULE, 100001).allowance], [10, 10001]);
    assert.deepEqual([calcRetailAllowance(RULE, 120000).rate, calcRetailAllowance(RULE, 120000).allowance], [10, 12000]);
});

test('物販手当: 1円未満は切り上げ', () => {
    assert.equal(calcRetailAllowance(RULE, 1234).allowance, 62); // 61.7 → 62
});

test('物販手当: 10%の手前なら「あと何円」を返す', () => {
    assert.deepEqual(calcRetailAllowance(RULE, 80000).next, { threshold: 100000, rate: 10, remaining: 20000 });
    assert.equal(calcRetailAllowance(RULE, 120000).next, null);
});

test('物販手当: 開店・締め作業の調整はしない・総支給に入る・残業精算のインセンティブ部分には入れない', () => {
    const base = { patternKey: 'short_atsugi', serviceSales: 850000, hours: { overtime: 4, statutoryOvertime: 2, totalHours: 160.92 } };
    const withRetail = calcMonthlyPay(RULE, { ...base, storeDuty: yes, retailSales: 120000 });
    const noRetail = calcMonthlyPay(RULE, { ...base, storeDuty: yes, retailSales: 0 });
    assert.equal(withRetail.retailRate, 10);
    assert.equal(withRetail.retailAllowance, 12000);
    assert.equal(withRetail.overtimeTopUp, noRetail.overtimeTopUp, '物販手当は残業精算に影響しない');
    assert.equal(withRetail.total, noRetail.total + 12000);
    assert.equal(calcMonthlyPay(RULE, { ...base, storeDuty: no, retailSales: 120000 }).retailRate, 10);
});

// ---- 適用開始月 ----
test('適用開始月: 2026年10月分までは新ルールを使わない（旧ルールのまま）、11月分から新ルール', () => {
    assert.equal(NEW_RULE_START_MONTH, '2026-11');
    assert.equal(ruleForMonth('2026-10'), null);
    assert.equal(ruleForMonth('2025-12'), null);
    assert.equal(ruleForMonth('2026-11'), PAYROLL_RULES[0]);
    assert.equal(ruleForMonth('2027-09'), PAYROLL_RULES[0]);
    assert.equal(ruleForMonth('不正'), null);
    assert.equal(ruleForMonth('2026-13'), null);
    assert.equal(ruleForMonth('2026-00'), null);
});

test('適用開始月: 新しい設定ブロックを追加すると、その月からだけ新しい金額になる（過去の月は当時の金額）', () => {
    const raised = {
        ...PAYROLL_RULES[0],
        from: '2027-10',
        patterns: { ...PAYROLL_RULES[0].patterns, full: { ...PAYROLL_RULES[0].patterns.full, baseSalary: 240000 } },
    };
    const rules = [PAYROLL_RULES[0], raised];
    assert.equal(ruleForMonth('2027-09', rules).patterns.full.baseSalary, 230300);
    assert.equal(ruleForMonth('2027-10', rules).patterns.full.baseSalary, 240000);
    assert.equal(ruleForMonth('2028-01', rules).patterns.full.baseSalary, 240000);
    assert.equal(ruleForMonth('2026-10', rules), null);
    // 並び順に関係なく正しく選ぶ
    assert.equal(ruleForMonth('2027-09', [raised, PAYROLL_RULES[0]]).patterns.full.baseSalary, 230300);
});

test('設定値: 月平均所定労働時間は 269日 × 実働 ÷ 12 を小数第2位までに丸めた値', () => {
    const p = RULE.patterns;
    assert.equal(p.full.monthlyScheduledHours, Math.round(269 * 8 / 12 * 100) / 100);
    assert.equal(p.short_atsugi.monthlyScheduledHours, Math.round(269 * 7 / 12 * 100) / 100);
    assert.equal(p.short_chiba.monthlyScheduledHours, Math.round(269 * 6.75 / 12 * 100) / 100);
});

// ---- 給与パターンの履歴（◯月から）----
test('給与パターンの履歴: その月に有効な設定を使い、過去の月は当時の設定のまま', () => {
    const history = [
        { from: '2026-11', pattern: 'short_atsugi', storeDuty: false },
        { from: '2027-02', pattern: 'full', storeDuty: true },
        { from: '2027-06', pattern: null },
    ];
    assert.equal(settingForMonth(history, '2026-10'), null);
    assert.deepEqual(settingForMonth(history, '2026-12'), { from: '2026-11', pattern: 'short_atsugi', storeDuty: false, rates: null });
    assert.deepEqual(settingForMonth(history, '2027-02'), { from: '2027-02', pattern: 'full', storeDuty: true, rates: null });
    assert.equal(settingForMonth(history, '2027-07'), null, '対象外にした月以降は計算しない');
    assert.equal(settingForMonth(undefined, '2026-11'), null);
    assert.deepEqual(settingForMonth([{ from: '2026-11', pattern: 'full' }], '2026-11'), { from: '2026-11', pattern: 'full', storeDuty: null, rates: null });
    assert.deepEqual(settingForMonth([{ from: '2026-11', pattern: 'full', storeDuty: true, rates: [40, 41, 42, 43, 44] }], '2026-12').rates, [40, 41, 42, 43, 44]);
});

// ---- SalonOne の売上から施術・物販を取り出す ----
test('売上の取り出し: 施術 = 会計済み売上 − 物販。粗売上や表示設定は使わない', () => {
    const row = { gross_sales: 2000000, digest_sales: 1000000, consumed_sales: 900000, product_sales: 20000, refund_amount: 0 };
    assert.deepEqual(salesForPayroll(row, RULE, 0), { total: 1000000, retail: 20000, service: 980000, retailSource: 'salonone', refund: 0 });
});

test('売上の取り出し: SalonOne の物販が0なら日報入力タブの手入力値を使う', () => {
    const row = { digest_sales: 1000000, product_sales: 0 };
    assert.deepEqual(salesForPayroll(row, RULE, 30000), { total: 1000000, retail: 30000, service: 970000, retailSource: 'manual', refund: 0 });
    assert.equal(salesForPayroll({ digest_sales: 500000 }, RULE, 0).retailSource, 'none');
});

test('売上の取り出し: 返金は SalonOne の数字をそのまま使い、返金額は確認用に返すだけ', () => {
    const r = salesForPayroll({ digest_sales: 900000, product_sales: 0, refund_amount: 15000 }, RULE, 0);
    assert.equal(r.service, 900000);
    assert.equal(r.refund, 15000);
});

test('端数: 浮動小数点の誤差で1円多く切り上げない', () => {
    assert.equal(ceilYen(1143.0000000000002), 1143);
    assert.equal(ceilYen(1142.47), 1143);
    assert.equal(ceilYen(360.36), 361);
});

// ---- スタッフ個別の掛け率（段階ごとの率を入力。金額の段階はパターンのまま・開店締めの±1%は重ねない）----
const CUSTOM = [40, 41, 42, 43, 44]; // フルの段階（606,000 / 900,000 / 950,000 / 1,000,000 / 1,050,000 超）に対する個別の率

test('個別の掛け率: 段階の金額はパターンのまま、率だけ入力した値を使う', () => {
    const r = calcIncentive(RULE, 'full', 980000, true, CUSTOM);
    assert.equal(r.rate, 42); // 950,000超 → 3段目
    assert.equal(r.incentive, 157080); // (980,000 − 606,000) × 42%
    assert.equal(calcIncentive(RULE, 'full', 1100000, true, CUSTOM).rate, 44);
    assert.equal(calcIncentive(RULE, 'full', 900000, true, CUSTOM).rate, 40); // ちょうど同額は前の段階
});

test('個別の掛け率: 開店・締め作業の±1%は重ねない（入力した率をそのまま使う）', () => {
    assert.equal(calcIncentive(RULE, 'full', 980000, false, CUSTOM).rate, 42);
    assert.equal(calcIncentive(RULE, 'short_atsugi', 850000, true, [30, 31, 32, 33, 34]).rate, 33);
});

test('個別の掛け率: 開始ライン以下は0円のまま', () => {
    assert.deepEqual([calcIncentive(RULE, 'full', 606000, true, CUSTOM).rate, calcIncentive(RULE, 'full', 606000, true, CUSTOM).incentive], [0, 0]);
});

test('個別の掛け率: 次の段階も個別の率で表示する', () => {
    assert.equal(nextTierLabel(nextTier(RULE, 'full', 980000, true, CUSTOM)), '1,000,000超で43%（あと20,000円）');
    assert.equal(nextTierLabel(nextTier(RULE, 'full', 500000, false, CUSTOM)), '606,000超で40%（あと106,000円）');
});

test('個別の掛け率: 小数の率も使え、1円未満は切り上げ', () => {
    assert.equal(calcIncentive(RULE, 'short_chiba', 514001, false, [36.5, 37, 38, 39, 40]).incentive, 366); // 1,001 × 36.5% = 365.365
});

test('個別の掛け率: 段階の数が合わない・範囲外のときは使わず、パターンの率で計算する', () => {
    const p = RULE.patterns.full;
    assert.equal(validCustomRates(p, [40, 41]), null);
    assert.equal(validCustomRates(p, [40, 41, 42, 43, 101]), null);
    assert.equal(validCustomRates(p, [40, 41, 'x', 43, 44]), null);
    assert.equal(calcIncentive(RULE, 'full', 980000, true, [40, 41]).rate, 40);
});

test('個別の掛け率: 画面に出す「各段階の率」は、個別の率があればそれ、なければパターン＋開店締め', () => {
    assert.deepEqual(effectiveTierRates(RULE, 'full', true, null), [38, 39, 40, 41, 42]);
    assert.deepEqual(effectiveTierRates(RULE, 'full', false, null), [37, 38, 39, 40, 41]);
    assert.deepEqual(effectiveTierRates(RULE, 'short_chiba', true, null), [37, 38, 39, 40, 41]);
    assert.deepEqual(effectiveTierRates(RULE, 'full', false, CUSTOM), CUSTOM);
});

test('個別の掛け率: 総支給に反映され、残業精算のインセンティブ部分にも使われる', () => {
    const base = { patternKey: 'full', storeDuty: true, serviceSales: 980000, hours: { overtime: 6, statutoryOvertime: 6, totalHours: 185.33 } };
    const pay = calcMonthlyPay(RULE, { ...base, customRates: CUSTOM });
    assert.equal(pay.customRates, true);
    assert.deepEqual(pay.tierRates, CUSTOM);
    assert.equal(pay.incentive, 157080);
    assert.equal(pay.totalBeforeOvertime, 230300 + 9700 + 157080);
    assert.equal(pay.overtimeTopUp, calcOvertimeTopUp(RULE, 'full', 157080, 6, 6, 185.33));
    const normal = calcMonthlyPay(RULE, base);
    assert.equal(normal.customRates, false);
    assert.equal(normal.incentive, 149600, '個別の率がなければ今までどおり');
});

// ---- API（保存・権限）----
test('API: 設定は「その月から」の履歴として保存し、過去の月の設定は変えない', () => {
    const { applySettingsPatch } = payrollApi;
    let s = applySettingsPatch({}, { 101: { pattern: 'full', storeDuty: true } }, '2026-11');
    assert.deepEqual(s, { 101: [{ from: '2026-11', pattern: 'full', storeDuty: true }] });
    s = applySettingsPatch(s, { 101: { pattern: 'short_chiba', storeDuty: false } }, '2027-01');
    assert.deepEqual(s[101], [
        { from: '2026-11', pattern: 'full', storeDuty: true },
        { from: '2027-01', pattern: 'short_chiba', storeDuty: false },
    ]);
    // 同じ月の設定は置き換え
    s = applySettingsPatch(s, { 101: { pattern: 'short_chiba', storeDuty: true } }, '2027-01');
    assert.deepEqual(s[101][1], { from: '2027-01', pattern: 'short_chiba', storeDuty: true });
    // 直前と同じ設定は履歴に残さない
    s = applySettingsPatch(s, { 101: { pattern: 'full', storeDuty: true } }, '2026-12');
    assert.equal(s[101].length, 2);
    // null はその月の設定を取り消す
    s = applySettingsPatch(s, { 101: null }, '2027-01');
    assert.deepEqual(s[101], [{ from: '2026-11', pattern: 'full', storeDuty: true }]);
    // 対象外（pattern: null）
    s = applySettingsPatch(s, { 101: { pattern: null } }, '2027-03');
    assert.deepEqual(s[101][1], { from: '2027-03', pattern: null });
    // 最初から対象外なら何も残さない
    assert.deepEqual(applySettingsPatch({}, { 102: { pattern: null } }, '2026-11'), {});
});

test('API: 個別の掛け率は「その月から」の設定として保存し、パターンの率に戻すと消える', () => {
    const { applySettingsPatch } = payrollApi;
    let s = applySettingsPatch({}, { 101: { pattern: 'full', storeDuty: true } }, '2026-11');
    s = applySettingsPatch(s, { 101: { pattern: 'full', storeDuty: true, rates: [40, '41', 42.5, 43, 44] } }, '2026-12');
    assert.deepEqual(s[101], [
        { from: '2026-11', pattern: 'full', storeDuty: true },
        { from: '2026-12', pattern: 'full', storeDuty: true, rates: [40, 41, 42.5, 43, 44] },
    ]);
    // 同じ月に「パターンの率に戻す」（rates: null）→ 前の月と同じ設定になるので履歴から消える
    s = applySettingsPatch(s, { 101: { pattern: 'full', storeDuty: true, rates: null } }, '2026-12');
    assert.deepEqual(s[101], [{ from: '2026-11', pattern: 'full', storeDuty: true }]);
    // 対象外（pattern: null）には率を残さない
    s = applySettingsPatch(s, { 101: { pattern: null, rates: [40, 41, 42, 43, 44] } }, '2027-01');
    assert.deepEqual(s[101][1], { from: '2027-01', pattern: null });
});

test('API: 個別の掛け率の不正な値は保存しない', () => {
    const { applySettingsPatch } = payrollApi;
    assert.throws(() => applySettingsPatch({}, { 101: { pattern: 'full', rates: [40, 41, 42, 43, 120] } }, '2026-11'));
    assert.throws(() => applySettingsPatch({}, { 101: { pattern: 'full', rates: [40, -1, 42, 43, 44] } }, '2026-11'));
    assert.throws(() => applySettingsPatch({}, { 101: { pattern: 'full', rates: 'abc' } }, '2026-11'));
    assert.throws(() => applySettingsPatch({}, { 101: { pattern: 'full', rates: [] } }, '2026-11'));
    assert.throws(() => applySettingsPatch({}, { 101: { pattern: 'full', rates: new Array(11).fill(40) } }, '2026-11'));
});

test('API: 不正な値は保存しない', () => {
    const { applySettingsPatch, applyHoursPatch } = payrollApi;
    assert.throws(() => applySettingsPatch({}, { abc: { pattern: 'full' } }, '2026-11'));
    assert.throws(() => applySettingsPatch({}, { 101: { pattern: 'FULL; drop' } }, '2026-11'));
    assert.throws(() => applySettingsPatch({}, { 101: { pattern: 'full', storeDuty: 'yes' } }, '2026-11'));
    assert.throws(() => applyHoursPatch({}, { 101: { overtime: -1 } }));
    assert.throws(() => applyHoursPatch({}, { 101: { overtime: 1000 } }));
    assert.throws(() => applyHoursPatch({}, { 101: { salary: 1 } }));
});

test('API: 勤怠は小数第2位までで保存し、空欄は消す', () => {
    const { applyHoursPatch } = payrollApi;
    let h = applyHoursPatch({}, { 101: { overtime: '6', statutoryOvertime: 6, totalHours: 185.333 } });
    assert.deepEqual(h, { 101: { overtime: 6, statutoryOvertime: 6, totalHours: 185.33 } });
    h = applyHoursPatch(h, { 101: { overtime: '' } });
    assert.deepEqual(h, { 101: { statutoryOvertime: 6, totalHours: 185.33 } });
    h = applyHoursPatch(h, { 101: null });
    assert.deepEqual(h, {});
});

test('API: スタッフには本人の分だけ返す。マネージャー・店長は見られない', () => {
    const { scopeFor, canRead } = payrollApi;
    const settings = { 101: [{ from: '2026-11', pattern: 'full' }], 102: [{ from: '2026-11', pattern: 'short_chiba' }] };
    const hours = { 101: { totalHours: 180 }, 102: { totalHours: 150 } };
    assert.deepEqual(scopeFor({ role: 'admin' }, settings, hours), { settings, hours });
    assert.deepEqual(scopeFor({ role: 'staff', staffId: 102 }, settings, hours), {
        settings: { 102: settings[102] }, hours: { 102: hours[102] },
    });
    assert.deepEqual(scopeFor({ role: 'staff', staffId: 999 }, settings, hours), { settings: {}, hours: {} });
    // 給与に使う売上もスタッフには本人の行だけ
    const sales = { 101: { digest_sales: 900000 }, 102: { digest_sales: 500000 } };
    assert.deepEqual(scopeFor({ role: 'staff', staffId: 102 }, settings, hours, sales).sales, { 102: sales[102] });
    assert.deepEqual(scopeFor({ role: 'staff', staffId: 999 }, settings, hours, sales).sales, {});
    assert.equal(scopeFor({ role: 'staff', staffId: 102 }, settings, hours, null).sales, null);
    assert.deepEqual(scopeFor({ role: 'admin' }, settings, hours, sales).sales, sales);
    assert.equal(canRead({ role: 'admin' }), true);
    assert.equal(canRead({ role: 'staff', staffId: 101 }), true);
    assert.equal(canRead({ role: 'manager' }), false);
    assert.equal(canRead({ role: 'store', shopId: 1 }), false);
});

test('API: 給与に使う売上は全店舗のスタッフ別売上から売上の項目だけを取り出す（担当者なしの売上は除く）', () => {
    const { pickSales, monthRange } = payrollApi;
    const summary = {
        by_staff: [
            { staff_id: 101, staff_name: 'A', digest_sales: 980000, gross_sales: 1200000, consumed_sales: 950000, product_sales: 20000, refund_amount: 0, new_count: 5, customer_count: 40 },
            { staff_id: null, staff_name: null, digest_sales: 15000 },
        ],
    };
    assert.deepEqual(pickSales(summary), {
        101: { staff_name: 'A', digest_sales: 980000, consumed_sales: 950000, gross_sales: 1200000, product_sales: 20000, refund_amount: 0 },
    });
    assert.deepEqual(monthRange('2026-11'), { from: '2026-11-01', to: '2026-11-30' });
    assert.deepEqual(monthRange('2027-02'), { from: '2027-02-01', to: '2027-02-28' });
    assert.deepEqual(monthRange('2028-02'), { from: '2028-02-01', to: '2028-02-29' });
});
