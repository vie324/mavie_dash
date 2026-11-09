// 給与・インセンティブの設定値（金額・率・時間）はこのファイルにだけ書く。
//
// ■ 適用開始月のしくみ
//   PAYROLL_RULES の各ブロックに from（適用開始月 'YYYY-MM'）を持たせる。
//   ある月の計算には「from がその月以前のブロックのうち、いちばん新しいもの」を使う。
//   いちばん古いブロックの from より前の月（= 2026年10月分まで）は、旧ルール
//   （インセンティブタブの従来の計算: 施術40% − 基本給・物販10%）のまま計算する。
//
// ■ 金額を変えるとき（例: 毎年10月の最低賃金改定）
//   既存のブロックは書き換えずに残し、末尾に from: '2027-10' のような新しいブロックを追加する。
//   こうすると過去の月は当時の金額のまま、新しい月だけ新しい金額で計算される。
//   給与パターンのキー（full / short_atsugi / short_chiba）はスタッフの設定と結びついているので変えないこと。

/**
 * @typedef {Object} PayPattern
 * @property {string}  label                  画面に出す名前
 * @property {string}  hours                  勤務時間の説明（表示用）
 * @property {number}  baseSalary             基本給（円/月）
 * @property {number}  fixedOtHours           固定残業時間（時間/月）
 * @property {number}  fixedOtPay             固定残業代（円/月）
 * @property {number}  monthlyScheduledHours  月平均所定労働時間（年間出勤日数 × 1日の実働 ÷ 12、小数第2位まで）
 * @property {boolean} defaultStoreDuty       標準で開店・締め作業を担うか
 * @property {[number, number][]} tiers       [この税込施術売上を超えたら, 率(%)]。1行目が開始ライン
 */

export const PAYROLL_RULES = [
    {
        from: '2026-11',
        label: '2026年11月〜',
        // 売上の基準: SalonOne の会計済み売上（税込）。端末ごとの「売上の基準」の表示設定には左右されない
        salesField: 'digest_sales',
        // 開店・締め作業がパターンの標準と違うときの率の調整（%）。施術インセンティブだけに適用
        storeDutyRateAdjust: 1,
        // 残業精算の割増率
        overtime: {
            baseMultiplier: 1.25,       // 基本給部分: 基本給 ÷ 月平均所定 × 1.25 × 残業時間
            incentiveMultiplier: 0.25,  // インセンティブ部分: 施術インセンティブ ÷ 総労働時間 × 0.25 × 法定時間外の時間
        },
        // 物販手当: 税込物販売上がこの金額を「超えたら」その率を物販売上の全額にかける（1円未満切り上げ）
        retailTiers: [[0, 5], [100000, 10]],
        patterns: {
            full: {
                label: 'フル', hours: 'シフト制（実働8h・休憩45分）',
                baseSalary: 230300, fixedOtHours: 6, fixedOtPay: 9700,
                monthlyScheduledHours: 179.33, defaultStoreDuty: true,
                tiers: [[606000, 38], [900000, 39], [950000, 40], [1000000, 41], [1050000, 42]],
            },
            short_atsugi: {
                label: '時短・本厚木', hours: '10:00〜17:45（実働7h・休憩45分）',
                baseSalary: 212600, fixedOtHours: 3, fixedOtPay: 5100,
                monthlyScheduledHours: 156.92, defaultStoreDuty: false,
                tiers: [[532000, 36], [559000, 37], [760000, 38], [830000, 39], [870000, 40]],
            },
            short_chiba: {
                label: '時短・千葉', hours: '10:00〜17:30（実働6h45m・休憩45分）',
                baseSalary: 205000, fixedOtHours: 3, fixedOtPay: 5100,
                monthlyScheduledHours: 151.31, defaultStoreDuty: false,
                tiers: [[513000, 36], [539000, 37], [760000, 38], [830000, 39], [870000, 40]],
            },
        },
    },
    // 例: 2027年10月に金額を変えるときは、ここに上と同じ形のブロックを from: '2027-10' で追加する
];

// 新ルールが始まる月（これより前は旧ルール）
export const NEW_RULE_START_MONTH = PAYROLL_RULES.map(r => r.from).sort()[0];

const MONTH_RE = /^\d{4}-(0[1-9]|1[0-2])$/;

/**
 * その月に使う設定ブロックを返す。旧ルールの月（2026年10月分まで）は null。
 * @param {string} monthKey 'YYYY-MM'
 * @param {Array} [rules] テスト用に差し替え可能
 */
export function ruleForMonth(monthKey, rules = PAYROLL_RULES) {
    if (!MONTH_RE.test(String(monthKey || ''))) return null;
    let found = null;
    for (const r of rules) {
        if (r.from <= monthKey && (!found || r.from > found.from)) found = r;
    }
    return found;
}
