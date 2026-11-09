// 給与・インセンティブの計算（2026年11月分〜の新ルール）
//
// 画面・データベース・通信には一切さわらない「計算だけ」の関数。
// 金額・率・時間はすべて config.js の設定ブロック（rule）から受け取る。
// テスト: tests/payroll.test.mjs（npm test）

/** 1円未満を切り上げる。浮動小数点の誤差（例: 1143.0000000001）で1円多くならないよう6桁で丸めてから切り上げる */
export function ceilYen(value) {
    return Math.ceil(Math.round(value * 1e6) / 1e6);
}

function money(v) {
    const n = Number(v);
    return isFinite(n) && n > 0 ? n : 0;
}

function hoursOf(v) {
    const n = Number(v);
    return isFinite(n) && n > 0 ? n : 0;
}

/** 設定ブロックから給与パターンを取り出す（無ければ null） */
export function getPattern(rule, patternKey) {
    return (rule && rule.patterns && Object.prototype.hasOwnProperty.call(rule.patterns, patternKey))
        ? rule.patterns[patternKey] : null;
}

function requirePattern(rule, patternKey) {
    const p = getPattern(rule, patternKey);
    if (!p) throw new Error(`unknown pay pattern: ${patternKey}`);
    return p;
}

/** 開店・締め作業がパターンの標準と違うときの率の調整（時短が担う → +1%、フルが担わない → −1%） */
function storeDutyAdjust(rule, pattern, storeDuty) {
    const a = rule.storeDutyRateAdjust ?? 1;
    return (storeDuty ? a : 0) - (pattern.defaultStoreDuty ? a : 0);
}

function resolveDuty(pattern, storeDuty) {
    return typeof storeDuty === 'boolean' ? storeDuty : pattern.defaultStoreDuty;
}

/**
 * 施術インセンティブ（全額方式）
 *   率: 税込施術売上が各段階の金額を「超えたら」その率（同額は前の段階のまま）
 *   インセンティブ = (売上 − 開始ライン) × 率、1円未満切り上げ。開始ライン以下は 0円
 */
export function calcIncentive(rule, patternKey, sales, storeDuty) {
    const p = requirePattern(rule, patternKey);
    const duty = resolveDuty(p, storeDuty);
    const s = money(sales);
    const startLine = p.tiers[0][0];
    if (s <= startLine) return { rate: 0, incentive: 0, startLine, storeDuty: duty };
    let rate = 0;
    for (const [threshold, r] of p.tiers) if (s > threshold) rate = r;
    rate += storeDutyAdjust(rule, p, duty);
    return { rate, incentive: ceilYen(((s - startLine) * rate) / 100), startLine, storeDuty: duty };
}

/**
 * 次の段階（金額・率・あと何円か）。最上位の段階に到達していれば null（＝次の段階なし）
 *   あと何円 = 次の段階の金額 − 売上。売上がちょうど段階の金額と同じときは、超えるのに必要な 1円 と表示する
 */
export function nextTier(rule, patternKey, sales, storeDuty) {
    const p = requirePattern(rule, patternKey);
    const duty = resolveDuty(p, storeDuty);
    const s = money(sales);
    const next = p.tiers.find(([threshold]) => s <= threshold);
    if (!next) return null;
    const [threshold, rate] = next;
    return { threshold, rate: rate + storeDutyAdjust(rule, p, duty), remaining: Math.max(1, Math.ceil(threshold - s)) };
}

function withCommas(n) {
    return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/** 次の段階の表示文（例: 「1,000,000超で41%（あと20,000円）」/「次の段階なし」） */
export function nextTierLabel(next) {
    if (!next) return '次の段階なし';
    return `${withCommas(next.threshold)}超で${next.rate}%（あと${withCommas(next.remaining)}円）`;
}

/**
 * 物販手当: 税込物販売上が各段階の金額を「超えたら」その率を物販売上の全額にかける（1円未満切り上げ）
 *   2026年11月〜: 0円超 → 5%、100,000円超 → 10%。開店・締め作業による調整はしない
 */
export function calcRetailAllowance(rule, retailSales) {
    const tiers = rule.retailTiers || [];
    const s = money(retailSales);
    let rate = 0;
    for (const [threshold, r] of tiers) if (s > threshold) rate = r;
    const allowance = rate > 0 ? ceilYen((s * rate) / 100) : 0;
    const nx = tiers.find(([threshold]) => threshold > 0 && s <= threshold);
    const next = nx ? { threshold: nx[0], rate: nx[1], remaining: Math.max(1, Math.ceil(nx[0] - s)) } : null;
    return { rate, allowance, next };
}

/**
 * 残業精算（固定残業代を超えた分）の内訳
 *   基本給部分     = 基本給 ÷ 月平均所定 × 1.25 × 残業時間（所定労働時間を超えた時間）
 *   インセンティブ部分 = 施術インセンティブ ÷ 総労働時間 × 0.25 × 法定時間外の時間（1日8h・週44hを超えた時間）
 *   追加支給       = 基本給部分 ＋ インセンティブ部分 − 固定残業代（マイナスは0円、1円未満切り上げ）
 *   物販手当はインセンティブ部分に含めない
 */
export function overtimeDetail(rule, patternKey, incentive, overtimeHours, statutoryOvertimeHours, totalWorkHours) {
    const p = requirePattern(rule, patternKey);
    const m = rule.overtime || { baseMultiplier: 1.25, incentiveMultiplier: 0.25 };
    const hourly = p.baseSalary / p.monthlyScheduledHours;
    const total = hoursOf(totalWorkHours);
    const basePart = hourly * m.baseMultiplier * hoursOf(overtimeHours);
    const incentivePart = total > 0 ? (money(incentive) / total) * m.incentiveMultiplier * hoursOf(statutoryOvertimeHours) : 0;
    const topUp = Math.max(0, ceilYen(basePart + incentivePart - p.fixedOtPay));
    return { hourly, basePart, incentivePart, fixedOtPay: p.fixedOtPay, topUp };
}

export function calcOvertimeTopUp(rule, patternKey, incentive, overtimeHours, statutoryOvertimeHours, totalWorkHours) {
    return overtimeDetail(rule, patternKey, incentive, overtimeHours, statutoryOvertimeHours, totalWorkHours).topUp;
}

/** 勤怠の手入力（残業・法定時間外・総労働時間）がどれか1つでも入っているか */
export function hasHours(hours) {
    if (!hours || typeof hours !== 'object') return false;
    return ['overtime', 'statutoryOvertime', 'totalHours'].some(k => hours[k] !== null && hours[k] !== undefined && hours[k] !== '' && isFinite(Number(hours[k])));
}

/**
 * スタッフ1人・1か月分の給与
 * @param {object} rule  config.js の設定ブロック（ruleForMonth の戻り値）
 * @param {object} input
 * @param {string} input.patternKey       給与パターン
 * @param {boolean|null} [input.storeDuty] 開店・締め作業を担うか（未指定はパターンの標準）
 * @param {number} input.serviceSales     税込施術売上
 * @param {number} [input.retailSales]    税込物販売上
 * @param {object|null} [input.hours]     { overtime, statutoryOvertime, totalHours }（未入力なら null）
 */
export function calcMonthlyPay(rule, { patternKey, storeDuty = null, serviceSales = 0, retailSales = 0, hours = null }) {
    const p = requirePattern(rule, patternKey);
    const inc = calcIncentive(rule, patternKey, serviceSales, storeDuty);
    const next = nextTier(rule, patternKey, serviceSales, storeDuty);
    const retail = calcRetailAllowance(rule, retailSales);
    const overtime = hasHours(hours)
        ? overtimeDetail(rule, patternKey, inc.incentive, hours.overtime, hours.statutoryOvertime, hours.totalHours)
        : null;
    const overtimeTopUp = overtime ? overtime.topUp : 0;
    const totalBeforeOvertime = p.baseSalary + p.fixedOtPay + inc.incentive + retail.allowance;
    return {
        patternKey,
        patternLabel: p.label,
        storeDuty: inc.storeDuty,
        storeDutyIsDefault: inc.storeDuty === p.defaultStoreDuty,
        baseSalary: p.baseSalary,
        fixedOtHours: p.fixedOtHours,
        fixedOtPay: p.fixedOtPay,
        serviceSales: money(serviceSales),
        rate: inc.rate,
        incentive: inc.incentive,
        startLine: inc.startLine,
        next,
        nextLabel: nextTierLabel(next),
        retailSales: money(retailSales),
        retailRate: retail.rate,
        retailAllowance: retail.allowance,
        retailNext: retail.next,
        hoursEntered: !!overtime,
        overtime,
        overtimeTopUp,
        totalBeforeOvertime,
        total: totalBeforeOvertime + overtimeTopUp,
    };
}

/**
 * 給与パターンの履歴から、その月に有効な設定を返す（無い・対象外なら null）
 * @param {Array<{from:string, pattern:string|null, storeDuty?:boolean}>} history  「◯月から」の設定の履歴
 * @param {string} monthKey 'YYYY-MM'
 */
export function settingForMonth(history, monthKey) {
    let found = null;
    for (const e of Array.isArray(history) ? history : []) {
        if (!e || typeof e.from !== 'string' || e.from > monthKey) continue;
        if (!found || e.from > found.from) found = e;
    }
    if (!found || !found.pattern) return null;
    return { from: found.from, pattern: found.pattern, storeDuty: typeof found.storeDuty === 'boolean' ? found.storeDuty : null };
}

/**
 * SalonOne のスタッフ別売上（by_staff の1行）から、給与計算に使う税込の施術売上・物販売上を取り出す
 *   施術売上 = 会計済み売上 − 物販売上
 *   物販売上 = SalonOne の product_sales（0 または無い場合は日報入力タブの手入力値）
 *   返金は SalonOne で管理する（会計済み売上の数字をそのまま使い、返金額は確認用に返すだけ）
 */
export function salesForPayroll(row, rule, manualRetail = 0) {
    const total = money(row?.[rule.salesField]);
    const apiRetail = money(row?.product_sales);
    const manual = money(manualRetail);
    const retailSource = apiRetail > 0 ? 'salonone' : manual > 0 ? 'manual' : 'none';
    const retail = Math.min(apiRetail > 0 ? apiRetail : manual, total);
    return { total, retail, service: Math.max(0, total - retail), retailSource, refund: money(row?.refund_amount) };
}
