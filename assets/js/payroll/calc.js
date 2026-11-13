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
 * スタッフ個別の掛け率（段階ごとの率・%）が使える形なら返す。使えなければ null（＝パターンの率を使う）
 *   段階の数がパターンと同じで、どれも 0〜100 の数のときだけ有効
 */
export function validCustomRates(pattern, rates) {
    if (!pattern || !Array.isArray(rates) || rates.length !== pattern.tiers.length) return null;
    const out = rates.map(Number);
    return out.every(r => isFinite(r) && r >= 0 && r <= 100) ? out : null;
}

/** スタッフ個別の基本給（円/月）が使える値なら返す。使えなければ null（＝パターンの基本給を使う） */
export function validBaseSalary(v) {
    if (v === null || v === undefined || v === '') return null;
    const n = Number(v);
    return isFinite(n) && n > 0 && n <= 2000000 ? Math.round(n) : null;
}

/**
 * 各段階で実際に使う率（%）
 *   個別の掛け率があればそのまま（開店・締めの±1%は重ねない）。なければパターンの率 ＋ 開店・締めの調整
 */
export function effectiveTierRates(rule, patternKey, storeDuty, customRates) {
    const p = requirePattern(rule, patternKey);
    const custom = validCustomRates(p, customRates);
    if (custom) return custom;
    const adj = storeDutyAdjust(rule, p, resolveDuty(p, storeDuty));
    return p.tiers.map(([, r]) => r + adj);
}

/**
 * 施術インセンティブ（全額方式）
 *   率: 税込施術売上が各段階の金額を「超えたら」その率（同額は前の段階のまま）
 *   インセンティブ = (売上 − 開始ライン) × 率、1円未満切り上げ。開始ライン以下は 0円
 *   customRates（スタッフ個別の段階ごとの率）があれば、パターンの率の代わりにそれを使う（開店・締めの調整なし）
 */
export function calcIncentive(rule, patternKey, sales, storeDuty, customRates = null) {
    const p = requirePattern(rule, patternKey);
    const duty = resolveDuty(p, storeDuty);
    const rates = effectiveTierRates(rule, patternKey, duty, customRates);
    const s = money(sales);
    const startLine = p.tiers[0][0];
    if (s <= startLine) return { rate: 0, incentive: 0, startLine, storeDuty: duty };
    let rate = 0;
    p.tiers.forEach(([threshold], i) => { if (s > threshold) rate = rates[i]; });
    return { rate, incentive: ceilYen(((s - startLine) * rate) / 100), startLine, storeDuty: duty };
}

/**
 * 次の段階（金額・率・あと何円か）。最上位の段階に到達していれば null（＝次の段階なし）
 *   あと何円 = 次の段階の金額 − 売上。売上がちょうど段階の金額と同じときは、超えるのに必要な 1円 と表示する
 */
export function nextTier(rule, patternKey, sales, storeDuty, customRates = null) {
    const p = requirePattern(rule, patternKey);
    const rates = effectiveTierRates(rule, patternKey, storeDuty, customRates);
    const s = money(sales);
    const i = p.tiers.findIndex(([threshold]) => s <= threshold);
    if (i < 0) return null;
    const threshold = p.tiers[i][0];
    return { threshold, rate: rates[i], remaining: Math.max(1, Math.ceil(threshold - s)) };
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
export function overtimeDetail(rule, patternKey, incentive, overtimeHours, statutoryOvertimeHours, totalWorkHours, baseSalary = null) {
    const p = requirePattern(rule, patternKey);
    const m = rule.overtime || { baseMultiplier: 1.25, incentiveMultiplier: 0.25 };
    // 個別の基本給があれば、残業の時給（基本給 ÷ 月平均所定）にもそれを使う
    const hourly = (validBaseSalary(baseSalary) ?? p.baseSalary) / p.monthlyScheduledHours;
    const total = hoursOf(totalWorkHours);
    const basePart = hourly * m.baseMultiplier * hoursOf(overtimeHours);
    const incentivePart = total > 0 ? (money(incentive) / total) * m.incentiveMultiplier * hoursOf(statutoryOvertimeHours) : 0;
    const topUp = Math.max(0, ceilYen(basePart + incentivePart - p.fixedOtPay));
    return { hourly, basePart, incentivePart, fixedOtPay: p.fixedOtPay, topUp };
}

export function calcOvertimeTopUp(rule, patternKey, incentive, overtimeHours, statutoryOvertimeHours, totalWorkHours, baseSalary = null) {
    return overtimeDetail(rule, patternKey, incentive, overtimeHours, statutoryOvertimeHours, totalWorkHours, baseSalary).topUp;
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
 * @param {number[]|null} [input.customRates] スタッフ個別の段階ごとの率（%）。あればパターンの率と開店・締めの調整の代わりに使う
 * @param {number|null} [input.baseSalary] スタッフ個別の基本給。あればパターンの基本給の代わりに使う（残業の時給にも使う）
 * @param {number} input.serviceSales     税込施術売上
 * @param {number} [input.retailSales]    税込物販売上
 * @param {object|null} [input.hours]     { overtime, statutoryOvertime, totalHours }（未入力なら null）
 */
export function calcMonthlyPay(rule, { patternKey, storeDuty = null, customRates = null, baseSalary = null, serviceSales = 0, retailSales = 0, hours = null }) {
    const p = requirePattern(rule, patternKey);
    const custom = validCustomRates(p, customRates);
    const customBase = validBaseSalary(baseSalary);
    const base = customBase ?? p.baseSalary;
    const inc = calcIncentive(rule, patternKey, serviceSales, storeDuty, custom);
    const next = nextTier(rule, patternKey, serviceSales, storeDuty, custom);
    const retail = calcRetailAllowance(rule, retailSales);
    const overtime = hasHours(hours)
        ? overtimeDetail(rule, patternKey, inc.incentive, hours.overtime, hours.statutoryOvertime, hours.totalHours, base)
        : null;
    const overtimeTopUp = overtime ? overtime.topUp : 0;
    const totalBeforeOvertime = base + p.fixedOtPay + inc.incentive + retail.allowance;
    return {
        patternKey,
        patternLabel: p.label,
        storeDuty: inc.storeDuty,
        storeDutyIsDefault: inc.storeDuty === p.defaultStoreDuty,
        customRates: !!custom,
        tierRates: effectiveTierRates(rule, patternKey, inc.storeDuty, custom),
        baseSalary: base,
        patternBaseSalary: p.baseSalary,
        customBaseSalary: customBase !== null,
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
 * スタッフ画面の「見通し」（やる気が出るように、今の位置と次の目標を数字で出す）
 *   - 今時点の率・インセンティブ
 *   - インセンティブ対象（開始ライン超え）まであと何円か
 *   - 次の段階まであと何円か・届いたらインセンティブがいくらになるか（今より＋いくらか）
 *   - 今のペース（今月の経過日数あたりの売上）のまま月末までいった場合の売上・率・インセンティブ
 *   - 残りの日数（今日を含む暦日）で次の段階に届くための1日あたりの売上
 *   - 段階の一覧（どこまで達成したか）
 * @param {object} input
 * @param {number} [input.dayOfMonth]  今日が何日か（今月を見ているときだけ。過去の月は省略 → ペースは出さない）
 * @param {number} [input.daysInMonth] その月の日数
 */
export function incentiveOutlook(rule, { patternKey, storeDuty = null, customRates = null, serviceSales = 0, dayOfMonth = null, daysInMonth = null }) {
    const p = requirePattern(rule, patternKey);
    const custom = validCustomRates(p, customRates);
    const rates = effectiveTierRates(rule, patternKey, storeDuty, custom);
    const sales = money(serviceSales);
    const cur = calcIncentive(rule, patternKey, sales, storeDuty, custom);
    const next = nextTier(rule, patternKey, sales, storeDuty, custom);
    const atNext = next ? calcIncentive(rule, patternKey, next.threshold + 1, storeDuty, custom).incentive : null;
    const eligible = sales > cur.startLine;

    const inMonth = Number.isInteger(dayOfMonth) && Number.isInteger(daysInMonth) && dayOfMonth >= 1 && dayOfMonth <= daysInMonth;
    const daysLeft = inMonth ? daysInMonth - dayOfMonth + 1 : null;
    let pace = null;
    if (inMonth && dayOfMonth < daysInMonth) {
        const projectedSales = Math.round((sales / dayOfMonth) * daysInMonth);
        const proj = calcIncentive(rule, patternKey, projectedSales, storeDuty, custom);
        pace = { projectedSales, projectedRate: proj.rate, projectedIncentive: proj.incentive };
    }
    const currentIndex = p.tiers.reduce((idx, [threshold], i) => (sales > threshold ? i : idx), -1);
    return {
        sales,
        startLine: cur.startLine,
        rate: cur.rate,
        incentive: cur.incentive,
        eligible,
        toStart: eligible ? 0 : Math.max(1, Math.ceil(cur.startLine - sales)),
        next,
        incentiveAtNext: atNext,
        gainAtNext: atNext === null ? null : atNext - cur.incentive,
        // 次の段階の率で、開始ラインを超えた売上1万円あたりいくらになるか（目安）
        perTenThousand: next ? Math.round(next.rate * 100) : Math.round(cur.rate * 100),
        daysLeft,
        perDayToNext: next && daysLeft ? Math.ceil(next.remaining / daysLeft) : null,
        pace,
        ladder: p.tiers.map(([threshold], i) => ({ threshold, rate: rates[i], reached: sales > threshold, current: i === currentIndex })),
    };
}

/**
 * 給与パターンの履歴から、その月に有効な設定を返す（無い・対象外なら null）
 * @param {Array<{from:string, pattern:string|null, storeDuty?:boolean, rates?:number[], baseSalary?:number}>} history  「◯月から」の設定の履歴
 * @param {string} monthKey 'YYYY-MM'
 */
export function settingForMonth(history, monthKey) {
    let found = null;
    for (const e of Array.isArray(history) ? history : []) {
        if (!e || typeof e.from !== 'string' || e.from > monthKey) continue;
        if (!found || e.from > found.from) found = e;
    }
    if (!found || !found.pattern) return null;
    return {
        from: found.from,
        pattern: found.pattern,
        storeDuty: typeof found.storeDuty === 'boolean' ? found.storeDuty : null,
        rates: Array.isArray(found.rates) ? found.rates : null,
        baseSalary: validBaseSalary(found.baseSalary),
    };
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
