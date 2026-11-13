// スタッフ向け「今月のインセンティブ」の見せ方（マイ成績とホームで共通）
//   今時点でいくらか / インセンティブ対象まであといくらか / 次の段階まであといくら・届くといくら増えるか /
//   今のペースだと月末いくらか / 段階のどこまで来たか を、やる気が出るように大きく・わかりやすく出す。
//   数字は payroll/calc.js の incentiveOutlook（画面と切り離した計算）から作る。

import { yen, esc, todayJst, daysInMonth } from '../core/format.js';
import { computeStaffPay } from '../data/payroll.js';
import { incentiveOutlook } from '../payroll/calc.js';

/**
 * スタッフ1人・1か月分の給与と見通しを作る（給与パターン未設定・旧ルールの月は null）
 * @param {string} month 'YYYY-MM'
 */
export function staffOutlook(month, staffId, fallbackRow) {
    const r = computeStaffPay(month, staffId, fallbackRow || {});
    if (!r?.pay) return null;
    const [y, m] = month.split('-').map(Number);
    const t = todayJst();
    const isCurrent = t.y === y && t.m === m;
    const outlook = incentiveOutlook(r.rule, {
        patternKey: r.setting.pattern,
        storeDuty: r.setting.storeDuty,
        customRates: r.setting.rates,
        serviceSales: r.sales.service,
        dayOfMonth: isCurrent ? t.d : null,
        daysInMonth: daysInMonth(y, m),
    });
    return { r, pay: r.pay, outlook, isCurrent, asOf: isCurrent ? `${t.m}/${t.d}時点` : `${m}月分` };
}

function progressPct(o) {
    if (!o.next) return 100;
    const reached = o.ladder.filter(t => t.reached);
    const from = reached.length ? reached[reached.length - 1].threshold : 0;
    const span = o.next.threshold - from;
    return span > 0 ? Math.max(0, Math.min(100, ((o.sales - from) / span) * 100)) : 0;
}

function heroHtml(o, pay) {
    if (o.eligible) {
        return `
            <p class="text-[11px] text-surface-500">今時点のインセンティブ</p>
            <p class="text-3xl md:text-4xl font-display font-bold text-accent-900 tabular-nums leading-tight">${yen(o.incentive)}</p>
            <p class="text-xs text-surface-500 mt-1">税込施術売上 <span class="tabular-nums">${yen(o.sales)}</span> ・ 率 <span class="font-semibold text-accent-900">${o.rate}%</span>${pay.customRates ? '（個別の率）' : ''}</p>`;
    }
    return `
        <p class="text-[11px] text-surface-500">インセンティブ対象まで</p>
        <p class="text-3xl md:text-4xl font-display font-bold text-primary-500 tabular-nums leading-tight">あと ${yen(o.toStart)}</p>
        <p class="text-xs text-surface-500 mt-1">税込施術売上 <span class="tabular-nums">${yen(o.sales)}</span>。${yen(o.startLine)}を超えると、超えた分の${o.next.rate}%がインセンティブになります（1万円ごとに約${yen(o.perTenThousand)}）。</p>`;
}

function goalHtml(o) {
    if (!o.next) {
        return `
            <div class="mt-4 rounded-xl bg-surface-50 dark:bg-gray-700/40 p-3 text-sm font-semibold text-sage-600">
                最上位の段階（${o.rate}%）に到達しています。すばらしいです！
            </div>`;
    }
    const title = o.eligible
        ? `あと <span class="tabular-nums">${yen(o.next.remaining)}</span> で <span class="text-primary-500">${o.next.rate}%</span> にアップ`
        : `あと <span class="tabular-nums">${yen(o.toStart)}</span> で <span class="text-primary-500">${o.next.rate}%</span> スタート`;
    return `
        <div class="mt-4 rounded-xl bg-surface-50 dark:bg-gray-700/40 p-3">
            <div class="flex items-baseline justify-between gap-2">
                <p class="text-sm font-semibold text-accent-900">${title}</p>
                <p class="text-[10px] text-surface-500 flex-shrink-0">${yen(o.next.threshold)}超</p>
            </div>
            <div class="h-2.5 rounded-full bg-surface-200 dark:bg-gray-600 overflow-hidden mt-2">
                <div class="h-full rounded-full gradient-warmgold" style="width:${progressPct(o).toFixed(1)}%"></div>
            </div>
            ${o.eligible && o.gainAtNext > 0 ? `<p class="text-xs text-accent-800 dark:text-gray-200 mt-2">届くとインセンティブは <span class="font-semibold tabular-nums">${yen(o.incentiveAtNext)}</span>（今より <span class="font-semibold text-sage-600 tabular-nums">＋${yen(o.gainAtNext)}</span>）</p>` : ''}
            ${o.perDayToNext ? `<p class="text-[11px] text-surface-500 mt-1">残り${o.daysLeft}日（今日を含む）で、1日あたり約 <span class="tabular-nums">${yen(o.perDayToNext)}</span></p>` : ''}
        </div>`;
}

function paceHtml(o) {
    if (!o.pace) return '';
    const p = o.pace;
    const result = p.projectedRate > 0
        ? `${p.projectedRate}%・インセンティブ 約<span class="font-semibold tabular-nums">${yen(p.projectedIncentive)}</span>`
        : `インセンティブ対象（${yen(o.startLine)}超）まであと少し`;
    return `<p class="text-[11px] text-surface-500 mt-3">このペースなら月末は 売上 約<span class="tabular-nums">${yen(p.projectedSales)}</span> → ${result}</p>`;
}

function ladderHtml(o) {
    return `
        <div class="mt-4">
            <p class="text-[11px] text-surface-500 mb-1">インセンティブの段階</p>
            <ol class="space-y-1">
                ${o.ladder.map(t => `
                    <li class="flex items-center justify-between gap-2 rounded-lg px-3 py-1.5 text-sm ${t.current ? 'bg-amber-50 dark:bg-amber-900/20 font-semibold' : ''}">
                        <span class="flex items-center gap-2 tabular-nums ${t.reached ? 'text-accent-900' : 'text-surface-400'}">
                            <span class="${t.reached ? 'text-sage-600' : 'text-surface-300'}" aria-hidden="true">${t.reached ? '✓' : '○'}</span>${yen(t.threshold)}超
                        </span>
                        <span class="tabular-nums ${t.reached ? 'text-accent-900' : 'text-surface-400'}">${t.rate}%${t.current ? '<span class="ml-1 text-[10px] text-primary-500">いまここ</span>' : ''}</span>
                    </li>`).join('')}
            </ol>
        </div>`;
}

/** マイ成績用（すべて）。compact: ホーム用（今いくら・次の目標・ペースだけ） */
export function outlookHtml({ pay, outlook: o, asOf }, { compact = false } = {}) {
    const note = '<p class="text-[10px] text-surface-400 mt-3">SalonOneの会計済み売上（会計が済んだ分）から計算した見込みです。確定した金額は給与明細でご確認ください。</p>';
    if (compact) {
        return `
            <div class="flex items-start justify-between gap-3">
                <div class="min-w-0">${heroHtml(o, pay)}</div>
                <span class="text-[10px] text-surface-500 flex-shrink-0">${esc(asOf)}</span>
            </div>
            ${goalHtml(o)}
            ${paceHtml(o)}`;
    }
    const cell = (label, value, sub) => `
        <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-3">
            <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">${label}</p>
            <p class="text-base font-display font-bold text-accent-900 tabular-nums">${value}</p>
            ${sub ? `<p class="text-[10px] text-surface-500 mt-0.5">${sub}</p>` : ''}
        </div>`;
    return `
        ${heroHtml(o, pay)}
        ${goalHtml(o)}
        ${paceHtml(o)}
        ${ladderHtml(o)}
        <div class="mt-5 flex items-baseline justify-between gap-3">
            <p class="text-sm font-semibold text-accent-900">総支給の見込み</p>
            <p class="text-xl font-display font-bold text-accent-900 tabular-nums">${yen(pay.total)}</p>
        </div>
        ${pay.hoursEntered ? '' : '<p class="text-[10px] text-surface-500 text-right">残業の精算は勤怠の入力後に反映されます</p>'}
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mt-3">
            ${cell('基本給', yen(pay.baseSalary), esc(pay.patternLabel))}
            ${cell('固定残業代', yen(pay.fixedOtPay), `${pay.fixedOtHours}時間分`)}
            ${cell('物販手当', yen(pay.retailAllowance), pay.retailSales > 0 ? `税込物販 ${yen(pay.retailSales)} × ${pay.retailRate}%` : '物販 5%（10万円超で10%）')}
            ${cell('残業の追加支給', pay.hoursEntered ? yen(pay.overtimeTopUp) : '—', pay.hoursEntered ? '固定残業代を超えた分' : '勤怠の入力待ち')}
        </div>
        ${note}`;
}
