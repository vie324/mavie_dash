// インセンティブ概算タブ（管理者専用）
// 従来ツールの計算式を踏襲:
//   施術売上(税込) = 粗売上 − 物販売上（物販は日報入力タブで手入力・未入力は0）
//   施術手当 = max(0, 施術売上(税抜) × 40% − 基本給)
//   物販手当 = 物販売上(税抜) × 10%
//   合計 = 基本給 + 施術手当 + 物販手当
// 税抜は従来ツールの計算（÷1.05）を踏襲。

import { state, on, isAdmin } from '../core/state.js';
import { yen, esc, monthLabel } from '../core/format.js';
import { kpisOf } from '../data/salonone.js';
import { getManual } from '../data/manual.js';

const SALARY_KEY = 'vie_base_salary_v1';
const TAX_DIVISOR = 1.05; // 従来ツールの業務ルールを踏襲
const SERVICE_RATE = 0.4;
const RETAIL_RATE = 0.1;

function loadSalaries() {
    try { return JSON.parse(localStorage.getItem(SALARY_KEY) || '{}'); } catch (_) { return {}; }
}
function saveSalaries(map) {
    try { localStorage.setItem(SALARY_KEY, JSON.stringify(map)); } catch (_) { /* ignore */ }
}

export function init() {
    on('data:core', render);
    on('data:manual', render);
    document.getElementById('incentive-table-body')?.addEventListener('change', ev => {
        const input = ev.target.closest('input[data-staff-id]');
        if (!input) return;
        const map = loadSalaries();
        const v = Number(input.value);
        if (isFinite(v) && v > 0) map[input.dataset.staffId] = v;
        else delete map[input.dataset.staffId];
        saveSalaries(map);
        render();
    });
}

function anchorMonthKey() {
    return `${state.filters.anchor.y}-${String(state.filters.anchor.m).padStart(2, '0')}`;
}

function render() {
    if (!isAdmin() || !state.data.summary) return;
    const body = document.getElementById('incentive-table-body');
    if (!body) return;

    // 基本給は月額のため、複数月の期間フィルタでは計算しない（給与関連の誤表示防止）
    if (state.filters.periodKind !== 'month') {
        setText('incentive-period', '');
        body.innerHTML = '<tr><td colspan="7" class="py-8 text-center text-surface-500">インセンティブは月単位で計算します。期間フィルタを「今月」にして対象月を選択してください。</td></tr>';
        return;
    }
    setText('incentive-period', `対象月: ${monthLabel(state.filters.anchor)}`);

    const byStaff = state.data.summary.by_staff || [];
    const salaries = loadSalaries();
    const manualMonthly = getManual(anchorMonthKey()).monthly || {};

    if (byStaff.length === 0) {
        body.innerHTML = '<tr><td colspan="7" class="py-8 text-center text-surface-500">この期間のスタッフ別データはありません</td></tr>';
        return;
    }

    const totals = { gross: 0, retail: 0, service: 0, sInc: 0, rInc: 0, sum: 0 };
    const rows = [...byStaff].sort((a, b) => (b.gross_sales || 0) - (a.gross_sales || 0)).map(r => {
        const k = kpisOf(r);
        const base = salaries[String(r.staff_id)] || 0;
        const retail = Math.min(manualMonthly[String(r.staff_id)]?.productSales || 0, k.gross);
        const service = Math.max(0, k.gross - retail);
        const serviceInc = Math.max(0, (service / TAX_DIVISOR) * SERVICE_RATE - base);
        const retailInc = (retail / TAX_DIVISOR) * RETAIL_RATE;
        const sum = base + serviceInc + retailInc;
        totals.gross += k.gross; totals.retail += retail; totals.service += service;
        totals.sInc += serviceInc; totals.rInc += retailInc; totals.sum += sum;
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(r.staff_name || '不明')}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(k.gross)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${retail ? '' : 'text-surface-400'}">${yen(retail)}</td>
            <td class="py-2 px-3 text-right">
                <input type="number" inputmode="numeric" min="0" step="1000" value="${base || ''}" placeholder="未設定"
                    data-staff-id="${r.staff_id}"
                    class="w-28 text-right px-2 py-1 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
            </td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold ${serviceInc > 0 ? 'text-sage-600' : 'text-surface-400'}">${yen(serviceInc)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${retailInc > 0 ? '' : 'text-surface-400'}">${yen(retailInc)}</td>
            <td class="py-2 px-3 text-right tabular-nums font-bold">${yen(sum)}</td>
        </tr>`;
    });

    rows.push(`
        <tr class="bg-surface-50 dark:bg-gray-800/50 font-semibold">
            <td class="py-2 px-3">合計</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.gross)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.retail)}</td>
            <td class="py-2 px-3"></td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.sInc)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.rInc)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.sum)}</td>
        </tr>`);
    body.innerHTML = rows.join('');
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
