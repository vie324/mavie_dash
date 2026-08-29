// インセンティブ概算タブ（管理者専用）
// 施術手当 = max(0, 施術売上(税抜) × 40% − 基本給)
// SalonOne APIは施術/物販の内訳を持たないため粗売上全体で概算する（画面に注記あり）。
// 税率は従来ツールの計算（÷1.05）を踏襲。

import { state, on, isAdmin } from '../core/state.js';
import { yen, num, esc, monthLabel } from '../core/format.js';
import { kpisOf } from '../data/salonone.js';

const SALARY_KEY = 'vie_base_salary_v1';
const TAX_DIVISOR = 1.05; // 従来ツールの業務ルールを踏襲
const SERVICE_RATE = 0.4;

function loadSalaries() {
    try { return JSON.parse(localStorage.getItem(SALARY_KEY) || '{}'); } catch (_) { return {}; }
}
function saveSalaries(map) {
    try { localStorage.setItem(SALARY_KEY, JSON.stringify(map)); } catch (_) { /* ignore */ }
}

export function init() {
    on('data:core', render);
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

function render() {
    if (!isAdmin() || !state.data.summary) return;
    const body = document.getElementById('incentive-table-body');
    if (!body) return;

    // 基本給は月額のため、複数月の期間フィルタでは計算しない（給与関連の誤表示防止）
    if (state.filters.periodKind !== 'month') {
        setText('incentive-period', '');
        body.innerHTML = '<tr><td colspan="6" class="py-8 text-center text-surface-500">インセンティブは月単位で計算します。期間フィルタを「今月」にして対象月を選択してください。</td></tr>';
        return;
    }
    setText('incentive-period', `対象月: ${monthLabel(state.filters.anchor)}`);

    const byStaff = state.data.summary.by_staff || [];
    const salaries = loadSalaries();
    const staffMeta = new Map(state.masters.staffs.map(s => [String(s.id), s]));

    if (byStaff.length === 0) {
        body.innerHTML = '<tr><td colspan="6" class="py-8 text-center text-surface-500">この期間のスタッフ別データはありません</td></tr>';
        return;
    }

    let totals = { gross: 0, net: 0, incentive: 0, sum: 0 };
    const rows = [...byStaff].sort((a, b) => (b.gross_sales || 0) - (a.gross_sales || 0)).map(r => {
        const k = kpisOf(r);
        const base = salaries[String(r.staff_id)] || 0;
        const net = k.gross / TAX_DIVISOR;
        const incentive = Math.max(0, net * SERVICE_RATE - base);
        const sum = base + incentive;
        totals.gross += k.gross; totals.net += net; totals.incentive += incentive; totals.sum += sum;
        const shop = staffMeta.get(String(r.staff_id));
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(r.staff_name || shop?.name || '不明')}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(k.gross)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(net)}</td>
            <td class="py-2 px-3 text-right">
                <input type="number" inputmode="numeric" min="0" step="1000" value="${base || ''}" placeholder="未設定"
                    data-staff-id="${r.staff_id}"
                    class="w-28 text-right px-2 py-1 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
            </td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold ${incentive > 0 ? 'text-sage-600' : 'text-surface-400'}">${yen(incentive)}</td>
            <td class="py-2 px-3 text-right tabular-nums font-bold">${yen(sum)}</td>
        </tr>`;
    });

    rows.push(`
        <tr class="bg-surface-50 dark:bg-gray-800/50 font-semibold">
            <td class="py-2 px-3">合計</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.gross)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.net)}</td>
            <td class="py-2 px-3"></td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.incentive)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(totals.sum)}</td>
        </tr>`);
    body.innerHTML = rows.join('');
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
