// インセンティブ概算タブ（管理者専用）
// 従来ツールの計算式を踏襲:
//   施術売上(税込) = 売上 − 物販売上（物販はSalonOneの product_sales。API側に無い/0の場合は日報入力タブの手入力値）
//   施術手当 = max(0, 施術売上(税抜) × 40% − 基本給)
//   物販手当 = 物販売上(税抜) × 10%
//   合計 = 基本給 + 施術手当 + 物販手当
// 税抜は従来ツールの計算（÷1.05）を踏襲。基本給はサーバー（/api/goals）に保存する。

import { state, on, isAdmin } from '../core/state.js';
import { yen, esc, monthLabel } from '../core/format.js';
import { kpisOf, salesOf } from '../data/salonone.js';
import { getManual } from '../data/manual.js';
import { getSalaries, setSalary, goalsStorage } from '../data/goals.js';
import { toast } from '../core/engage.js';

const TAX_DIVISOR = 1.05; // 従来ツールの業務ルールを踏襲
const SERVICE_RATE = 0.4;
const RETAIL_RATE = 0.1;

export function init() {
    on('data:core', render);
    on('data:manual', render);
    on('data:goals', () => {
        const body = document.getElementById('incentive-table-body');
        if (body && body.contains(document.activeElement)) return; // 入力中は再描画しない
        render();
    });
    document.getElementById('incentive-table-body')?.addEventListener('change', async ev => {
        const input = ev.target.closest('input[data-staff-id]');
        if (!input) return;
        input.disabled = true;
        try {
            const res = await setSalary(input.dataset.staffId, input.value);
            toast(res.storage === 'local' ? '基本給を保存しました（この端末のみ）' : '基本給を保存しました', 'success');
            render();
        } catch (e) {
            console.error('salary save', e);
            toast(e?.body?.detail || '基本給の保存に失敗しました', 'error');
        } finally {
            input.disabled = false;
        }
    });
}

function anchorMonthKey() {
    return `${state.filters.anchor.y}-${String(state.filters.anchor.m).padStart(2, '0')}`;
}

function render() {
    if (!isAdmin() || !state.data.summary) return;
    const body = document.getElementById('incentive-table-body');
    if (!body) return;

    const note = document.getElementById('incentive-storage-note');
    if (note) note.textContent = goalsStorage() === 'local'
        ? '基本給はこの端末にのみ保存されています（サーバー保存が未設定）。変更すると自動保存されます。'
        : '基本給はサーバーに保存されます（オーナーのみ閲覧・変更可）。変更すると自動保存されます。';

    // 基本給は月額のため、複数月の期間フィルタでは計算しない（給与関連の誤表示防止）
    if (state.filters.periodKind !== 'month') {
        setText('incentive-period', '');
        body.innerHTML = '<tr><td colspan="7" class="py-8 text-center text-surface-500">インセンティブは月単位で計算します。期間フィルタを「今月」にして対象月を選択してください。</td></tr>';
        return;
    }
    setText('incentive-period', `対象月: ${monthLabel(state.filters.anchor)}`);

    const byStaff = state.data.summary.by_staff || [];
    const salaries = getSalaries();
    const manualMonthly = getManual(anchorMonthKey()).monthly || {};

    if (byStaff.length === 0) {
        body.innerHTML = '<tr><td colspan="7" class="py-8 text-center text-surface-500">この期間のスタッフ別データはありません</td></tr>';
        return;
    }

    const totals = { gross: 0, retail: 0, service: 0, sInc: 0, rInc: 0, sum: 0 };
    const rows = [...byStaff].sort((a, b) => salesOf(b) - salesOf(a)).map(r => {
        const k = kpisOf(r);
        const base = salaries[String(r.staff_id)] || 0;
        // 物販売上: 実APIの product_sales（>0）を優先し、無い場合は日報入力タブの手入力値
        const apiRetail = (r.product_sales !== undefined && r.product_sales !== null) ? Number(r.product_sales) : null;
        const manualRetail = manualMonthly[String(r.staff_id)]?.productSales || 0;
        const useManual = !(apiRetail > 0) && manualRetail > 0;
        const retailRaw = apiRetail > 0 ? apiRetail : manualRetail;
        const retail = Math.min(retailRaw, k.sales);
        const service = Math.max(0, k.sales - retail);
        const serviceInc = Math.max(0, (service / TAX_DIVISOR) * SERVICE_RATE - base);
        const retailInc = (retail / TAX_DIVISOR) * RETAIL_RATE;
        const sum = base + serviceInc + retailInc;
        totals.gross += k.sales; totals.retail += retail; totals.service += service;
        totals.sInc += serviceInc; totals.rInc += retailInc; totals.sum += sum;
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(r.staff_name || '不明')}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(k.sales)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${retail ? '' : 'text-surface-400'}">${yen(retail)}${useManual ? ' <span class="text-[9px] text-surface-400 align-middle">手入力</span>' : ''}</td>
            <td class="py-2 px-3 text-right">
                <input type="number" inputmode="numeric" min="0" step="1000" value="${base || ''}" placeholder="未設定"
                    data-staff-id="${r.staff_id}" aria-label="${esc(r.staff_name || '')}の基本給"
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
