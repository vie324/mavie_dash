// 売上目標設定タブ（管理者専用）
// 目標はlocalStorageに保存。店舗ごとにカードを作り、店舗目標+所属スタッフ目標を編集できる。

import { state, on, emit, isAdmin } from '../core/state.js';
import { esc, monthLabel } from '../core/format.js';
import { getGoalRaw, setGoal, monthKey, exportGoals, importGoals } from '../data/goals.js';
import { toast } from '../core/engage.js';

const FIELDS = [
    { key: 'sales', label: '売上目標（円）', step: 100000, placeholder: '例: 1100000' },
    { key: 'newVisits', label: '新規来店（名）', step: 5, placeholder: '例: 30' },
    { key: 'joins', label: '入会数（名）', step: 1, placeholder: '例: 10' },
];

export function init() {
    on('masters', render);
    on('data:core', render);
    document.getElementById('goal-editor')?.addEventListener('change', onEdit);
    document.getElementById('goal-export-btn')?.addEventListener('click', doExport);
    document.getElementById('goal-import-btn')?.addEventListener('click', () => document.getElementById('goal-import-file')?.click());
    document.getElementById('goal-import-file')?.addEventListener('change', doImport);
}

function mk() {
    return monthKey(state.filters.anchor);
}

function render() {
    if (!isAdmin()) return;
    const editor = document.getElementById('goal-editor');
    if (!editor) return;
    setText('goal-month-label', `${monthLabel(state.filters.anchor)}の目標`);

    const key = mk();
    editor.innerHTML = state.masters.shops.map(shop => {
        const shopGoal = getGoalRaw(key, `shop:${shop.id}`) || {};
        const staffs = state.masters.staffs.filter(s => String(s.shop_id) === String(shop.id));
        return `
        <div class="border border-surface-200 dark:border-gray-700 rounded-xl p-4">
            <h4 class="font-display font-bold text-accent-900 mb-3">${esc(shop.name)}</h4>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-3 mb-4">
                ${FIELDS.map(f => `
                    <div>
                        <label class="block text-xs text-surface-500 mb-1">店舗${f.label}</label>
                        <input type="number" inputmode="numeric" min="0" step="${f.step}" value="${shopGoal[f.key] || ''}"
                            placeholder="${f.placeholder}"
                            data-scope="shop:${shop.id}" data-field="${f.key}"
                            class="w-full px-3 py-2 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                    </div>`).join('')}
            </div>
            ${staffs.length ? `
            <div class="overflow-x-auto">
                <table class="w-full text-sm whitespace-nowrap">
                    <thead>
                        <tr class="border-b border-surface-200 dark:border-gray-700 text-surface-500">
                            <th class="text-left py-2 px-2 font-semibold">スタッフ</th>
                            ${FIELDS.map(f => `<th class="text-right py-2 px-2 font-semibold">${f.label}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${staffs.map(st => {
                            const g = getGoalRaw(key, `staff:${st.id}`) || {};
                            return `<tr class="border-b border-surface-100 dark:border-gray-800">
                                <td class="py-2 px-2 font-medium">${esc(st.name)}</td>
                                ${FIELDS.map(f => `
                                    <td class="py-1.5 px-2 text-right">
                                        <input type="number" inputmode="numeric" min="0" step="${f.step}" value="${g[f.key] || ''}"
                                            data-scope="staff:${st.id}" data-field="${f.key}"
                                            class="w-28 text-right px-2 py-1 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                                    </td>`).join('')}
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>` : ''}
        </div>`;
    }).join('');
}

function onEdit(ev) {
    const input = ev.target.closest('input[data-scope]');
    if (!input) return;
    const scope = input.dataset.scope;
    const key = mk();
    const current = getGoalRaw(key, scope) || {};
    const v = Number(input.value);
    if (isFinite(v) && v > 0) current[input.dataset.field] = v;
    else delete current[input.dataset.field];
    setGoal(key, scope, current);
    toast('目標を保存しました');
    emit('data:core'); // リング・ペースを再描画
}

function doExport() {
    const blob = new Blob([exportGoals()], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `vie_goals_${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

async function doImport(ev) {
    const file = ev.target.files?.[0];
    if (!file) return;
    try {
        importGoals(await file.text());
        toast('目標をインポートしました');
        render();
        emit('data:core');
    } catch (e) {
        toast('インポートに失敗しました: ' + e.message, 'error');
    } finally {
        ev.target.value = '';
    }
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
