// 売上目標設定タブ（オーナー・マネージャー・店長）
// 目標はサーバー（Supabase / Upstash）に保存され、全端末・スタッフのマイダッシュボードに反映される。
// サーバー保存が未設定の場合のみ、この端末のlocalStorageに退避する。

import { state, on, emit, isAdmin, isManager, isStoreLocked } from '../core/state.js';
import { esc, monthLabel, yen, num, todayJst } from '../core/format.js';
import { getGoalRaw, setGoal, monthKey, exportGoals, importGoals, goalsStorage } from '../data/goals.js';
import { kpisOf } from '../data/salonone.js';
import { monthlyTotalsByStaff } from '../data/manual.js';
import { toast } from '../core/engage.js';

// 次回予約数 = 日報の次回予約（新規 + 2回目以降）の月合計
const FIELDS = [
    { key: 'sales', label: '売上目標（円）', step: 100000, placeholder: '例: 1100000', fmt: v => yen(v) },
    { key: 'newVisits', label: '新規来店（名）', step: 5, placeholder: '例: 30', fmt: v => `${num(v)}名` },
    { key: 'nextBookings', label: '次回予約数（件）', step: 5, placeholder: '例: 60', fmt: v => `${num(v)}件` },
];

export function init() {
    on('masters', render);
    on('data:core', render);
    // 入力中に再描画するとフォーカスが飛ぶため、編集中は値をそのまま維持する
    const renderUnlessEditing = () => {
        const editor = document.getElementById('goal-editor');
        if (editor && editor.contains(document.activeElement)) return;
        render();
    };
    on('data:goals', renderUnlessEditing);
    on('data:manual', renderUnlessEditing);
    document.getElementById('goal-editor')?.addEventListener('change', onEdit);
    document.getElementById('goal-export-btn')?.addEventListener('click', doExport);
    document.getElementById('goal-import-btn')?.addEventListener('click', () => document.getElementById('goal-import-file')?.click());
    document.getElementById('goal-import-file')?.addEventListener('change', doImport);
}

function mk() {
    return monthKey(state.filters.anchor);
}

// 対象月の実績（単月表示のときはサマリ、それ以外は今月の実績）を目安として表示する
// 次回予約数は対象月の日報の合計（新規 + 2回目以降）
function actualsFor() {
    const t = todayJst();
    const isAnchorNow = state.filters.anchor.y === t.y && state.filters.anchor.m === t.m;
    const summary = state.filters.periodKind === 'month' ? state.data.summary : (isAnchorNow ? state.data.nowMonth : null);
    const totals = monthlyTotalsByStaff(mk());
    const nextOf = staffId => { const tt = totals[String(staffId)]; return tt ? tt.nextNew + tt.nextRepeat : 0; };
    return {
        shop: shopId => {
            // 店舗別の実績は全店舗表示のサマリには無いため、所属スタッフ行の合計で代用
            const staffs = state.masters.staffs.filter(s => String(s.shop_id) === String(shopId));
            const agg = { sales: 0, newVisits: 0, nextBookings: staffs.reduce((a, s) => a + nextOf(s.id), 0) };
            for (const r of summary?.by_staff || []) {
                if (!staffs.some(s => String(s.id) === String(r.staff_id))) continue;
                const k = kpisOf(r);
                agg.sales += k.sales;
                agg.newVisits += k.newVisits;
            }
            return agg;
        },
        staff: staffId => {
            const r = (summary?.by_staff || []).find(x => String(x.staff_id) === String(staffId));
            const k = r ? kpisOf(r) : { sales: 0, newVisits: 0 };
            return { sales: k.sales, newVisits: k.newVisits, nextBookings: nextOf(staffId) };
        },
    };
}

function render() {
    // オーナー・マネージャー・店長が編集可（店長はサーバー側で自店舗のみのマスタが返る）
    if (!isAdmin() && !isManager() && !isStoreLocked()) return;
    const editor = document.getElementById('goal-editor');
    if (!editor) return;
    setText('goal-month-label', `${monthLabel(state.filters.anchor)}の目標`);
    const notice = document.getElementById('goal-storage-notice');
    if (notice) {
        notice.textContent = goalsStorage() === 'local'
            ? '⚠️ サーバー保存が未設定のため、目標はこの端末（ブラウザ）にのみ保存されます。他の端末と共有する場合はエクスポート/インポートをご利用ください。'
            : '目標はサーバーに保存され、全端末とスタッフのホーム・マイ成績（目標リング・進捗）に反映されます。店舗目標を空欄にすると所属スタッフの合計が自動で使われます。';
    }

    const key = mk();
    const actuals = actualsFor();
    const actualText = (a, f) => {
        const v = a?.[f.key];
        return v ? `<p class="goal-actual">実績 ${f.fmt(v)}</p>` : '';
    };
    editor.innerHTML = state.masters.shops.map(shop => {
        const shopGoal = getGoalRaw(key, `shop:${shop.id}`) || {};
        const staffs = state.masters.staffs.filter(s => String(s.shop_id) === String(shop.id));
        const shopActual = actuals.shop(shop.id);
        return `
        <div class="border border-surface-200 dark:border-gray-700 rounded-xl p-4">
            <h4 class="font-display font-bold text-accent-900 mb-3">${esc(shop.name)}</h4>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-3 mb-4">
                ${FIELDS.map(f => `
                    <div>
                        <label class="block text-xs text-surface-500 mb-1" for="goal-shop-${shop.id}-${f.key}">店舗${f.label}</label>
                        <input type="number" inputmode="numeric" min="0" step="${f.step}" value="${shopGoal[f.key] || ''}"
                            id="goal-shop-${shop.id}-${f.key}" placeholder="${f.placeholder}"
                            data-scope="shop:${shop.id}" data-field="${f.key}"
                            class="w-full px-3 py-2 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                        ${actualText(shopActual, f)}
                    </div>`).join('')}
            </div>
            ${staffs.length ? `
            <div class="overflow-x-auto">
                <table class="w-full text-sm whitespace-nowrap m-cards m-cards-goal">
                    <thead>
                        <tr class="border-b border-surface-200 dark:border-gray-700 text-surface-500">
                            <th class="text-left py-2 px-2 font-semibold">スタッフ</th>
                            ${FIELDS.map(f => `<th class="text-right py-2 px-2 font-semibold">${f.label}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${staffs.map(st => {
                            const g = getGoalRaw(key, `staff:${st.id}`) || {};
                            const a = actuals.staff(st.id);
                            return `<tr class="border-b border-surface-100 dark:border-gray-800">
                                <td class="py-2 px-2 font-medium">${esc(st.name)}</td>
                                ${FIELDS.map(f => `
                                    <td class="py-1.5 px-2 text-right" data-label="${f.label.replace(/（.*）/, '')}">
                                        <input type="number" inputmode="numeric" min="0" step="${f.step}" value="${g[f.key] || ''}"
                                            data-scope="staff:${st.id}" data-field="${f.key}" aria-label="${esc(st.name)}の${f.label}"
                                            class="w-28 text-right px-2 py-1 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                                        ${actualText(a, f)}
                                    </td>`).join('')}
                            </tr>`;
                        }).join('')}
                    </tbody>
                </table>
            </div>` : ''}
        </div>`;
    }).join('') || '<p class="text-sm text-surface-500">店舗情報がありません</p>';
}

async function onEdit(ev) {
    const input = ev.target.closest('input[data-scope]');
    if (!input) return;
    const scope = input.dataset.scope;
    const key = mk();
    const current = { ...(getGoalRaw(key, scope) || {}) };
    const v = Number(input.value);
    if (isFinite(v) && v > 0) current[input.dataset.field] = Math.round(v);
    else delete current[input.dataset.field];
    input.disabled = true;
    try {
        const res = await setGoal(key, scope, current);
        toast(res.storage === 'local' ? '目標を保存しました（この端末のみ）' : '目標を保存しました', 'success');
        emit('data:core'); // リング・ペースを再描画
    } catch (e) {
        console.error('goal save', e);
        toast(e?.body?.detail || '目標の保存に失敗しました', 'error');
    } finally {
        input.disabled = false;
    }
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
        await importGoals(await file.text());
        toast('目標をインポートしました', 'success');
        render();
        emit('data:core');
    } catch (e) {
        toast('インポートに失敗しました: ' + (e?.body?.detail || e.message), 'error');
    } finally {
        ev.target.value = '';
    }
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
