// 入金突合タブ: SalonOneの支払い方法別記録 × 実際額（レジ実査・端末集計）の照合
// オーナー・マネージャー・店長が利用（スタッフ不可）。店舗単位で照合する。

import { state, on, isAdminLike, isStoreLocked, currentShopId, shopName } from '../core/state.js';
import { yen, esc, todayStr, monthLabel } from '../core/format.js';
import { loadManual, saveManualPatch, getManual } from '../data/manual.js';
import { toast } from '../core/engage.js';

let selectedDate = todayStr();

export function init() {
    on('tab:shown', id => { if (id === 'recon') refresh(); });
    on('data:core', render);
    on('data:manual', render);
    document.getElementById('recon-date')?.addEventListener('change', ev => {
        selectedDate = ev.target.value || todayStr();
        render();
    });
    document.getElementById('recon-daily-body')?.addEventListener('change', ev => {
        const input = ev.target.closest('input[data-method-id]');
        if (input) saveActual(input);
    });
    document.getElementById('recon-memo')?.addEventListener('change', saveMemo);
    document.getElementById('recon-month-body')?.addEventListener('click', ev => {
        const row = ev.target.closest('tr[data-date]');
        if (row) {
            selectedDate = row.dataset.date;
            const dateInput = document.getElementById('recon-date');
            if (dateInput) dateInput.value = selectedDate;
            render();
        }
    });
}

function activeShopId() {
    const id = currentShopId();
    return id === 'all' ? null : id;
}

function month() {
    // 表示月はヘッダーの対象月に追従
    return `${state.filters.anchor.y}-${String(state.filters.anchor.m).padStart(2, '0')}`;
}

async function refresh() {
    // 選択日が対象月の外なら月初に合わせる
    if (!selectedDate.startsWith(month())) {
        const today = todayStr();
        selectedDate = today.startsWith(month()) ? today : `${month()}-01`;
    }
    const dateInput = document.getElementById('recon-date');
    if (dateInput) dateInput.value = selectedDate;
    await loadManual(month());
    render();
}

// 対象月の日別データ（summaryのby_dayから対象月分を切り出し）
function monthDays() {
    return (state.data.summary?.by_day || []).filter(d => d.date?.startsWith(month()));
}

function methodsOf(dayRow) {
    return (dayRow?.payment_breakdown || []).filter(p => p.is_sales !== false && (p.amount || 0) > 0);
}

function reconKey(date) {
    return `${date}:${activeShopId()}`;
}

function render() {
    if (!isAdminLike() && !isStoreLocked()) return;
    const guard = document.getElementById('recon-shop-guard');
    const bodyWrap = document.getElementById('recon-body');
    const shopId = activeShopId();
    if (guard && bodyWrap) {
        // 全店舗表示のままでは照合できない（レジは店舗単位のため）
        guard.classList.toggle('hidden', !!shopId);
        bodyWrap.classList.toggle('hidden', !shopId);
    }
    if (!shopId) return;

    setText('recon-shop-label', `${shopName(shopId)} / ${monthLabel(state.filters.anchor)}`);

    const recon = getManual(month()).recon || {};
    const dayRow = monthDays().find(d => d.date === selectedDate);
    const entry = recon[reconKey(selectedDate)] || {};

    // ---- 日次照合カード ----
    const body = document.getElementById('recon-daily-body');
    if (body) {
        const methods = methodsOf(dayRow);
        // 過去に入力済みだが当日の記録に無い方法も表示する
        const extraIds = Object.keys(entry).filter(k => /^m\d+$/.test(k) && !methods.some(p => `m${p.payment_method_id}` === k));
        if (methods.length === 0 && extraIds.length === 0) {
            body.innerHTML = `<tr><td colspan="4" class="py-8 text-center text-surface-500">${esc(selectedDate)} の売上記録はありません</td></tr>`;
        } else {
            let totalRec = 0, totalAct = 0, anyActual = false;
            const rows = methods.map(p => {
                const k = `m${p.payment_method_id}`;
                const actual = entry[k];
                const has = actual !== undefined && actual !== null;
                if (has) { totalAct += actual; anyActual = true; }
                totalRec += p.amount;
                const diff = has ? actual - p.amount : null;
                return `
                <tr class="border-b border-surface-100 dark:border-accent-800">
                    <td class="py-2 px-3 font-medium">${esc(p.name || '不明')}</td>
                    <td class="py-2 px-3 text-right tabular-nums">${yen(p.amount)}</td>
                    <td class="py-2 px-3 text-right">
                        <input type="number" inputmode="numeric" min="0" step="1" value="${has ? actual : ''}" placeholder="未入力"
                            data-method-id="${p.payment_method_id}"
                            class="w-32 text-right px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                    </td>
                    <td class="py-2 px-3 text-right tabular-nums font-semibold ${diffClass(diff)}">${diffLabel(diff)}</td>
                </tr>`;
            });
            for (const k of extraIds) {
                totalAct += entry[k]; anyActual = true;
                rows.push(`
                <tr class="border-b border-surface-100 dark:border-accent-800 bg-amber-50/50 dark:bg-amber-900/10">
                    <td class="py-2 px-3 font-medium text-surface-500">支払方法ID ${esc(k.slice(1))}（当日の記録なし）</td>
                    <td class="py-2 px-3 text-right tabular-nums">¥0</td>
                    <td class="py-2 px-3 text-right">
                        <input type="number" inputmode="numeric" min="0" step="1" value="${entry[k]}"
                            data-method-id="${k.slice(1)}"
                            class="w-32 text-right px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                    </td>
                    <td class="py-2 px-3 text-right tabular-nums font-semibold text-rose-500">+¥${entry[k].toLocaleString('ja-JP')}</td>
                </tr>`);
            }
            const totalDiff = anyActual ? totalAct - totalRec : null;
            rows.push(`
                <tr class="bg-surface-50 dark:bg-gray-800/50 font-semibold">
                    <td class="py-2 px-3">合計</td>
                    <td class="py-2 px-3 text-right tabular-nums">${yen(totalRec)}</td>
                    <td class="py-2 px-3 text-right tabular-nums">${anyActual ? yen(totalAct) : '—'}</td>
                    <td class="py-2 px-3 text-right tabular-nums ${diffClass(totalDiff)}">${diffLabel(totalDiff)}</td>
                </tr>`);
            body.innerHTML = rows.join('');
        }
    }
    const memoInput = document.getElementById('recon-memo');
    if (memoInput) memoInput.value = entry.memo || '';

    renderMonthTable(recon);
    renderMethodSummary(recon);
}

function dayState(dayRow, entry) {
    const methods = methodsOf(dayRow);
    if (methods.length === 0) return { label: '—', cls: 'text-surface-400', rec: 0, act: null };
    const rec = methods.reduce((a, p) => a + p.amount, 0);
    const entered = methods.filter(p => entry && entry[`m${p.payment_method_id}`] !== undefined);
    if (entered.length === 0) return { label: '未入力', cls: 'text-surface-400', rec, act: null };
    const act = methods.reduce((a, p) => a + (entry[`m${p.payment_method_id}`] ?? 0), 0)
        + Object.entries(entry).filter(([k]) => /^m\d+$/.test(k) && !methods.some(p => `m${p.payment_method_id}` === k)).reduce((a, [, v]) => a + v, 0);
    const diff = act - rec;
    if (entered.length < methods.length) return { label: '一部入力', cls: 'text-amber-600', rec, act, diff };
    return diff === 0
        ? { label: '✅ 一致', cls: 'text-sage-600 font-semibold', rec, act, diff }
        : { label: '⚠ 差異あり', cls: 'text-rose-500 font-semibold', rec, act, diff };
}

function renderMonthTable(recon) {
    const body = document.getElementById('recon-month-body');
    if (!body) return;
    const today = todayStr();
    const days = monthDays().filter(d => d.date <= today && ((d.gross_sales || 0) > 0 || recon[reconKey(d.date)]));
    if (days.length === 0) {
        body.innerHTML = '<tr><td colspan="5" class="py-6 text-center text-surface-500">対象月の営業データがありません</td></tr>';
        return;
    }
    body.innerHTML = days.map(d => {
        const st = dayState(d, recon[reconKey(d.date)]);
        return `
        <tr data-date="${d.date}" class="border-b border-surface-100 dark:border-accent-800 cursor-pointer hover:bg-surface-50 dark:hover:bg-gray-800/40 ${d.date === selectedDate ? 'bg-primary-50 dark:bg-primary-900/20' : ''}">
            <td class="py-2 px-3 font-medium">${Number(d.date.slice(8))}日</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(st.rec)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${st.act === null ? '—' : yen(st.act)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${diffClass(st.diff ?? null)}">${diffLabel(st.diff ?? null)}</td>
            <td class="py-2 px-3 text-right ${st.cls}">${st.label}</td>
        </tr>`;
    }).join('');
}

function renderMethodSummary(recon) {
    const body = document.getElementById('recon-summary-body');
    if (!body) return;
    const today = todayStr();
    const agg = new Map(); // methodId -> {name, rec, act, entered}
    for (const d of monthDays()) {
        if (d.date > today) continue;
        const entry = recon[reconKey(d.date)] || {};
        for (const p of methodsOf(d)) {
            const k = String(p.payment_method_id);
            if (!agg.has(k)) agg.set(k, { name: p.name || '不明', rec: 0, act: 0, entered: 0, days: 0 });
            const a = agg.get(k);
            a.rec += p.amount;
            a.days++;
            const actual = entry[`m${k}`];
            if (actual !== undefined && actual !== null) { a.act += actual; a.entered++; }
        }
    }
    if (agg.size === 0) {
        body.innerHTML = '<tr><td colspan="5" class="py-6 text-center text-surface-500">データがありません</td></tr>';
        return;
    }
    body.innerHTML = [...agg.values()].sort((a, b) => b.rec - a.rec).map(a => {
        const diff = a.entered > 0 ? a.act - a.rec : null;
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(a.name)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(a.rec)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${a.entered > 0 ? yen(a.act) : '—'}</td>
            <td class="py-2 px-3 text-right tabular-nums ${diffClass(diff)}">${diffLabel(diff)}</td>
            <td class="py-2 px-3 text-right text-surface-500 tabular-nums">${a.entered}/${a.days}日</td>
        </tr>`;
    }).join('');
}

async function saveActual(input) {
    const shopId = activeShopId();
    if (!shopId) return;
    const v = input.value === '' ? null : Number(input.value);
    if (v !== null && (!isFinite(v) || v < 0)) return;
    try {
        await saveManualPatch(month(), {
            recon: { [reconKey(selectedDate)]: { [`m${input.dataset.methodId}`]: v } },
        });
        toast('実際額を保存しました');
    } catch (e) {
        console.error('recon save', e);
        toast('保存に失敗しました', 'error');
    }
}

async function saveMemo(ev) {
    const shopId = activeShopId();
    if (!shopId) return;
    try {
        await saveManualPatch(month(), {
            recon: { [reconKey(selectedDate)]: { memo: ev.target.value || null } },
        });
    } catch (e) {
        console.error('recon memo save', e);
    }
}

function diffClass(diff) {
    if (diff === null || diff === undefined) return 'text-surface-400';
    return diff === 0 ? 'text-sage-600' : 'text-rose-500';
}

function diffLabel(diff) {
    if (diff === null || diff === undefined) return '—';
    if (diff === 0) return '±0';
    return (diff > 0 ? '+¥' : '−¥') + Math.abs(diff).toLocaleString('ja-JP');
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
