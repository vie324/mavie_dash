// 売上詳細タブ: 3種の売上・曜日別分析・日別テーブル・CSVエクスポート

import { state, on, currentStaffId, shopName, staffName, currentShopId } from '../core/state.js';
import { yen, num, pct, shortDate, dowJa, dowIndex, todayStr } from '../core/format.js';
import { kpisOf, scopedRow, dowAnalysis, monthlyBuckets, downloadCsv } from '../data/salonone.js';
import { ensureChart, applyChartData, chartCommonOptions, chartTheme, BrandColors, makeVGradient } from '../core/charts.js';

export function init() {
    on('data:core', render);
    on('theme', render);
    document.getElementById('sales-csv-btn')?.addEventListener('click', exportCsv);
}

function render() {
    if (!state.data.summary) return;
    const row = scopedRow(state.data.summary);
    const k = kpisOf(row);
    setText('sale-gross', yen(k.gross));
    setText('sale-consumed', yen(k.consumed));
    setText('sale-digest', yen(k.digest));

    // 稼働率（実APIの拡張フィールド。スタッフ選択時はそのスタッフの値）
    const util = row.utilization_rate ?? row.period_utilization_rate ?? state.data.summary.period_utilization_rate;
    const card = document.getElementById('sale-utilization-card');
    if (card) {
        card.classList.toggle('hidden', util === null || util === undefined);
        if (util !== null && util !== undefined) {
            setText('sale-utilization', pct(util, 1));
            const op = row.operating_minutes ?? state.data.summary.period_operating_minutes;
            const av = row.available_minutes ?? state.data.summary.period_available_minutes;
            setText('sale-utilization-sub', op && av ? `施術 ${Math.round(op / 60)}h / 枠 ${Math.round(av / 60)}h` : '');
        }
    }

    renderDowChart();
    renderTable();
}

function renderDowChart() {
    const t = chartTheme();
    const rows = dowAnalysis(state.data.summary?.by_day);
    // 月曜始まりに並べ替え
    const order = [1, 2, 3, 4, 5, 6, 0];
    const labels = order.map(w => ['日', '月', '火', '水', '木', '金', '土'][w] + '曜');
    const chart = ensureChart('dowChart', {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            ...chartCommonOptions(),
            interaction: { mode: 'index', intersect: false },
            scales: {
                x: { grid: { display: false }, ticks: { color: t.textMuted } },
                y: { position: 'left', title: { display: true, text: '平均売上 (¥)', color: t.textMuted }, grid: { color: t.grid }, ticks: { color: t.textMuted } },
                y1: { position: 'right', title: { display: true, text: '平均来店 (名)', color: t.textMuted }, grid: { display: false }, ticks: { color: t.textMuted } },
            },
        },
    });
    applyChartData(chart, {
        labels,
        datasets: [
            { type: 'bar', label: '平均売上', data: order.map(w => rows[w].avgSales), backgroundColor: ctx => makeVGradient(ctx, '#d4b896', '#b8956a'), borderRadius: 8, yAxisID: 'y' },
            { type: 'line', label: '平均来店数', data: order.map(w => rows[w].avgVisits), borderColor: BrandColors.primary, backgroundColor: BrandColors.primary, borderWidth: 2.5, tension: 0.3, yAxisID: 'y1', pointRadius: 3 },
        ],
    });
}

function tableRows() {
    const multiMonth = state.filters.periodKind !== 'month';
    const byDay = state.data.summary?.by_day || [];
    if (multiMonth) {
        return monthlyBuckets(byDay).map(b => ({ label: b.month.replace('-', '年') + '月', ...b }));
    }
    const today = todayStr();
    return byDay.filter(d => d.date <= today).map(d => ({ label: `${shortDate(d.date)} (${dowJa(d.date)})`, isWeekend: [0, 6].includes(dowIndex(d.date)), ...d }));
}

function renderTable() {
    const body = document.getElementById('sales-table-body');
    if (!body) return;
    const rows = tableRows();
    setText('sales-table-title', state.filters.periodKind === 'month' ? '日別売上' : '月別売上');
    if (rows.length === 0) {
        body.innerHTML = '<tr><td colspan="8" class="py-8 text-center text-surface-500">この期間のデータはありません</td></tr>';
        return;
    }
    body.innerHTML = rows.map(r => {
        const visits = (r.new_visit_count || 0) + (r.repeat_visit_count || 0);
        const unit = visits > 0 ? Math.round(r.gross_sales / visits) : 0;
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800 ${r.isWeekend ? 'bg-surface-50/60 dark:bg-gray-800/40' : ''}">
            <td class="py-2 px-3 font-medium">${r.label}</td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold">${yen(r.gross_sales)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(r.consumed_sales)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(visits)}</td>
            <td class="py-2 px-3 text-right tabular-nums text-primary-600">${num(r.new_visit_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.repeat_visit_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${(r.cancel_count + (r.no_show_count || 0)) > 0 ? 'text-rose-500' : ''}">${num((r.cancel_count || 0) + (r.no_show_count || 0))}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yen(unit)}</td>
        </tr>`;
    }).join('');
}

function exportCsv() {
    const rows = tableRows();
    const scope = `${shopName(currentShopId())}${currentStaffId() !== 'all' ? '_' + staffName(currentStaffId()) : ''}`;
    downloadCsv(
        `vie_売上_${scope}_${state.filters.anchor.y}-${state.filters.anchor.m}.csv`,
        ['日付', '粗売上', '消化売上', '会計済み売上', '来店数', '新規', '再来', 'キャンセル', '無断キャンセル'],
        rows.map(r => [
            r.label,
            r.gross_sales || 0,
            r.consumed_sales || 0,
            r.digest_sales || 0,
            (r.new_visit_count || 0) + (r.repeat_visit_count || 0),
            r.new_visit_count || 0,
            r.repeat_visit_count || 0,
            r.cancel_count || 0,
            r.no_show_count || 0,
        ])
    );
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
