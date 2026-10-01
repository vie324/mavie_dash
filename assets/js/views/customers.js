// 顧客分析タブ: 年代分布・来店回数の分布（どちらも個人情報なしの集計のみ使用）

import { state, on } from '../core/state.js';
import { num } from '../core/format.js';
import { ensureChart, applyChartData, chartCommonOptions, chartTheme, makeVGradient } from '../core/charts.js';

const VISIT_LABELS = { '1': '1回', '2': '2回', '3': '3回', '4-5': '4〜5回', '6-9': '6〜9回', '10+': '10回以上' };

export function init() {
    on('data:agedist', render);
    on('theme', render);
}

function render() {
    const dist = state.data.ageDist;
    if (!dist) return;
    renderAges(dist);
    renderVisits(dist);
}

function renderAges(dist) {
    setText('age-total', num(dist.total) + (dist.truncated ? '+（一部のみ集計）' : ''));

    const brackets = Object.keys(dist.buckets || {}).map(Number).sort((a, b) => a - b);
    const labels = brackets.map(b => b < 10 ? '10歳未満' : `${b}代`);
    const data = brackets.map(b => dist.buckets[String(b)]);
    if (dist.unknown > 0) {
        labels.push('不明');
        data.push(dist.unknown);
    }

    const t = chartTheme();
    const chart = ensureChart('ageChart', {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            ...chartCommonOptions(),
            plugins: { ...chartCommonOptions().plugins, legend: { display: false } },
            scales: {
                x: { grid: { display: false }, ticks: { color: t.textMuted } },
                y: { grid: { color: t.grid }, ticks: { color: t.textMuted, precision: 0 } },
            },
        },
    });
    applyChartData(chart, {
        labels,
        datasets: [{
            label: '人数',
            data,
            backgroundColor: ctx => makeVGradient(ctx, '#d4b896', '#b8956a'),
            borderRadius: 8,
        }],
    });
}

// 来店回数の分布: 1回だけ（まだ2回目に来ていない）と2回以上（リピート）を色で分ける
// 色は日報の「新規 / 2回目以降」と同じ（CSS変数 --next-new / --next-rep。ダークモードは別の段）
function renderVisits(dist) {
    const cards = document.getElementById('cust-visit-cards');
    const v = dist.visits;
    if (!v || !v.total) {
        if (cards) cards.innerHTML = '<p class="col-span-full text-sm text-surface-500">SalonOneの顧客データに来店回数がないため表示できません</p>';
        return;
    }
    const b = v.buckets || {};
    const once = b['1'] || 0;
    const three = (b['3'] || 0) + (b['4-5'] || 0) + (b['6-9'] || 0) + (b['10+'] || 0);
    const share = n => `${Math.round(n / v.total * 100)}%`;
    if (cards) {
        const tile = (label, value, sub) => `
            <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-3 text-center">
                <p class="text-[10px] text-surface-500 mb-0.5">${label}</p>
                <p class="text-lg font-display font-bold text-accent-900">${value}</p>
                <p class="text-[10px] text-surface-500">${sub}</p>
            </div>`;
        cards.innerHTML = tile('1回だけ', share(once), `${num(once)}名`)
            + tile('2回以上（リピート）', share(v.total - once), `${num(v.total - once)}名`)
            + tile('3回以上', share(three), `${num(three)}名`);
    }
    const keys = Object.keys(VISIT_LABELS);
    const css = getComputedStyle(document.documentElement);
    const colorNew = css.getPropertyValue('--next-new').trim() || '#bf8a45';
    const colorRep = css.getPropertyValue('--next-rep').trim() || '#3f6aa3';
    const t = chartTheme();
    const chart = ensureChart('custVisitChart', {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            ...chartCommonOptions(),
            plugins: {
                ...chartCommonOptions().plugins,
                legend: { display: false },
                // チャートは使い回すため、割合は表示中のデータから毎回計算する
                tooltip: { ...(chartCommonOptions().plugins?.tooltip || {}), callbacks: { label: ctx => {
                    const sum = ctx.dataset.data.reduce((a, x) => a + (x || 0), 0);
                    return `${num(ctx.parsed.y)}名（${sum > 0 ? Math.round(ctx.parsed.y / sum * 100) : 0}%）`;
                } } },
            },
            scales: {
                x: { grid: { display: false }, ticks: { color: t.textMuted } },
                y: { grid: { color: t.grid }, ticks: { color: t.textMuted, precision: 0 } },
            },
        },
    });
    applyChartData(chart, {
        labels: keys.map(k => VISIT_LABELS[k]),
        datasets: [{
            label: '人数',
            data: keys.map(k => b[k] || 0),
            backgroundColor: keys.map(k => (k === '1' ? colorNew : colorRep)),
            borderRadius: 4,
            maxBarThickness: 44,
        }],
    });
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
