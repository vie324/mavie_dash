// 顧客分析タブ: 年代分布（個人情報なしの集計のみ使用）
// 継続分析のカード・チャートは marketing.js の renderRetention が描画する

import { state, on } from '../core/state.js';
import { num } from '../core/format.js';
import { ensureChart, applyChartData, chartCommonOptions, chartTheme, makeVGradient } from '../core/charts.js';

export function init() {
    on('data:agedist', render);
    on('theme', render);
}

function render() {
    const dist = state.data.ageDist;
    if (!dist) return;
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

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
