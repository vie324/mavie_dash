// Chart.js 共通テーマ・ヘルパー（旧dashboard.jsから移植し、ダークモード連動を維持）

export const BrandColors = {
    primary: '#47566b',
    primaryLight: '#6e819c',
    primaryDark: '#3d4859',
    accent: '#b8956a',
    accentLight: '#d4b896',
    accentDark: '#a07d52',
    gold: '#c9a96e',
    brown: '#47566b',
    beige: '#dcc9b3',
    light: '#f6f4f1',
    white: '#ffffff',
    darkBrown: '#3d4859',
    sage: '#739977',
    rose: '#b08f8a',
    warmgold: '#d4b896',
    success: '#739977',
    warning: '#c9a96e',
    purple: '#b08f8a',
    danger: '#c0705e',
};

// カテゴリ配色（媒体・年代など系列が多いチャート用）
export const Palette = ['#b8956a', '#47566b', '#739977', '#b08f8a', '#c9a96e', '#6e819c', '#dcc9b3', '#8a9bb4'];

export const charts = {}; // id -> Chart インスタンス

export function chartTheme() {
    const isDark = document.documentElement.classList.contains('dark');
    return {
        isDark,
        text: isDark ? '#e8e6e1' : '#47566b',
        textMuted: isDark ? '#aea69a' : '#7a7167',
        grid: isDark ? 'rgba(232,230,225,0.08)' : 'rgba(71,86,107,0.08)',
        tooltipBg: isDark ? 'rgba(38,42,46,0.96)' : 'rgba(255,255,255,0.97)',
        tooltipText: isDark ? '#f6f4f1' : '#3d4859',
        tooltipBorder: isDark ? 'rgba(232,230,225,0.12)' : 'rgba(184,149,106,0.25)',
        donutBorder: isDark ? '#1e2024' : '#ffffff',
    };
}

export function makeVGradient(context, colorTop, colorBottom) {
    const { ctx, chartArea } = context.chart;
    if (!chartArea) return colorTop;
    const g = ctx.createLinearGradient(0, chartArea.top, 0, chartArea.bottom);
    g.addColorStop(0, colorTop);
    g.addColorStop(1, colorBottom);
    return g;
}

// ドーナツ中央合計表示プラグイン
const centerTextPlugin = {
    id: 'centerText',
    afterDraw(chart) {
        const opt = chart.options.plugins?.centerText;
        if (!opt || !opt.text) return;
        const { ctx, chartArea } = chart;
        if (!chartArea) return;
        const x = (chartArea.left + chartArea.right) / 2;
        const y = (chartArea.top + chartArea.bottom) / 2;
        const t = chartTheme();
        ctx.save();
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.font = '700 18px Inter, sans-serif';
        ctx.fillStyle = t.text;
        ctx.fillText(opt.text, x, opt.sub ? y - 8 : y);
        if (opt.sub) {
            ctx.font = '500 10px Inter, sans-serif';
            ctx.fillStyle = t.textMuted;
            ctx.fillText(opt.sub, x, y + 12);
        }
        ctx.restore();
    },
};

export function initChartDefaults() {
    if (!window.Chart) return;
    Chart.defaults.font.family = 'Inter, -apple-system, sans-serif';
    try { Chart.register(centerTextPlugin); } catch (_) { /* 二重登録 */ }
}

export function chartCommonOptions() {
    const t = chartTheme();
    return {
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 700, easing: 'easeOutQuart' },
        plugins: {
            legend: {
                position: 'bottom',
                labels: { usePointStyle: true, pointStyle: 'circle', boxWidth: 8, padding: 14, color: t.text, font: { family: 'Inter, sans-serif', size: 11 } },
            },
            tooltip: {
                backgroundColor: t.tooltipBg,
                titleColor: t.tooltipText,
                bodyColor: t.tooltipText,
                borderColor: t.tooltipBorder,
                borderWidth: 1,
                cornerRadius: 10,
                padding: 10,
                boxPadding: 4,
                usePointStyle: true,
                titleFont: { family: 'Inter, sans-serif', weight: '700' },
                bodyFont: { family: 'Inter, sans-serif' },
            },
        },
    };
}

// 構造が同じなら破棄せず滑らかにトランジション
export function applyChartData(chart, newData) {
    if (!chart) return;
    const cur = chart.data;
    const sameShape = Array.isArray(cur?.datasets)
        && cur.datasets.length === newData.datasets.length
        && cur.datasets.every((ds, i) => (ds.type || chart.config.type) === (newData.datasets[i].type || chart.config.type) && ds.label === newData.datasets[i].label);
    if (sameShape) {
        cur.labels = newData.labels;
        newData.datasets.forEach((ds, i) => Object.assign(cur.datasets[i], ds));
    } else {
        chart.data = newData;
    }
    chart.update();
}

// キャンバスIDに対してチャートを生成 or 取得（タブ再表示時の二重生成を防ぐ）
export function ensureChart(id, config) {
    if (!window.Chart) return null;
    const el = document.getElementById(id);
    if (!el) return null;
    if (charts[id]) return charts[id];
    charts[id] = new Chart(el, config);
    return charts[id];
}

// ダークモード切替時: 全チャートの軸・凡例色を作り直す
export function refreshChartsTheme() {
    const t = chartTheme();
    const common = chartCommonOptions();
    for (const chart of Object.values(charts)) {
        if (!chart) continue;
        chart.options.plugins.legend.labels.color = t.text;
        Object.assign(chart.options.plugins.tooltip, common.plugins.tooltip);
        for (const axis of Object.values(chart.options.scales || {})) {
            if (axis.ticks) axis.ticks.color = t.textMuted;
            if (axis.grid && axis.grid.color) axis.grid.color = t.grid;
            if (axis.title) axis.title.color = t.textMuted;
        }
        chart.update('none');
    }
}

// スパークライン（軽量インラインSVG）
export function sparklineSvg(values, { width = 120, height = 34, color = BrandColors.accent } = {}) {
    const vals = (values || []).filter(v => v !== null && v !== undefined);
    if (vals.length < 2) return '';
    const max = Math.max(...vals), min = Math.min(...vals);
    const range = max - min || 1;
    const pts = vals.map((v, i) => {
        const x = (i / (vals.length - 1)) * (width - 4) + 2;
        const y = height - 4 - ((v - min) / range) * (height - 10);
        return `${x.toFixed(1)},${y.toFixed(1)}`;
    });
    const first = pts[0].split(',');
    const last = pts[pts.length - 1].split(',');
    return `<svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" aria-hidden="true">
        <polygon points="${first[0]},${height} ${pts.join(' ')} ${last[0]},${height}" fill="${color}18"/>
        <polyline points="${pts.join(' ')}" fill="none" stroke="${color}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
    </svg>`;
}
