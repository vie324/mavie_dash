// マーケティングタブ: 媒体別・担当者別の新規獲得、広告指標、リテンション

import { state, on } from '../core/state.js';
import { yen, yenShort, num, pct, esc } from '../core/format.js';
import { ensureChart, applyChartData, chartCommonOptions, chartTheme, BrandColors, makeVGradient } from '../core/charts.js';
import { getManual } from '../data/manual.js';

export function init() {
    on('data:marketing', render);
    on('data:retention', renderRetention);
    on('data:manual', render);
    on('theme', () => { render(); renderRetention(); });
}

function anchorMonthKey() {
    return `${state.filters.anchor.y}-${String(state.filters.anchor.m).padStart(2, '0')}`;
}

// APIの広告費がない媒体は手入力の広告費（日報入力タブ）で補完する
function effectiveAdSpend(c, manualAdCosts) {
    if (c.ad_spend > 0) return { amount: c.ad_spend, manual: false };
    const m = manualAdCosts[String(c.visit_source_id)];
    if (m > 0) return { amount: m, manual: true };
    return { amount: null, manual: false };
}

function render() {
    const channels = state.data.channels;
    const mkStaff = state.data.mkStaff;
    if (!channels) return;
    const manualAdCosts = (state.filters.periodKind === 'month' ? getManual(anchorMonthKey()).adCosts : {}) || {};

    // サマリカード
    const total = { booking: 0, visit: 0, joins: 0, adSpend: 0, adVisit: 0, hasAd: false, sales: 0 };
    for (const c of channels) {
        total.booking += c.booking_count || 0;
        total.visit += c.visit_count || 0;
        total.joins += c.join_in_period_count || 0;
        const ad = effectiveAdSpend(c, manualAdCosts);
        if (ad.amount > 0) {
            total.adSpend += ad.amount;
            total.adVisit += c.visit_count || 0; // CPAの分母は広告媒体の来店のみ
            total.hasAd = true;
        }
        total.sales += c.sales || 0;
    }
    setText('mk-bookings', num(total.booking));
    setText('mk-bookings-sub', `来店率 ${total.booking ? pct(total.visit / total.booking * 100, 0) : '—'}`);
    setText('mk-visits', num(total.visit));
    setText('mk-visits-sub', `新規売上(1〜3回) ${yenShort(total.sales)}`);
    setText('mk-joins', num(total.joins));
    setText('mk-joins-sub', `入会率 ${total.booking ? pct(total.joins / total.booking * 100) : '—'}（予約数ベース）`);
    setText('mk-adspend', total.hasAd ? yen(total.adSpend) : '—');
    setText('mk-adspend-sub', total.hasAd && total.adVisit ? `CPA ${yen(total.adSpend / total.adVisit)}（広告媒体のみ）` : total.hasAd ? '広告媒体の来店なし' : '広告費データなし');

    renderChannelTable(channels, manualAdCosts);
    if (mkStaff) renderStaffTable(mkStaff);
}

function renderChannelTable(channels, manualAdCosts) {
    const body = document.getElementById('channel-table-body');
    if (!body) return;
    const rows = [...channels].sort((a, b) => (b.booking_count || 0) - (a.booking_count || 0));
    if (rows.length === 0) {
        body.innerHTML = '<tr><td colspan="10" class="py-8 text-center text-surface-500">この期間のデータはありません</td></tr>';
        return;
    }
    body.innerHTML = rows.map(c => {
        const ad = effectiveAdSpend(c, manualAdCosts);
        const cpa = ad.manual ? (c.visit_count > 0 ? Math.round(ad.amount / c.visit_count) : null) : c.cpa;
        const roas = ad.manual ? (ad.amount > 0 ? Math.round((c.sales || 0) / ad.amount * 100) : null) : c.roas;
        const manualMark = ad.manual ? ' <span class="text-[9px] text-surface-400 align-middle">手入力</span>' : '';
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(c.name || '未設定・その他')}${platformBadge(c.platform_type)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(c.booking_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(c.visit_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${(c.cancel_rate || 0) >= 20 ? 'text-rose-500 font-semibold' : ''}">${pct(c.cancel_rate)}</td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold text-primary-600">${num(c.join_in_period_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${pct(c.join_in_period_rate_by_booking)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yenShort(c.sales)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${ad.amount ? yen(ad.amount) + manualMark : '—'}</td>
            <td class="py-2 px-3 text-right tabular-nums">${cpa ? yen(cpa) : '—'}</td>
            <td class="py-2 px-3 text-right tabular-nums ${roas >= 300 ? 'text-sage-600 font-semibold' : ''}">${roas ? pct(roas, 0) : '—'}</td>
        </tr>`;
    }).join('');
}

function platformBadge(type) {
    if (type === 'meta') return ' <span class="text-[9px] font-bold text-white bg-[#6e819c] px-1.5 py-0.5 rounded align-middle">Meta</span>';
    if (type === 'tiktok') return ' <span class="text-[9px] font-bold text-white bg-[#3d4859] px-1.5 py-0.5 rounded align-middle">TikTok</span>';
    return '';
}

function renderStaffTable(mkStaff) {
    const body = document.getElementById('mk-staff-table-body');
    if (!body) return;
    const totalRow = mkStaff.find(r => r.is_total);
    const rows = mkStaff.filter(r => !r.is_total).sort((a, b) => (b.purchase_in_period_amount || 0) - (a.purchase_in_period_amount || 0));
    const tr = (r, isTotal) => `
        <tr class="border-b border-surface-100 dark:border-accent-800 ${isTotal ? 'bg-surface-50 dark:bg-gray-800/50 font-semibold' : ''}">
            <td class="py-2 px-3 font-medium">${isTotal ? '全体' : esc(r.staff_name || '未割当')}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.new_booking_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.new_visit_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.cancel_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold text-primary-600">${num(r.purchase_in_period_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${pct(r.purchase_in_period_rate)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yenShort(r.purchase_in_period_amount)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yenShort(r.new_customer_sales_total)}</td>
        </tr>`;
    body.innerHTML = rows.map(r => tr(r, false)).join('') + (totalRow ? tr(totalRow, true) : '');
}

// ---- リテンション ----
// 仕様書§6はフィールド名を明記していないため、想定される別名にも耐性を持たせる
function pick(obj, ...keys) {
    for (const k of keys) {
        if (obj && obj[k] !== undefined && obj[k] !== null) return obj[k];
    }
    return null;
}

function normalizeRetSummary(raw) {
    return {
        join_count: pick(raw, 'join_count', 'joined_count', 'joins'),
        active_count: pick(raw, 'active_count', 'continue_count', 'retained_count', 'continuing_count'),
        churn_count: pick(raw, 'churn_count', 'churned_count', 'leave_count', 'left_count'),
        retention_rate: pick(raw, 'retention_rate', 'continuation_rate', 'continue_rate', 'retained_rate'),
        avg_purchase_count: pick(raw, 'avg_purchase_count', 'average_purchase_count', 'avg_purchases'),
    };
}

function renderRetention() {
    const ret = state.data.retention;
    if (!ret) return;
    const s = normalizeRetSummary(ret.summary || {});
    if (ret.summary && Object.values(s).every(v => v === null)) {
        console.warn('リテンションAPIのフィールド名が想定と異なります。実レスポンス:', Object.keys(ret.summary));
    }
    const grid = document.getElementById('ret-summary-grid');
    if (grid) {
        const card = (label, value, accent) => `
            <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-4 text-center">
                <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">${label}</p>
                <p class="text-xl font-display font-bold ${accent || 'text-accent-900'}">${value}</p>
            </div>`;
        grid.innerHTML = [
            card('入会数', num(s.join_count)),
            card('継続中', num(s.active_count), 'text-sage-600'),
            card('離反', num(s.churn_count), 'text-rose-500'),
            card('継続率', pct(s.retention_rate), 'text-primary-500'),
            card('平均購入回数', s.avg_purchase_count != null ? `${s.avg_purchase_count}回` : '—'),
        ].join('');
    }

    // 購入回数別の残存バー
    const byP = (ret.by_purchases || []).map(r => ({
        purchases: pick(r, 'purchases', 'purchase_count', 'times'),
        count: pick(r, 'count', 'customer_count', 'active_count') || 0,
    }));
    if (byP.length) {
        const t = chartTheme();
        const chart = ensureChart('retPurchasesChart', {
            type: 'bar',
            data: { labels: [], datasets: [] },
            options: {
                ...chartCommonOptions(),
                scales: {
                    x: { grid: { display: false }, ticks: { color: t.textMuted } },
                    y: { grid: { color: t.grid }, ticks: { color: t.textMuted, precision: 0 } },
                },
            },
        });
        applyChartData(chart, {
            labels: byP.map(r => `${r.purchases}回`),
            datasets: [{
                label: '人数',
                data: byP.map(r => r.count || 0),
                backgroundColor: ctx => makeVGradient(ctx, '#8fae92', '#739977'),
                borderRadius: 8,
            }],
        });
    }

    fillRetTable('ret-source-body', ret.by_source || [], r => esc(r.name || '未設定'));
    fillRetTable('ret-staff-body', ret.by_staff || [], r => esc(r.staff_name || '未割当'));
    // 顧客分析タブ側のカードも更新
    const cards = document.getElementById('cust-ret-cards');
    if (cards) {
        const mini = (label, value, cls) => `
            <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-3 text-center">
                <p class="text-[10px] text-surface-500 mb-0.5">${label}</p>
                <p class="text-lg font-display font-bold ${cls}">${value}</p>
            </div>`;
        cards.innerHTML = mini('入会', num(s.join_count), 'text-accent-900') + mini('継続', num(s.active_count), 'text-sage-600') + mini('継続率', pct(s.retention_rate), 'text-primary-500');
    }
    const custChartEl = document.getElementById('custRetChart');
    if (custChartEl && byP.length) {
        const t = chartTheme();
        const chart = ensureChart('custRetChart', {
            type: 'line',
            data: { labels: [], datasets: [] },
            options: {
                ...chartCommonOptions(),
                scales: {
                    x: { grid: { display: false }, ticks: { color: t.textMuted } },
                    y: { grid: { color: t.grid }, ticks: { color: t.textMuted, precision: 0 } },
                },
            },
        });
        applyChartData(chart, {
            labels: byP.map(r => `${r.purchases}回`),
            datasets: [{
                label: '残存人数',
                data: byP.map(r => r.count || 0),
                borderColor: BrandColors.sage,
                backgroundColor: BrandColors.sage + '30',
                fill: true, tension: 0.35, pointRadius: 3,
            }],
        });
    }
}

function fillRetTable(id, rows, nameOf) {
    const body = document.getElementById(id);
    if (!body) return;
    if (!rows.length) {
        body.innerHTML = '<tr><td colspan="5" class="py-6 text-center text-surface-500">データがありません</td></tr>';
        return;
    }
    const norm = rows.map(r => ({ ...normalizeRetSummary(r), _raw: r }));
    body.innerHTML = norm.sort((a, b) => (b.join_count || 0) - (a.join_count || 0)).map(r => `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${nameOf(r._raw)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.join_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums text-sage-600">${num(r.active_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums text-rose-500">${num(r.churn_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold">${pct(r.retention_rate)}</td>
        </tr>`).join('');
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
