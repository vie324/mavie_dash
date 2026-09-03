// サマリータブ + 共通ヘッダー部（スナップショット・KPIカード）

import { state, on, isAdmin, isAdminLike, currentShopId, currentStaffId, staffName } from '../core/state.js';
import { yen, yenShort, num, pct, esc, delta, applyDeltaBadge, countUp, todayStr, todayJst, monthLabel, shortDate, daysInMonth, ymd, dowJa } from '../core/format.js';
import { kpisOf, scopedRow, monthlyBuckets, monthToDate, forecastMonth, salesOf, salesBasis, SALES_BASES } from '../data/salonone.js';
import { getGoal, monthKey, scopeKey } from '../data/goals.js';
import { renderRings, greeting, maybeCelebrate } from '../core/engage.js';
import { ensureChart, applyChartData, chartCommonOptions, chartTheme, BrandColors, Palette, makeVGradient, sparklineSvg } from '../core/charts.js';
import { apiGetCached } from '../core/api.js';
import { monthlyTotalsByStaff } from '../data/manual.js';
import { staffsOfShop } from '../core/state.js';

const BLOG_TARGET = 10;

export function init() {
    on('data:core', renderCore);
    on('data:marketing', renderChannelShare);
    on('data:manual', renderBlogProgress);
    on('theme', () => { renderCore(); renderChannelShare(); });
}

function currentGoal() {
    const t = todayJst();
    return getGoal(monthKey({ y: t.y, m: t.m }), currentShopId(), currentStaffId());
}

function anchorGoal() {
    return getGoal(monthKey(state.filters.anchor), currentShopId(), currentStaffId());
}

function renderCore() {
    if (!state.data.summary) return;
    renderSnapshot();
    renderKpis();
    renderRingsSection();
    renderOverviewChart();
    renderRatio();
    renderHighlights();
    renderStoreRace();
    renderStaffSummary();
}

// ---- Today's Snapshot ----
function renderSnapshot() {
    const nowMonth = scopedRowMonth();
    const today = todayStr();
    const t = todayJst();
    const todayRow = (nowMonthByDay() || []).find(d => d.date === today);
    const tk = kpisOf(todayRow || {});

    const dateEl = document.getElementById('snapshot-date');
    if (dateEl) dateEl.textContent = `${t.y}年${t.m}月${t.d}日 (${dowJa(today)})`;

    setText('snapshot-sales', yen(tk.sales));
    // 本日サマリは店舗単位の集計（スタッフ別の日次データはAPIにないため）
    if (currentStaffId() !== 'all') {
        setText('snapshot-sales-sub', '店舗全体の実績');
    } else {
        // 前日比
        const yesterday = (nowMonthByDay() || []).filter(d => d.date < today).at(-1);
        const d = delta(tk.sales, yesterday ? salesOf(yesterday) : null);
        setText('snapshot-sales-sub', d.text === '—' ? '—' : `前日比 ${d.text}`);
    }
    setHtml('snapshot-customers', `${num(tk.visits)}<span class="text-sm font-sans font-normal ml-1 text-surface-500">名</span>`);
    setText('snapshot-customers-sub', `新規 ${num(tk.newVisits)} / 再来 ${num(tk.repeatVisits)}`);
    setHtml('snapshot-cancel', `${num(tk.cancels)}<span class="text-sm font-sans font-normal ml-1 text-surface-500">件</span>`);
    setText('snapshot-cancel-sub', `無断 ${num(tk.noShows)}件`);

    renderPaymentBreakdown(todayRow);

    // 月次ペース
    // スタッフ選択時は個人の当月実績（by_staff行）を目標と比較する。
    // 店舗全体のMTDを個人目標と比べると進捗が数倍に化けるため、必ずスコープを揃える。
    const staffScoped = currentStaffId() !== 'all';
    const storeMtd = monthToDate({ by_day: nowMonthByDay() });
    const mtdSales = staffScoped ? salesOf(nowMonth) : storeMtd.sales;
    const goal = currentGoal();
    const goalSales = goal?.sales || 0;
    const storeFc = forecastMonth({ by_day: nowMonthByDay() });
    // 個人の着地予測は店舗予測をシェアで按分した概算
    const fc = staffScoped
        ? (storeFc && storeMtd.sales > 0
            ? (r => ({ forecast: Math.round(storeFc.forecast * r), low: Math.round(storeFc.low * r), high: Math.round(storeFc.high * r) }))(mtdSales / storeMtd.sales)
            : null)
        : storeFc;
    const progress = goalSales > 0 ? mtdSales / goalSales * 100 : 0;
    const expected = (t.d / daysInMonth(t.y, t.m)) * 100;

    const fill = document.getElementById('pace-meter-fill');
    const marker = document.getElementById('pace-marker');
    if (fill) fill.style.width = `${Math.min(progress, 100)}%`;
    if (marker) marker.style.left = `${Math.min(expected, 100)}%`;
    setText('pace-expected', `${expected.toFixed(0)}%`);
    const status = document.getElementById('pace-status');
    if (status) {
        if (!goalSales) { status.textContent = '目標未設定'; }
        else {
            const onTrack = progress >= expected - 3;
            status.textContent = `${progress.toFixed(0)}% ${onTrack ? '順調です' : '追い上げましょう'}`;
            status.style.color = onTrack ? '#5d7d60' : '#b08f8a';
        }
    }
    const remaining = document.getElementById('pace-remaining');
    if (remaining) {
        if (goalSales > 0) {
            const left = goalSales - mtdSales;
            remaining.innerHTML = left > 0
                ? `<span class="remaining-chip">🎯 目標まであと <b>${yen(left)}</b></span>`
                : `<span class="remaining-chip done">🏆 月間目標達成！</span>`;
        } else remaining.innerHTML = '';
    }
    setText('pace-forecast', fc ? yen(fc.forecast) : '—');
    setText('pace-forecast-range', fc ? `${yenShort(fc.low)} 〜 ${yenShort(fc.high)}${staffScoped ? '（概算）' : ''}` : 'データが揃うと表示されます');

    const greetEl = document.getElementById('snapshot-greeting');
    if (greetEl) {
        const name = state.session?.staffName || (staffScoped ? staffName(currentStaffId()) : '');
        greetEl.textContent = greeting(name, goalSales > 0 ? mtdSales / goalSales / (expected / 100 || 1) : null);
    }

    maybeCelebrate(scopeKey(currentShopId(), currentStaffId()), monthKey({ y: t.y, m: t.m }), mtdSales, goalSales);
}

// 本日の支払い方法別内訳（実APIの by_day.payment_breakdown。媒体ごとに動的に色を割り当て）
const PAYMENT_COLORS = ['#739977', '#566882', '#c9a96e', '#b08f8a', '#6e819c', '#a07d52', '#8ba88e', '#94a3b8'];

function renderPaymentBreakdown(todayRow) {
    const row = document.getElementById('snapshot-payment-row');
    const bar = document.getElementById('snapshot-payment-bar');
    const legend = document.getElementById('snapshot-payment-legend');
    if (!row || !bar || !legend) return;
    const items = (todayRow?.payment_breakdown || [])
        .filter(p => p.is_sales !== false && (p.amount || 0) > 0)
        .sort((a, b) => b.amount - a.amount);
    const total = items.reduce((a, p) => a + p.amount, 0);
    if (items.length === 0 || total <= 0) { row.classList.add('hidden'); return; }
    row.classList.remove('hidden');
    setText('snapshot-payment-total', yen(total));
    bar.innerHTML = items.map((p, i) => `
        <div class="payment-bar-seg" style="width:${(p.amount / total * 100).toFixed(1)}%; background:${PAYMENT_COLORS[i % PAYMENT_COLORS.length]}" title="${esc(p.name || '')} ${yen(p.amount)}"></div>`).join('');
    legend.innerHTML = items.map((p, i) => `
        <span class="payment-legend-item">
            <span class="w-2.5 h-2.5 rounded-full inline-block" style="background:${PAYMENT_COLORS[i % PAYMENT_COLORS.length]}"></span>
            ${esc(p.name || '不明')} <b class="tabular-nums">${yen(p.amount)}</b>
        </span>`).join('');
}

// 当月サマリ（スタッフ選択時はby_staff行にフォールバック不可のため、日別は店舗全体を表示しつつ合計はスタッフ分）
function nowMonthByDay() {
    return state.data.nowMonth?.by_day || [];
}

function scopedRowMonth() {
    return scopedRow(state.data.nowMonth);
}

// ---- KPIカード ----
function renderKpis() {
    const cur = kpisOf(scopedRow(state.data.summary));
    const prev = kpisOf(scopedRow(state.data.summaryPrev));
    const yoy = kpisOf(scopedRow(state.data.summaryYoy));

    countUp(document.getElementById('kpi-sales'), cur.sales, { prefix: '¥' });
    applyDeltaBadge(document.getElementById('delta-sales'), delta(cur.sales, prev?.sales));
    const basis = SALES_BASES[salesBasis()];
    setText('kpi-sales-label', `総売上（${basis.short}）`);
    setText('kpi-sales-alt', salesBasis() === 'gross' ? `会計済み ${yen(cur.digest)}` : `粗売上 ${yen(cur.gross)}`);
    const goal = anchorGoal();
    setText('kpi-goal-ratio', goal?.sales ? pct(cur.sales / goal.sales * 100, 0) : '—');
    setYoy('yoy-sales', cur.sales, yoy?.sales);

    const custEl = document.getElementById('kpi-customers');
    if (custEl) {
        countUp(custEl, cur.visits, { suffix: '名' });
    }
    applyDeltaBadge(document.getElementById('delta-customers'), delta(cur.visits, prev?.visits));
    setText('kpi-new', num(cur.newVisits));
    setText('kpi-existing', num(cur.repeatVisits));
    setYoy('yoy-customers', cur.visits, yoy?.visits);

    countUp(document.getElementById('kpi-unit-price'), cur.unitPrice, { prefix: '¥' });
    applyDeltaBadge(document.getElementById('delta-unit-price'), delta(cur.unitPrice, prev?.unitPrice));

    countUp(document.getElementById('kpi-new-visits'), cur.newVisits, { suffix: '名' });
    applyDeltaBadge(document.getElementById('delta-new-visits'), delta(cur.newVisits, prev?.newVisits));
    setText('kpi-new-rate', pct(cur.newRate, 0));

    setText('kpi-cancel-rate', pct(cur.cancelRate));
    // キャンセル率は下がる方が良いので向きを反転
    const d = delta(cur.cancelRate, prev?.cancelRate);
    const inverted = { text: d.text, dir: d.dir === 'up' ? 'down' : d.dir === 'down' ? 'up' : 'flat' };
    applyDeltaBadge(document.getElementById('delta-cancel-rate'), inverted);
    setText('kpi-noshow', num(cur.noShows));

    renderSparklines();
}

function setYoy(id, cur, prevYear) {
    const el = document.getElementById(id);
    if (!el) return;
    if (!prevYear) { el.textContent = ''; return; }
    const d = delta(cur, prevYear);
    el.textContent = `前年比 ${d.text}`;
    el.style.color = d.dir === 'up' ? '#5d7d60' : d.dir === 'down' ? '#b08f8a' : '';
}

function renderSparklines() {
    const byDay = state.data.summary?.by_day || [];
    const multiMonth = state.filters.periodKind !== 'month';
    const series = multiMonth ? monthlyBuckets(byDay) : byDay.filter(d => d.date <= todayStr());
    const sales = series.map(d => salesOf(d));
    const visits = series.map(d => (d.new_visit_count || 0) + (d.repeat_visit_count || 0));
    const unit = series.map(d => {
        const v = (d.new_visit_count || 0) + (d.repeat_visit_count || 0);
        return v > 0 ? Math.round(salesOf(d) / v) : 0;
    });
    setHtml('spark-sales', sparklineSvg(sales, { color: BrandColors.accent }));
    setHtml('spark-customers', sparklineSvg(visits, { color: BrandColors.primary }));
    setHtml('spark-unit-price', sparklineSvg(unit, { color: BrandColors.sage }));
}

// ---- 目標リング（常に「今月」の実績 ÷ 今月の目標）----
function renderRingsSection() {
    const t = todayJst();
    const goal = currentGoal() || {};
    const staffScoped = currentStaffId() !== 'all';
    // スタッフ選択時は個人の当月by_staff行、それ以外は店舗by_dayのMTD
    const row = staffScoped ? scopedRowMonth() : monthToDate({ by_day: nowMonthByDay() });
    const rings = [
        {
            label: '売上', color: '#b8956a',
            pct: goal.sales > 0 ? salesOf(row) / goal.sales * 100 : 0,
            value: yen(salesOf(row)),
            sub: goal.sales > 0 ? `目標 ${yen(goal.sales)}` : '目標未設定',
        },
        {
            label: '新規来店', color: '#739977',
            pct: goal.newVisits > 0 ? (row.new_visit_count || 0) / goal.newVisits * 100 : 0,
            value: `${num(row.new_visit_count || 0)}名`,
            sub: goal.newVisits > 0 ? `目標 ${num(goal.newVisits)}名` : '目標未設定',
        },
    ];
    // 入会リングはマーケ集計が「今月」を指しているときだけ表示
    // （期間フィルタが別期間だと今月の目標と分子がずれるため）
    const anchorIsNow = state.filters.periodKind === 'month'
        && state.filters.anchor.y === t.y && state.filters.anchor.m === t.m;
    if (anchorIsNow && state.data.mkStaff) {
        const mkRow = staffScoped
            ? state.data.mkStaff.find(r => String(r.staff_id) === String(currentStaffId()))
            : state.data.mkStaff.find(r => r.is_total);
        if (mkRow) {
            rings.push({
                label: '入会（今月）', color: '#c9a96e',
                pct: goal.joins > 0 ? (mkRow.purchase_in_period_count || 0) / goal.joins * 100 : 0,
                value: `${num(mkRow.purchase_in_period_count || 0)}名`,
                sub: goal.joins > 0 ? `目標 ${num(goal.joins)}名` : '目標未設定',
            });
        }
    }
    renderRings('goal-rings', rings);
}

// ---- 推移チャート ----
function renderOverviewChart() {
    const t = chartTheme();
    const byDay = state.data.summary?.by_day || [];
    const multiMonth = state.filters.periodKind !== 'month';
    const series = multiMonth ? monthlyBuckets(byDay) : byDay;
    const labels = series.map(d => multiMonth ? d.month.replace('-', '/') : shortDate(d.date));
    const newData = series.map(d => d.new_visit_count || 0);
    const repeatData = series.map(d => d.repeat_visit_count || 0);
    const unitData = series.map(d => {
        const v = (d.new_visit_count || 0) + (d.repeat_visit_count || 0);
        return v > 0 ? Math.round(salesOf(d) / v) : null;
    });

    const chart = ensureChart('overviewChart', {
        type: 'bar',
        data: { labels: [], datasets: [] },
        options: {
            ...chartCommonOptions(),
            interaction: { mode: 'index', intersect: false },
            scales: {
                x: { stacked: true, grid: { display: false }, ticks: { color: t.textMuted, maxRotation: 0, autoSkip: true } },
                y: { type: 'linear', position: 'left', title: { display: true, text: '単価 (¥)', color: t.textMuted }, grid: { display: false }, ticks: { color: t.textMuted } },
                y1: { type: 'linear', position: 'right', stacked: true, title: { display: true, text: '来店数 (名)', color: t.textMuted }, grid: { color: t.grid }, ticks: { color: t.textMuted, precision: 0 } },
            },
        },
    });
    applyChartData(chart, {
        labels,
        datasets: [
            { type: 'line', label: '平均客単価', data: unitData, borderColor: BrandColors.brown, backgroundColor: BrandColors.brown, borderWidth: 2.5, tension: 0.35, yAxisID: 'y', pointRadius: 2, pointHoverRadius: 5, spanGaps: true },
            { type: 'bar', label: '新規', data: newData, backgroundColor: ctx => makeVGradient(ctx, '#d4b896', '#b8956a'), borderRadius: 6, yAxisID: 'y1', stack: 'v' },
            { type: 'bar', label: '再来', data: repeatData, backgroundColor: ctx => makeVGradient(ctx, '#6e819c', '#47566b'), borderRadius: 6, yAxisID: 'y1', stack: 'v' },
        ],
    });
}

function renderRatio() {
    const k = kpisOf(scopedRow(state.data.summary));
    const t = chartTheme();
    const chart = ensureChart('customerRatioChart', {
        type: 'doughnut',
        data: { labels: [], datasets: [] },
        options: { ...chartCommonOptions(), cutout: '68%' },
    });
    if (!chart) return;
    chart.options.plugins.centerText = { text: `${num(k.visits)}名`, sub: '総来店' };
    applyChartData(chart, {
        labels: ['新規', '再来'],
        datasets: [{
            data: [k.newVisits, k.repeatVisits],
            backgroundColor: [BrandColors.accent, BrandColors.primary],
            borderColor: t.donutBorder, borderWidth: 3, hoverOffset: 6,
        }],
    });
    setText('ratio-new-count', `${num(k.newVisits)}名`);
    setText('ratio-exist-count', `${num(k.repeatVisits)}名`);
    const newRate = k.visits > 0 ? k.newVisits / k.visits * 100 : 0;
    const el1 = document.getElementById('bar-new-rate');
    const el2 = document.getElementById('bar-exist-rate');
    if (el1) el1.style.width = `${newRate}%`;
    if (el2) el2.style.width = `${100 - newRate}%`;
}

function renderChannelShare() {
    const channels = state.data.channels;
    if (!channels) return;
    const t = chartTheme();
    const rows = channels.filter(c => (c.visit_count || 0) > 0).sort((a, b) => b.visit_count - a.visit_count);
    const totalVisits = rows.reduce((a, c) => a + c.visit_count, 0);
    const chart = ensureChart('channelShareChart', {
        type: 'doughnut',
        data: { labels: [], datasets: [] },
        options: { ...chartCommonOptions(), cutout: '55%' },
    });
    if (!chart) return;
    chart.options.plugins.centerText = { text: `${num(totalVisits)}名`, sub: '新規来店' };
    applyChartData(chart, {
        labels: rows.map(c => c.name || '未設定'),
        datasets: [{
            data: rows.map(c => c.visit_count),
            backgroundColor: rows.map((_, i) => Palette[i % Palette.length]),
            borderColor: t.donutBorder, borderWidth: 3, hoverOffset: 6,
        }],
    });
    renderRingsSection(); // 入会リングはマーケデータ到着後に完成する
}

// ---- ハイライト ----
function renderHighlights() {
    const grid = document.getElementById('highlights-grid');
    const section = document.getElementById('highlights-section');
    if (!grid || !section) return;
    const cur = state.data.summary?.by_staff || [];
    const prev = state.data.summaryPrev?.by_staff || [];
    const cards = [];

    if (cur.length > 0) {
        // MVP: 前期比の売上増加率トップ
        let mvp = null;
        for (const row of cur) {
            const p = prev.find(x => String(x.staff_id) === String(row.staff_id));
            if (!p || salesOf(p) < 10000) continue;
            const growth = (salesOf(row) - salesOf(p)) / salesOf(p) * 100;
            if (!mvp || growth > mvp.growth) mvp = { name: row.staff_name, growth };
        }
        if (mvp && mvp.growth > 3) {
            cards.push({ icon: '🏆', cls: 'hl-gold', title: `MVP: ${esc(mvp.name || '')}さん`, body: `売上が前期間比 +${mvp.growth.toFixed(0)}% と絶好調です` });
        }
        // トップセールス
        const top = [...cur].sort((a, b) => salesOf(b) - salesOf(a))[0];
        if (top && salesOf(top) > 0) {
            cards.push({ icon: '👑', cls: 'hl-primary', title: `売上トップ: ${esc(top.staff_name || '')}さん`, body: `${yen(salesOf(top))}（期間内）` });
        }
    }
    const total = kpisOf(scopedRow(state.data.summary));
    const prevTotal = kpisOf(scopedRow(state.data.summaryPrev));
    if (prevTotal && prevTotal.sales > 0) {
        const g = (total.sales - prevTotal.sales) / prevTotal.sales * 100;
        if (g <= -20) cards.push({ icon: '⚠️', cls: 'hl-rose', title: '売上が減少しています', body: `前期間比 ${g.toFixed(0)}%。要因を確認しましょう` });
        else if (g >= 15) cards.push({ icon: '📈', cls: 'hl-sage', title: '売上が好調です', body: `前期間比 +${g.toFixed(0)}%` });
    }
    if (total && total.cancelRate >= 15) {
        cards.push({ icon: '🚨', cls: 'hl-rose', title: 'キャンセル率が高めです', body: `${total.cancelRate.toFixed(1)}%（無断 ${total.noShows}件）。リマインド強化を検討しましょう` });
    }

    section.classList.toggle('hidden', cards.length === 0);
    grid.innerHTML = cards.map(c => `
        <div class="highlight-card ${c.cls}">
            <span class="hl-icon">${c.icon}</span>
            <div><p class="hl-title">${c.title}</p><p class="hl-body">${c.body}</p></div>
        </div>`).join('');
}

// ---- 店舗対抗レース + ベンチマーク（管理者・全店舗表示時のみ）----
const RACE_COLORS = ['#b8956a', '#739977', '#b08f8a', '#6e819c', '#c9a96e'];

async function renderStoreRace() {
    const section = document.getElementById('store-race-section');
    if (!section) return;
    const show = isAdminLike() && currentShopId() === 'all' && state.masters.shops.length > 1;
    section.classList.toggle('hidden', !show);
    if (!show) return;

    const t = todayJst();
    const from = ymd(t.y, t.m, 1);
    const to = ymd(t.y, t.m, daysInMonth(t.y, t.m));
    let perShop;
    try {
        perShop = await Promise.all(state.masters.shops.map(async shop => ({
            shop,
            summary: await apiGetCached('sales/summary', { from, to, shop_id: shop.id }, 300000),
        })));
    } catch (e) {
        console.warn('店舗別サマリ取得失敗', e);
        section.classList.add('hidden');
        return;
    }

    const mk = monthKey({ y: t.y, m: t.m });
    const rows = perShop.map(({ shop, summary }) => {
        const mtd = monthToDate(summary);
        const goal = getGoal(mk, shop.id, 'all');
        const k = kpisOf(summary);
        return {
            shop, k, mtd,
            goal: goal?.sales || 0,
            pct: goal?.sales > 0 ? mtd.sales / goal.sales * 100 : null,
        };
    }).sort((a, b) => (b.pct ?? -1) - (a.pct ?? -1));

    const bars = document.getElementById('store-race-bars');
    if (bars) {
        const medals = ['🥇', '🥈', '🥉'];
        bars.innerHTML = rows.map((r, i) => `
            <div class="race-row">
                <div class="race-head">
                    <span class="race-rank">${medals[i] || `${i + 1}位`}</span>
                    <span class="race-name">${esc(r.shop.name)}</span>
                    <span class="race-pct" style="color:${RACE_COLORS[i % RACE_COLORS.length]}">${r.pct === null ? '—' : r.pct.toFixed(0) + '%'}</span>
                    <span class="race-detail">${yen(r.mtd.sales)}${r.goal ? ` / ${yen(r.goal)}` : '（目標未設定）'}</span>
                </div>
                <div class="race-track">
                    <div class="race-goalline"></div>
                    ${r.pct === null ? '' : `<div class="race-fill" style="width:${Math.min(r.pct, 100)}%; background:linear-gradient(90deg, ${RACE_COLORS[i % RACE_COLORS.length]}88, ${RACE_COLORS[i % RACE_COLORS.length]})">
                        <span class="race-runner">🏃‍♀️</span>
                    </div>`}
                </div>
            </div>`).join('');
    }

    const body = document.getElementById('store-benchmark-body');
    if (body) {
        body.innerHTML = rows.map(r => `
            <tr class="border-b border-surface-100 dark:border-accent-800">
                <td class="py-2 px-3 font-semibold">${esc(r.shop.name)}</td>
                <td class="py-2 px-3 text-right">${yen(r.mtd.sales)}</td>
                <td class="py-2 px-3 text-right font-semibold">${r.pct === null ? '—' : pct(r.pct, 0)}</td>
                <td class="py-2 px-3 text-right">${yen(r.k.unitPrice)}</td>
                <td class="py-2 px-3 text-right">${num(r.k.newVisits)}名</td>
                <td class="py-2 px-3 text-right">${pct(r.k.cancelRate)}</td>
            </tr>`).join('');
    }

    // レーダーチャート（各指標を店舗間の最大値=100で正規化）
    const theme = chartTheme();
    const axes = ['売上', '来店数', '客単価', '新規', '低キャンセル'];
    const maxOf = f => Math.max(...rows.map(f), 1);
    const maxSales = maxOf(r => r.k.sales), maxVisits = maxOf(r => r.k.visits),
        maxUnit = maxOf(r => r.k.unitPrice), maxNew = maxOf(r => r.k.newVisits);
    const chart = ensureChart('storeRadarChart', {
        type: 'radar',
        data: { labels: axes, datasets: [] },
        options: {
            ...chartCommonOptions(),
            scales: {
                r: {
                    min: 0, max: 100, ticks: { display: false },
                    grid: { color: theme.grid }, angleLines: { color: theme.grid },
                    pointLabels: { color: theme.textMuted, font: { size: 10 } },
                },
            },
        },
    });
    applyChartData(chart, {
        labels: axes,
        datasets: rows.map((r, i) => ({
            label: r.shop.name,
            data: [
                r.k.sales / maxSales * 100,
                r.k.visits / maxVisits * 100,
                r.k.unitPrice / maxUnit * 100,
                r.k.newVisits / maxNew * 100,
                100 - Math.min(r.k.cancelRate * 4, 100),
            ],
            borderColor: RACE_COLORS[i % RACE_COLORS.length],
            backgroundColor: RACE_COLORS[i % RACE_COLORS.length] + '22',
            pointRadius: 2,
        })),
    });
}

// ---- ブログ・SNS更新進捗（手入力データ）----
function renderBlogProgress() {
    const section = document.getElementById('blog-progress-section');
    const container = document.getElementById('blog-progress-container');
    if (!section || !container) return;
    const t = todayJst();
    const month = `${t.y}-${String(t.m).padStart(2, '0')}`;
    const totals = monthlyTotalsByStaff(month);
    const shopId = currentShopId();
    const staffs = shopId === 'all' ? state.masters.staffs : staffsOfShop(shopId);
    if (staffs.length === 0) { section.classList.add('hidden'); return; }
    section.classList.remove('hidden');
    // 次回予約率の分母はSalonOneの当月スタッフ別来店数
    const byStaffNow = state.data.nowMonth?.by_staff || [];
    const rows = staffs.map(s => {
        const tt = totals[String(s.id)] || { blog: 0, sns: 0, reviews: 0, nextNew: 0, nextRepeat: 0 };
        const r = byStaffNow.find(x => String(x.staff_id) === String(s.id));
        const newV = r?.new_visit_count || 0, repV = r?.repeat_visit_count || 0;
        return { s, tt, newV, repV, visits: newV + repV };
    });
    const totalNext = rows.reduce((a, r) => a + r.tt.nextNew + r.tt.nextRepeat, 0);
    const totalVisits = rows.reduce((a, r) => a + r.visits, 0);
    const totalNewNext = rows.reduce((a, r) => a + r.tt.nextNew, 0);
    const totalNewVisits = rows.reduce((a, r) => a + r.newV, 0);
    const rateText = (n, d) => d > 0 ? (n / d * 100).toFixed(0) + '%' : '—';
    container.innerHTML = `
        <div class="grid grid-cols-3 gap-3 mb-4">
            <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-3 text-center">
                <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">総次回予約率</p>
                <p class="text-xl font-display font-bold text-primary-500">${rateText(totalNext, totalVisits)}</p>
                <p class="text-[10px] text-surface-500">${num(totalNext)} / ${num(totalVisits)}名</p>
            </div>
            <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-3 text-center">
                <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">新規次回予約率</p>
                <p class="text-xl font-display font-bold text-accent-900">${rateText(totalNewNext, totalNewVisits)}</p>
                <p class="text-[10px] text-surface-500">${num(totalNewNext)} / ${num(totalNewVisits)}名</p>
            </div>
            <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-3 text-center">
                <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">既存次回予約率</p>
                <p class="text-xl font-display font-bold text-accent-900">${rateText(totalNext - totalNewNext, totalVisits - totalNewVisits)}</p>
                <p class="text-[10px] text-surface-500">${num(totalNext - totalNewNext)} / ${num(totalVisits - totalNewVisits)}名</p>
            </div>
        </div>
        <div class="overflow-x-auto">
            <table class="w-full text-sm whitespace-nowrap">
                <thead>
                    <tr class="border-b border-surface-200 dark:border-accent-700 text-surface-500">
                        <th class="text-left py-2 px-3 font-semibold">スタッフ</th>
                        <th class="text-right py-2 px-3 font-semibold">次回予約率</th>
                        <th class="text-right py-2 px-3 font-semibold">新規</th>
                        <th class="text-right py-2 px-3 font-semibold">既存</th>
                        <th class="text-left py-2 px-3 font-semibold w-1/3">ブログ（目標${BLOG_TARGET}）</th>
                        <th class="text-right py-2 px-3 font-semibold">SNS</th>
                        <th class="text-right py-2 px-3 font-semibold">★5</th>
                    </tr>
                </thead>
                <tbody>
                ${rows.map(({ s, tt, newV, repV, visits }) => {
                    const pctVal = Math.min(tt.blog / BLOG_TARGET * 100, 100);
                    const done = tt.blog >= BLOG_TARGET;
                    return `
                    <tr class="border-b border-surface-100 dark:border-accent-800">
                        <td class="py-2 px-3 font-medium">${esc(s.name)}</td>
                        <td class="py-2 px-3 text-right tabular-nums font-semibold text-primary-600">${rateText(tt.nextNew + tt.nextRepeat, visits)}</td>
                        <td class="py-2 px-3 text-right tabular-nums">${num(tt.nextNew)}<span class="text-[10px] text-surface-400">/${num(newV)}</span></td>
                        <td class="py-2 px-3 text-right tabular-nums">${num(tt.nextRepeat)}<span class="text-[10px] text-surface-400">/${num(repV)}</span></td>
                        <td class="py-2 px-3">
                            <div class="flex items-center gap-2">
                                <div class="flex-1 bg-surface-100 dark:bg-gray-700 h-2 rounded-full overflow-hidden">
                                    <div class="h-2 rounded-full ${done ? 'bg-sage-500' : 'bg-primary-400'}" style="width:${pctVal}%"></div>
                                </div>
                                <span class="text-xs tabular-nums ${done ? 'text-sage-600 font-semibold' : 'text-surface-600'}">${tt.blog}${done ? ' ✅' : ''}</span>
                            </div>
                        </td>
                        <td class="py-2 px-3 text-right tabular-nums">${num(tt.sns)}</td>
                        <td class="py-2 px-3 text-right tabular-nums">${num(tt.reviews)}</td>
                    </tr>`;
                }).join('')}
                </tbody>
            </table>
        </div>`;
}

// ---- スタッフパフォーマンス ----
function renderStaffSummary() {
    const byStaff = state.data.summary?.by_staff || [];
    const prevByStaff = state.data.summaryPrev?.by_staff || [];
    const section = document.getElementById('staff-summary-section');
    if (!section) return;
    section.classList.toggle('hidden', byStaff.length === 0);
    setText('staff-summary-period', `${monthLabel(state.filters.anchor)}${state.filters.periodKind === 'month' ? '' : 'までの' + { '3months': '3ヶ月', '6months': '6ヶ月', year: '1年' }[state.filters.periodKind]}・売上順`);

    const rows = byStaff.map(r => ({ ...r, k: kpisOf(r) })).sort((a, b) => b.k.sales - a.k.sales);
    const limit = state.session?.role === 'staff' ? 3 : rows.length;

    const cards = document.getElementById('staff-cards-view');
    if (cards) {
        cards.innerHTML = rows.slice(0, Math.min(limit, 9)).map((r, i) => `
            <div class="premium-card p-4 flex items-center gap-3">
                <div class="w-10 h-10 rounded-full gradient-${['primary', 'accent', 'sage'][i % 3]} flex items-center justify-center text-white font-display">${esc((r.staff_name || '?').charAt(0))}</div>
                <div class="min-w-0 flex-1">
                    <p class="font-semibold text-sm text-accent-900 truncate">${esc(r.staff_name || '不明')}</p>
                    <p class="text-xs text-surface-500">${num(r.k.visits)}名来店 / 新規${num(r.k.newVisits)}</p>
                </div>
                <div class="text-right">
                    <p class="font-display font-bold text-accent-900">${yenShort(r.k.sales)}</p>
                    <p class="text-[10px] text-surface-500">単価 ${yenShort(r.k.unitPrice)}</p>
                </div>
            </div>`).join('');
    }

    renderRanking('sales-ranking-list', rows, r => r.k.sales, v => yen(v), prevByStaff, p => salesOf(p), limit);
    renderRanking('new-customers-ranking-list', [...rows].sort((a, b) => b.k.newVisits - a.k.newVisits), r => r.k.newVisits, v => `${num(v)}名`, prevByStaff, p => p.new_visit_count || 0, limit);
    renderRanking('unit-price-ranking-list', [...rows].sort((a, b) => b.k.unitPrice - a.k.unitPrice), r => r.k.unitPrice, v => yen(v), null, null, limit);
}

function renderRanking(containerId, sorted, valueOf, fmt, prevRows, prevValueOf, limit) {
    const el = document.getElementById(containerId);
    if (!el) return;
    const max = Math.max(...sorted.map(valueOf), 1);
    // 前期の順位（変動バッジ用）
    let prevOrder = null;
    if (prevRows && prevValueOf) {
        prevOrder = [...prevRows].sort((a, b) => prevValueOf(b) - prevValueOf(a)).map(r => String(r.staff_id));
    }
    el.innerHTML = sorted.slice(0, limit === 3 ? 3 : 5).map((r, i) => {
        let moveBadge = '';
        if (prevOrder) {
            const prevIdx = prevOrder.indexOf(String(r.staff_id));
            if (prevIdx === -1) moveBadge = '<span class="rank-move new">NEW</span>';
            else if (prevIdx > i) moveBadge = `<span class="rank-move up">↑${prevIdx - i}</span>`;
            else if (prevIdx < i) moveBadge = `<span class="rank-move down">↓${i - prevIdx}</span>`;
        }
        return `
        <div class="flex items-center gap-2 p-2 rounded-lg ${i === 0 ? 'bg-primary-50 dark:bg-primary-900/20' : ''}">
            <span class="w-6 text-center font-display font-bold ${i < 3 ? 'text-primary-500' : 'text-surface-400'}">${i + 1}</span>
            <div class="flex-1 min-w-0">
                <div class="flex items-center gap-1.5">
                    <p class="text-sm font-medium text-accent-900 truncate">${esc(r.staff_name || '不明')}</p>${moveBadge}
                </div>
                <div class="w-full bg-surface-100 dark:bg-gray-700 h-1 rounded-full mt-1">
                    <div class="rank-bar-fill bg-primary-400 h-1 rounded-full" style="width:${valueOf(r) / max * 100}%"></div>
                </div>
            </div>
            <span class="text-sm font-semibold tabular-nums">${fmt(valueOf(r))}</span>
        </div>`;
    }).join('');
}

// ---- util ----
function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
function setHtml(id, html) {
    const el = document.getElementById(id);
    if (el) el.innerHTML = html;
}
