// マイダッシュボード（スタッフ個人ビュー）

import { state, on, currentStaffId, currentShopId, staffName, shopName, isStaffLocked } from '../core/state.js';
import { monthlyTotalsByStaff } from '../data/manual.js';
import { yen, yenShort, num, pct, esc, delta, todayJst, todayStr } from '../core/format.js';
import { kpisOf, salesOf } from '../data/salonone.js';
import { getGoal, monthKey } from '../data/goals.js';
import { renderRings } from '../core/engage.js';
import { ensureChart, applyChartData, chartCommonOptions, chartTheme, BrandColors } from '../core/charts.js';
import { aiGenerate } from '../core/api.js';

export function init() {
    on('data:core', render);
    on('data:marketing', render);
    on('data:manual', render);
    on('meta', updateAiVisibility);
    on('theme', render);
    document.getElementById('st-ai-btn')?.addEventListener('click', generateAdvice);
}

// マイダッシュボードは常に本人（スタッフロック時は「店舗全体」を選んでいても本人の数値）
function targetStaffId() {
    if (isStaffLocked()) return state.session.staffId;
    return currentStaffId();
}

function active() {
    return targetStaffId() !== 'all';
}

function updateAiVisibility() {
    document.getElementById('st-ai-section')?.classList.toggle('hidden', !state.aiAvailable);
}

function render() {
    if (!active() || !state.data.summary) return;
    const staffId = targetStaffId();
    const name = state.session?.staffName || staffName(staffId);
    const byStaff = state.data.summary.by_staff || [];
    const row = byStaff.find(r => String(r.staff_id) === String(staffId));
    const k = kpisOf(row || {});
    const prevRow = (state.data.summaryPrev?.by_staff || []).find(r => String(r.staff_id) === String(staffId));
    const pk = prevRow ? kpisOf(prevRow) : null;

    setText('st-name', name);
    setText('st-avatar', (name || '?').charAt(0));
    setText('st-shop', shopName(currentShopId()));

    const sorted = [...byStaff].sort((a, b) => salesOf(b) - salesOf(a));
    const rank = sorted.findIndex(r => String(r.staff_id) === String(staffId));
    setText('st-rank', rank >= 0 ? `${rank + 1}位 / ${sorted.length}人` : '—');

    setText('st-sales', yen(k.sales));
    const d = delta(k.sales, pk?.sales);
    const deltaEl = document.getElementById('st-sales-delta');
    if (deltaEl) {
        deltaEl.textContent = d.text === '—' ? '' : `前期間比 ${d.text}`;
        deltaEl.style.color = d.dir === 'up' ? '#5d7d60' : d.dir === 'down' ? '#b08f8a' : '';
    }
    setText('st-visits', `${num(k.visits)}名`);
    setText('st-visits-sub', `新規 ${num(k.newVisits)} / 再来 ${num(k.repeatVisits)}`);
    setText('st-unit', yen(k.unitPrice));
    setText('st-unit-sub', pk ? `前期間 ${yen(pk.unitPrice)}` : '');
    setText('st-new', `${num(k.newVisits)}名`);
    setText('st-new-sub', `キャンセル ${num(k.cancels)}件`);

    renderPersonalRings(staffId);
    renderRadar(staffId, byStaff);
    renderMkGrid(staffId);
}

function renderPersonalRings(staffId) {
    const t = todayJst();
    const goal = getGoal(monthKey({ y: t.y, m: t.m }), currentShopId(), staffId) || {};
    // 当月の個人実績
    const nowRow = (state.data.nowMonth?.by_staff || []).find(r => String(r.staff_id) === String(staffId));
    const k = kpisOf(nowRow || {});
    renderRings('st-rings', [
        {
            label: '売上', color: '#b8956a',
            pct: goal.sales > 0 ? k.sales / goal.sales * 100 : 0,
            value: yen(k.sales),
            sub: goal.sales > 0 ? `目標 ${yen(goal.sales)}` : '目標未設定',
        },
        {
            label: '新規来店', color: '#739977',
            pct: goal.newVisits > 0 ? k.newVisits / goal.newVisits * 100 : 0,
            value: `${num(k.newVisits)}名`,
            sub: goal.newVisits > 0 ? `目標 ${num(goal.newVisits)}名` : '目標未設定',
        },
    ]);
}

function renderRadar(staffId, byStaff) {
    const rows = byStaff.map(r => ({ id: r.staff_id, k: kpisOf(r) }));
    const mine = rows.find(r => String(r.id) === String(staffId));
    if (!mine) return;
    const maxOf = f => Math.max(...rows.map(f), 1);
    const t = chartTheme();
    const mkRow = (state.data.mkStaff || []).find(r => String(r.staff_id) === String(staffId));
    const mkMax = Math.max(...(state.data.mkStaff || []).filter(r => !r.is_total).map(r => r.purchase_in_period_count || 0), 1);

    const chart = ensureChart('staffRadarChart', {
        type: 'radar',
        data: { labels: [], datasets: [] },
        options: {
            ...chartCommonOptions(),
            scales: {
                r: {
                    min: 0, max: 100, ticks: { display: false },
                    grid: { color: t.grid }, angleLines: { color: t.grid },
                    pointLabels: { color: t.textMuted, font: { size: 10 } },
                },
            },
        },
    });
    applyChartData(chart, {
        labels: ['売上', '来店数', '客単価', '新規獲得', '入会獲得'],
        datasets: [{
            label: staffName(staffId),
            data: [
                mine.k.sales / maxOf(r => r.k.sales) * 100,
                mine.k.visits / maxOf(r => r.k.visits) * 100,
                mine.k.unitPrice / maxOf(r => r.k.unitPrice) * 100,
                mine.k.newVisits / maxOf(r => r.k.newVisits) * 100,
                (mkRow?.purchase_in_period_count || 0) / mkMax * 100,
            ],
            borderColor: BrandColors.accent,
            backgroundColor: BrandColors.accent + '33',
            pointRadius: 3,
        }],
    });
}

function renderMkGrid(staffId) {
    const grid = document.getElementById('st-mk-grid');
    if (!grid) return;
    const cell = (label, value, sub) => `
        <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-4 text-center">
            <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">${label}</p>
            <p class="text-xl font-display font-bold text-accent-900">${value}</p>
            ${sub ? `<p class="text-[10px] text-surface-500 mt-0.5">${sub}</p>` : ''}
        </div>`;

    const cells = [];
    const row = (state.data.mkStaff || []).find(r => String(r.staff_id) === String(staffId));
    if (row) {
        cells.push(
            cell('新規予約', num(row.new_booking_count)),
            cell('新規来店', num(row.new_visit_count), `キャンセル ${num(row.cancel_count)}`),
            cell('購入（期間内）', num(row.purchase_in_period_count), pct(row.purchase_in_period_rate)),
            cell('購入金額', yenShort(row.purchase_amount), `単価 ${yenShort(row.purchase_unit_price)}`),
        );
    }
    // 日報の次回予約（当月・手入力）
    const t = todayJst();
    const mt = monthlyTotalsByStaff(`${t.y}-${String(t.m).padStart(2, '0')}`)[String(staffId)];
    const nowRow = (state.data.nowMonth?.by_staff || []).find(r => String(r.staff_id) === String(staffId));
    if (mt) {
        const visits = nowRow ? (nowRow.new_visit_count || 0) + (nowRow.repeat_visit_count || 0) : 0;
        const nextTotal = (mt.nextNew || 0) + (mt.nextRepeat || 0);
        cells.push(cell('今月の次回予約率', visits > 0 ? pct(nextTotal / visits * 100, 0) : '—', `新規 ${num(mt.nextNew || 0)} / 既存 ${num(mt.nextRepeat || 0)}（来店 ${num(visits)}名）`));
    }
    // 実APIの拡張フィールド（稼働率・口コミ獲得数）があれば表示
    const sRow = (state.data.summary?.by_staff || []).find(r => String(r.staff_id) === String(staffId));
    if (sRow?.utilization_rate !== undefined && sRow?.utilization_rate !== null) {
        cells.push(cell('稼働率', pct(sRow.utilization_rate, 0), sRow.operating_minutes ? `施術 ${Math.round(sRow.operating_minutes / 60)}h` : ''));
    }
    if (sRow && (sRow.google_review_count !== undefined || sRow.hotpepper_review_count !== undefined)) {
        cells.push(cell('口コミ獲得', num((sRow.google_review_count || 0) + (sRow.hotpepper_review_count || 0)), `Google ${num(sRow.google_review_count || 0)} / HPB ${num(sRow.hotpepper_review_count || 0)}`));
    }
    grid.innerHTML = cells.join('') || '<p class="text-sm text-surface-500 col-span-full">この期間のマーケ集計データはありません</p>';
}

async function generateAdvice() {
    const btn = document.getElementById('st-ai-btn');
    const out = document.getElementById('st-ai-output');
    if (!btn || !out) return;
    const staffId = targetStaffId();
    const row = (state.data.summary?.by_staff || []).find(r => String(r.staff_id) === String(staffId));
    const k = kpisOf(row || {});
    const mkRow = (state.data.mkStaff || []).find(r => String(r.staff_id) === String(staffId));
    btn.disabled = true;
    out.innerHTML = '<p class="text-surface-500 text-sm animate-pulse">アドバイスを生成しています…</p>';
    try {
        const prompt = [
            'あなたはアイラッシュサロンの経験豊富なマネージャーです。以下のスタッフの実績を見て、日本語で具体的で前向きなアドバイスを3点、箇条書きで簡潔に書いてください。',
            `スタッフ名: ${staffName(staffId)}`,
            `期間売上: ${k.sales}円 / 来店数: ${k.visits}名（新規${k.newVisits}・再来${k.repeatVisits}）`,
            `客単価: ${k.unitPrice}円 / キャンセル率: ${k.cancelRate.toFixed(1)}%`,
            mkRow ? `新規予約${mkRow.new_booking_count}件・購入率${mkRow.purchase_in_period_rate}%・購入金額${mkRow.purchase_amount}円` : '',
        ].join('\n');
        const res = await aiGenerate(prompt);
        out.innerHTML = `<div class="ai-advice-body">${esc(res.text).replace(/\n/g, '<br>')}</div>`;
    } catch (e) {
        out.innerHTML = '<p class="text-rose-500 text-sm">アドバイスの生成に失敗しました。時間をおいて再度お試しください。</p>';
    } finally {
        btn.disabled = false;
    }
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
