// カレンダータブ: 日別売上ヒートマップ（表示月 = ヘッダーの対象月）

import { state, on, emit } from '../core/state.js';
import { yen, yenShort, num, monthLabel, daysInMonth, ymd, todayStr } from '../core/format.js';
import { salesOf } from '../data/salonone.js';

export function init() {
    on('data:core', render);
    document.getElementById('cal-prev')?.addEventListener('click', () => shiftMonth(-1));
    document.getElementById('cal-next')?.addEventListener('click', () => shiftMonth(1));
}

function shiftMonth(diff) {
    let { y, m } = state.filters.anchor;
    m += diff;
    while (m < 1) { m += 12; y -= 1; }
    while (m > 12) { m -= 12; y += 1; }
    state.filters.anchor = { y, m };
    const sel = document.getElementById('date-selector');
    if (sel) {
        sel.value = `${y}-${m}`;
        sel.dispatchEvent(new Event('change'));
    } else {
        emit('filters');
    }
}

function render() {
    const grid = document.getElementById('calendar-grid');
    if (!grid || !state.data.summary) return;
    const { y, m } = state.filters.anchor;
    setText('calendar-month-label', monthLabel(state.filters.anchor));

    // 対象月のby_day（期間フィルタが複数月でも対象月分のみ切り出し）
    const prefix = `${y}-${String(m).padStart(2, '0')}`;
    const byDay = new Map((state.data.summary.by_day || []).filter(d => d.date.startsWith(prefix)).map(d => [d.date, d]));
    const dim = daysInMonth(y, m);
    const firstDow = new Date(`${prefix}-01T12:00:00Z`).getUTCDay();
    const today = todayStr();
    const max = Math.max(...[...byDay.values()].map(d => salesOf(d)), 1);

    const cells = [];
    for (let i = 0; i < firstDow; i++) cells.push('<div></div>');
    for (let day = 1; day <= dim; day++) {
        const date = ymd(y, m, day);
        const d = byDay.get(date);
        const sales = d ? salesOf(d) : 0;
        const visits = d ? (d.new_visit_count || 0) + (d.repeat_visit_count || 0) : 0;
        const future = date > today;
        const intensity = sales / max;
        const bg = future || sales === 0 ? '' :
            `background:rgba(184,149,106,${(0.12 + intensity * 0.55).toFixed(2)});`;
        cells.push(`
            <div class="cal-cell rounded-lg p-1.5 md:p-2 min-h-[64px] md:min-h-[76px] border border-surface-100 dark:border-gray-700/60 ${future ? 'opacity-40' : ''} ${date === today ? 'ring-2 ring-primary-400' : ''}" style="${bg}">
                <p class="text-[10px] md:text-xs font-semibold ${date === today ? 'text-primary-600' : 'text-surface-500'}">${day}</p>
                ${sales > 0 ? `<p class="text-[10px] md:text-xs font-bold text-accent-900 mt-0.5 tabular-nums">${yenShort(sales)}</p><p class="text-[9px] text-surface-500">${num(visits)}名</p>` : ''}
            </div>`);
    }
    grid.innerHTML = cells.join('');

    // 月間サマリ
    const summaryEl = document.getElementById('calendar-summary');
    if (summaryEl) {
        const days = [...byDay.values()].filter(d => d.date <= today);
        const total = days.reduce((a, d) => a + salesOf(d), 0);
        const visits = days.reduce((a, d) => a + (d.new_visit_count || 0) + (d.repeat_visit_count || 0), 0);
        const active = days.filter(d => salesOf(d) > 0);
        const best = active.length ? active.reduce((a, d) => salesOf(d) > salesOf(a) ? d : a) : null;
        const card = (label, value, sub) => `
            <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-4 text-center">
                <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">${label}</p>
                <p class="text-lg font-display font-bold text-accent-900">${value}</p>
                ${sub ? `<p class="text-[10px] text-surface-500 mt-0.5">${sub}</p>` : ''}
            </div>`;
        summaryEl.innerHTML = [
            card('月間売上', yen(total)),
            card('月間来店', `${num(visits)}名`),
            card('営業日数', `${active.length}日`),
            card('ベスト日', best ? `${Number(best.date.slice(8))}日` : '—', best ? yenShort(salesOf(best)) : ''),
        ].join('');
    }
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
