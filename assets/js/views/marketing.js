// マーケティングタブ: 媒体別・担当者別の新規獲得、広告指標、媒体別の次回予約（日報）
// 次回予約の数字は日報（手入力）から集計する。分母の来店数はSalonOneのマーケ集計。

import { state, on, currentShopId, staffsOfShop } from '../core/state.js';
import { yen, yenShort, num, pct, esc } from '../core/format.js';
import { currentRange } from '../data/salonone.js';
import { getManual, nextBySource, monthlyTotalsByStaff, monthsBetween, OTHER_SOURCE } from '../data/manual.js';

export function init() {
    on('data:marketing', render);
    on('data:manual', render);
    on('masters', render);
}

function anchorMonthKey() {
    return `${state.filters.anchor.y}-${String(state.filters.anchor.m).padStart(2, '0')}`;
}

// 表示期間に含まれる月（日報の集計範囲）
function rangeMonths() {
    const r = currentRange();
    return monthsBetween(r.from, r.to);
}

// 日報を集計するスタッフ（店舗を選んでいればその店舗の所属スタッフ、全店舗なら全員）
function scopeStaffIds() {
    const shopId = currentShopId();
    return shopId === 'all' ? null : staffsOfShop(shopId).map(s => s.id);
}

// APIの広告費がない媒体は手入力の広告費（日報入力タブ）で補完する
function effectiveAdSpend(c, manualAdCosts) {
    if (c.ad_spend > 0) return { amount: c.ad_spend, manual: false };
    const m = manualAdCosts[String(c.visit_source_id)];
    if (m > 0) return { amount: m, manual: true };
    return { amount: null, manual: false };
}

function hasSource(c) {
    return c.visit_source_id !== null && c.visit_source_id !== undefined && c.visit_source_id !== '';
}

function render() {
    const channels = state.data.channels;
    const mkStaff = state.data.mkStaff;
    if (!channels) return;
    const manualAdCosts = (state.filters.periodKind === 'month' ? getManual(anchorMonthKey()).adCosts : {}) || {};
    const months = rangeMonths();
    const nb = nextBySource(months, scopeStaffIds());

    // サマリカード
    const total = { booking: 0, visit: 0, adSpend: 0, adVisit: 0, hasAd: false, sales: 0 };
    for (const c of channels) {
        total.booking += c.booking_count || 0;
        total.visit += c.visit_count || 0;
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
    setText('mk-next-new', num(nb.total.n));
    setText('mk-next-new-sub', total.visit > 0
        ? `次回予約率 ${pct(nb.total.n / total.visit * 100, 0)}（新規の来店 ${num(total.visit)}名）`
        : '新規の来店がありません');
    setText('mk-adspend', total.hasAd ? yen(total.adSpend) : '—');
    setText('mk-adspend-sub', total.hasAd && total.adVisit ? `CPA ${yen(total.adSpend / total.adVisit)}（広告媒体のみ）` : total.hasAd ? '広告媒体の来店なし' : '広告費データなし');

    renderChannelTable(channels, manualAdCosts, nb);
    if (mkStaff) renderStaffTable(mkStaff, months);
    renderNextBySource(channels, nb, months);
}

function renderChannelTable(channels, manualAdCosts, nb) {
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
        // 流入元が未設定の行は日報の「その他・不明」と意味が違うため次回予約を出さない
        const next = hasSource(c) ? (nb.bySource[String(c.visit_source_id)]?.n || 0) : null;
        const nextRate = next !== null && c.visit_count > 0 ? next / c.visit_count * 100 : null;
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(c.name || '未設定・その他')}${platformBadge(c.platform_type)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(c.booking_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(c.visit_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${(c.cancel_rate || 0) >= 20 ? 'text-rose-500 font-semibold' : ''}">${pct(c.cancel_rate)}</td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold text-primary-600">${next === null ? '—' : num(next)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${rateClass(nextRate)}">${nextRate === null ? '—' : pct(nextRate, 0)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yenShort(c.sales)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${ad.amount ? yen(ad.amount) + manualMark : '—'}</td>
            <td class="py-2 px-3 text-right tabular-nums">${cpa ? yen(cpa) : '—'}</td>
            <td class="py-2 px-3 text-right tabular-nums ${roas >= 300 ? 'text-sage-600 font-semibold' : ''}">${roas ? pct(roas, 0) : '—'}</td>
        </tr>`;
    }).join('');
}

// 次回予約率の色分け（日報サマリと同じ: 70%以上=緑、50%以上=黄、50%未満=赤）
function rateClass(rate) {
    if (rate === null || rate === undefined) return 'text-surface-400';
    return rate >= 70 ? 'text-sage-600 font-semibold' : rate >= 50 ? 'text-primary-600 font-semibold' : 'text-rose-500 font-semibold';
}

function platformBadge(type) {
    if (type === 'meta') return ' <span class="text-[9px] font-bold text-white bg-[#6e819c] px-1.5 py-0.5 rounded align-middle">Meta</span>';
    if (type === 'tiktok') return ' <span class="text-[9px] font-bold text-white bg-[#3d4859] px-1.5 py-0.5 rounded align-middle">TikTok</span>';
    return '';
}

function renderStaffTable(mkStaff, months) {
    const body = document.getElementById('mk-staff-table-body');
    if (!body) return;
    const totals = monthlyTotalsByStaff(months);
    const nextOf = r => totals[String(r.staff_id)]?.nextNew || 0;
    const totalRow = mkStaff.find(r => r.is_total);
    const rows = mkStaff.filter(r => !r.is_total)
        .sort((a, b) => (b.new_visit_count || 0) - (a.new_visit_count || 0) || nextOf(b) - nextOf(a));
    const tr = (r, isTotal) => {
        const next = isTotal ? rows.reduce((a, x) => a + nextOf(x), 0) : nextOf(r);
        const rate = (r.new_visit_count || 0) > 0 ? next / r.new_visit_count * 100 : null;
        return `
        <tr class="border-b border-surface-100 dark:border-accent-800 ${isTotal ? 'bg-surface-50 dark:bg-gray-800/50 font-semibold' : ''}">
            <td class="py-2 px-3 font-medium">${isTotal ? '全体' : esc(r.staff_name || '未割当')}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.new_booking_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.new_visit_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(r.cancel_count)}</td>
            <td class="py-2 px-3 text-right tabular-nums font-semibold text-primary-600">${num(next)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${rateClass(rate)}">${rate === null ? '—' : pct(rate, 0)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${yenShort(r.new_customer_sales_total)}</td>
        </tr>`;
    };
    body.innerHTML = rows.map(r => tr(r, false)).join('') + (totalRow ? tr(totalRow, true) : '');
}

// ---- 媒体別の次回予約（日報）: 新規の次回予約を媒体別に（合計が多い順）。
// 2回目以降は媒体を問わない入力なので合計だけを添える（媒体別の r は媒体別に入力していた頃の日報の値） ----
function sourceName(key, channels) {
    if (key === OTHER_SOURCE) return 'その他・不明';
    const master = state.masters.visitSources.find(s => String(s.id) === key);
    if (master?.name) return master.name;
    const ch = channels.find(c => String(c.visit_source_id) === key);
    return ch?.name || `媒体#${key}`;
}

function renderNextBySource(channels, nb, months) {
    const list = document.getElementById('mk-next-list');
    if (!list) return;
    const [y0, m0] = months[0].split('-').map(Number);
    const [y1, m1] = months[months.length - 1].split('-').map(Number);
    const label = months.length === 1 ? `${y0}年${m0}月` : `${y0}年${m0}月〜${y1 === y0 ? '' : `${y1}年`}${m1}月`;
    const repTotal = nb.total.r + nb.total.no;
    setText('mk-next-sub', `${label}の日報から集計。どの媒体から来た新規のお客様が次回予約につながっているか（合計の多い順）。2回目以降は媒体を問わず ${num(nb.total.r)}件${repTotal > 0 && nb.total.no > 0 ? `（${num(repTotal)}名中・${Math.round(nb.total.r / repTotal * 100)}%）` : ''}`);

    const rows = Object.entries(nb.bySource).map(([key, c]) => {
        const ch = key === OTHER_SOURCE ? null : channels.find(x => String(x.visit_source_id) === key);
        return { key, name: sourceName(key, channels), n: c.n, r: c.r, total: c.n + c.r, newVisits: ch ? ch.visit_count || 0 : null };
    }).filter(r => r.total > 0)
        // 「その他・不明」は件数に関わらず最後に置く（ランキングの対象ではない）
        .sort((a, b) => (a.key === OTHER_SOURCE) - (b.key === OTHER_SOURCE) || b.total - a.total || b.n - a.n);
    const legacy = nb.noBreakdown.n;
    if (legacy > 0) rows.push({ key: 'legacy', name: '媒体の内訳なし（以前の日報）', n: nb.noBreakdown.n, r: 0, total: legacy, newVisits: null });

    document.getElementById('mk-next-legend')?.classList.toggle('hidden', rows.length === 0);
    if (rows.length === 0) {
        list.innerHTML = '<li class="next-rank-empty">この期間の日報に次回予約の入力はまだありません。スタッフの日報（媒体別）が入ると、ここに媒体ごとの次回予約が並びます。</li>';
        return;
    }
    const max = Math.max(...rows.map(r => r.total), 1);
    let rank = 0;
    list.innerHTML = rows.map(r => {
        const ranked = r.key !== OTHER_SOURCE && r.key !== 'legacy';
        if (ranked) rank++;
        const rate = r.newVisits > 0 ? r.n / r.newVisits * 100 : null;
        const share = nb.total.n + nb.total.r > 0 ? r.total / (nb.total.n + nb.total.r) * 100 : 0;
        const aria = `${r.name}: 次回予約 ${r.total}件（新規 ${r.n}件、2回目以降 ${r.r}件）`;
        return `
        <li class="next-rank-row${ranked ? '' : ' muted'}">
            <div class="next-rank-head">
                <span class="next-rank-no">${ranked ? rank : '–'}</span>
                <span class="next-rank-name" title="${esc(r.name)}">${esc(r.name)}</span>
                <span class="next-rank-total"><b>${num(r.total)}</b>件<small>${share.toFixed(0)}%</small></span>
            </div>
            <div class="next-rank-track" role="img" aria-label="${esc(aria)}">
                <div class="next-rank-bar" style="width:${(r.total / max * 100).toFixed(1)}%">
                    ${r.n > 0 ? `<span class="next-seg new" style="flex:${r.n} 1 0" title="新規の次回予約 ${r.n}件"></span>` : ''}
                    ${r.r > 0 ? `<span class="next-seg rep" style="flex:${r.r} 1 0" title="2回目以降の次回予約 ${r.r}件"></span>` : ''}
                </div>
            </div>
            <p class="next-rank-meta">
                <span class="next-meta-item"><i class="next-key new"></i>新規 <b>${num(r.n)}</b>${rate !== null ? `（次回予約率 <b class="next-rate ${rate >= 70 ? 'good' : rate >= 50 ? 'ok' : 'low'}">${pct(rate, 0)}</b> / 来店 ${num(r.newVisits)}名）` : ''}</span>
                ${r.r > 0 ? `<span class="next-meta-item"><i class="next-key rep"></i>2回目以降 <b>${num(r.r)}</b><small>（媒体別に入力していた頃）</small></span>` : ''}
            </p>
        </li>`;
    }).join('');
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
