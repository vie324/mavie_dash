// 日報入力タブ: SalonOne APIにない項目の手入力
//   - 日次: 新規/既存の次回予約数・ブログ更新・SNS更新・★5口コミ（スタッフ本人 or 管理者/店長が入力）
//   - 月次: 物販売上（管理者/店長のみ・インセンティブ計算用のフォールバック）
//   - 広告費: 媒体別の手入力（オーナー/マネージャーのみ・APIに広告費がない媒体用）

import { state, on, isAdminLike, isStaffLocked, isStoreLocked, currentShopId, staffsOfShop, shopName } from '../core/state.js';
import { esc, num, yen, todayStr, ymd, daysInMonth, dowJa, dowIndex } from '../core/format.js';
import { apiGetCached } from '../core/api.js';
import {
    loadManual, saveManualPatch, getManual, monthlyTotalsByStaff, daysWithEntry, getDailyEntry,
    staffWithEntryOn, hasValues, emptyTotals, monthOf, BLOG_TARGET, DAILY_FIELDS,
} from '../data/manual.js';
import { toast } from '../core/engage.js';

const FIELD_IDS = { nextNew: 'input-next-new', nextRepeat: 'input-next-repeat', blog: 'input-blog', sns: 'input-sns', reviews: 'input-reviews' };

let currentDate = todayStr();
let dirty = false;
const monthSummaries = {}; // 'YYYY-MM' → sales/summary（次回予約率の分母用）

export function init() {
    on('tab:shown', id => { if (id === 'input') refresh(); });
    on('data:manual', () => { if (active()) render(); });
    on('data:core', () => { if (active()) { renderContextHint(); renderSummary(); } });
    on('masters', () => { renderStaffOptions(); });
    on('filters', () => {
        syncStaffFromHeader();
        if (active()) { renderStaffOptions(); render(); }
    });

    document.getElementById('input-date')?.addEventListener('change', ev => setDate(ev.target.value || todayStr()));
    document.getElementById('input-prev-day')?.addEventListener('click', () => shiftDay(-1));
    document.getElementById('input-next-day')?.addEventListener('click', () => shiftDay(1));
    document.getElementById('input-today')?.addEventListener('click', () => setDate(todayStr()));
    document.getElementById('input-staff')?.addEventListener('change', () => {
        if (!confirmDiscard()) { renderStaffOptions(); return; }
        fillDailyForm();
        renderDayStrip();
    });

    const fields = document.getElementById('input-fields');
    fields?.addEventListener('click', ev => {
        const btn = ev.target.closest('button[data-step]');
        if (!btn) return;
        const input = document.getElementById(btn.dataset.for);
        if (!input) return;
        const max = Number(input.max) || 999;
        const v = Math.min(max, Math.max(0, (Number(input.value) || 0) + Number(btn.dataset.step)));
        input.value = String(v);
        markDirty();
    });
    fields?.addEventListener('input', markDirty);
    fields?.addEventListener('keydown', ev => {
        if (ev.key === 'Enter') { ev.preventDefault(); saveDaily(); }
    });
    document.getElementById('input-day-strip')?.addEventListener('click', ev => {
        const cell = ev.target.closest('[data-date]');
        if (cell && !cell.classList.contains('future')) setDate(cell.dataset.date);
    });
    document.getElementById('input-daily-save')?.addEventListener('click', saveDaily);
    document.getElementById('input-daily-clear')?.addEventListener('click', clearDaily);

    document.getElementById('input-monthly-body')?.addEventListener('change', ev => {
        const input = ev.target.closest('input[data-staff-id]');
        if (input) saveMonthly(input);
    });
    document.getElementById('input-adcost-body')?.addEventListener('change', ev => {
        const input = ev.target.closest('input[data-source-id]');
        if (input) saveAdCost(input);
    });
    // 未保存のまま離れようとしたら確認
    window.addEventListener('beforeunload', ev => {
        if (dirty && active()) { ev.preventDefault(); ev.returnValue = ''; }
    });
}

function active() {
    return state.ui.activeTab === 'input';
}

// ---- 日付 ----
function clampDate(date) {
    const today = todayStr();
    return date > today ? today : date;
}

function confirmDiscard() {
    if (!dirty) return true;
    return window.confirm('保存していない変更があります。破棄して移動しますか？');
}

function setDate(date) {
    if (!confirmDiscard()) { const el = document.getElementById('input-date'); if (el) el.value = currentDate; return; }
    currentDate = clampDate(date);
    dirty = false;
    refresh();
}

function shiftDay(diff) {
    const d = new Date(currentDate + 'T12:00:00Z');
    d.setUTCDate(d.getUTCDate() + diff);
    setDate(ymd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()));
}

async function refresh() {
    currentDate = clampDate(currentDate);
    const dateInput = document.getElementById('input-date');
    if (dateInput) { dateInput.value = currentDate; dateInput.max = todayStr(); }
    renderStaffOptions();
    renderDateLabel();
    const month = monthOf(currentDate);
    await Promise.all([loadManual(month), ensureMonthSummary(month)]);
    render();
}

// ---- 対象スタッフ ----
function inputStaffs() {
    if (isStaffLocked()) return state.masters.staffs.filter(s => String(s.id) === String(state.session.staffId));
    const shopId = currentShopId();
    return shopId === 'all' ? state.masters.staffs : staffsOfShop(shopId);
}

function renderStaffOptions() {
    const sel = document.getElementById('input-staff');
    const wrap = document.getElementById('input-staff-wrap');
    const fixed = document.getElementById('input-staff-fixed');
    if (!sel) return;
    if (isStaffLocked()) {
        wrap?.classList.add('hidden');
        fixed?.classList.remove('hidden');
        const name = state.session.staffName || '';
        const nameEl = document.getElementById('input-staff-name');
        const avEl = document.getElementById('input-staff-avatar');
        if (nameEl) nameEl.textContent = name;
        if (avEl) avEl.textContent = name.charAt(0) || '–';
        sel.innerHTML = `<option value="${state.session.staffId}">${esc(name)}</option>`;
        sel.value = String(state.session.staffId);
        return;
    }
    wrap?.classList.remove('hidden');
    fixed?.classList.add('hidden');
    const staffs = inputStaffs();
    const prev = sel.value;
    if (currentShopId() === 'all' && state.masters.shops.length > 1) {
        // 全店舗表示のときは店舗ごとにグループ化
        sel.innerHTML = state.masters.shops.map(shop => {
            const list = staffs.filter(s => String(s.shop_id) === String(shop.id));
            if (!list.length) return '';
            return `<optgroup label="${esc(shop.name)}">${list.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('')}</optgroup>`;
        }).join('');
    } else {
        sel.innerHTML = staffs.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
    }
    if (staffs.some(s => String(s.id) === prev)) sel.value = prev;
    syncStaffFromHeader();
}

// ヘッダーでスタッフを選んでいる場合は入力対象も合わせる
function syncStaffFromHeader() {
    if (isStaffLocked()) return;
    const sel = document.getElementById('input-staff');
    if (!sel) return;
    const headerStaff = state.filters.staffId;
    if (headerStaff !== 'all' && inputStaffs().some(s => String(s.id) === String(headerStaff))) {
        if (sel.value !== String(headerStaff)) { sel.value = String(headerStaff); dirty = false; }
    }
}

function selectedStaffId() {
    if (isStaffLocked()) return String(state.session.staffId);
    return document.getElementById('input-staff')?.value || '';
}

function selectedStaffName() {
    const id = selectedStaffId();
    return state.masters.staffs.find(s => String(s.id) === id)?.name || '';
}

// ---- 来店数（次回予約率の分母）----
function monthRange(month) {
    const [y, m] = month.split('-').map(Number);
    return { from: ymd(y, m, 1), to: ymd(y, m, daysInMonth(y, m)) };
}

async function ensureMonthSummary(month) {
    const t = todayStr().slice(0, 7);
    if (month === t && state.data.nowMonth) { monthSummaries[month] = state.data.nowMonth; return; }
    const shopId = currentShopId();
    const key = `${month}:${shopId}`;
    if (monthSummaries[key]) return;
    try {
        const { from, to } = monthRange(month);
        monthSummaries[key] = await apiGetCached('sales/summary', { from, to, ...(shopId === 'all' ? {} : { shop_id: shopId }) }, 300000);
    } catch (e) {
        console.warn('month summary', e);
    }
}

function monthSummary(month) {
    const t = todayStr().slice(0, 7);
    if (month === t && state.data.nowMonth) return state.data.nowMonth;
    return monthSummaries[`${month}:${currentShopId()}`] || null;
}

function visitsOf(month, staffId) {
    const r = (monthSummary(month)?.by_staff || []).find(x => String(x.staff_id) === String(staffId));
    return r ? { visits: (r.new_visit_count || 0) + (r.repeat_visit_count || 0), newV: r.new_visit_count || 0, repV: r.repeat_visit_count || 0 } : { visits: 0, newV: 0, repV: 0 };
}

// ---- 描画 ----
function render() {
    const notice = document.getElementById('input-storage-notice');
    if (notice) notice.classList.toggle('hidden', state.manualStorage !== 'local');
    renderDateLabel();
    if (!dirty) fillDailyForm();
    else renderSaveState();
    renderDayStrip();
    renderContextHint();
    renderSummary();
    renderMonthlySection();
    renderAdCostSection();
}

function renderDateLabel() {
    const el = document.getElementById('input-date-label');
    if (!el) return;
    const [y, m, d] = currentDate.split('-').map(Number);
    const isToday = currentDate === todayStr();
    el.innerHTML = `${y}年${m}月${d}日（${dowJa(currentDate)}）${isToday ? '<span class="ml-2 text-[10px] font-bold text-primary-600 bg-primary-50 dark:bg-primary-900/30 px-2 py-0.5 rounded-full">今日</span>' : ''}`;
}

function fillDailyForm() {
    const entry = getDailyEntry(currentDate, selectedStaffId());
    for (const [f, id] of Object.entries(FIELD_IDS)) setValue(id, entry?.[f]);
    dirty = false;
    renderSaveState(entry);
}

function markDirty() {
    dirty = true;
    renderSaveState();
}

function renderSaveState(entry) {
    const el = document.getElementById('input-save-state');
    if (!el) return;
    if (entry === undefined) entry = getDailyEntry(currentDate, selectedStaffId());
    let cls = 'empty', text = '未入力';
    if (dirty) { cls = 'dirty'; text = '未保存の変更があります'; }
    else if (hasValues(entry)) {
        cls = 'saved';
        text = entry.at ? `✓ 保存済み ${fmtTime(entry.at)}` : '✓ 保存済み';
    }
    el.className = `save-state ${cls}`;
    el.textContent = text;
}

function fmtTime(sec) {
    const d = new Date(sec * 1000);
    const now = new Date();
    const time = d.toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Tokyo' });
    const sameDay = d.toDateString() === now.toDateString();
    return sameDay ? time : `${d.toLocaleDateString('ja-JP', { month: 'numeric', day: 'numeric', timeZone: 'Asia/Tokyo' })} ${time}`;
}

// この月の入力状況ストリップ
function renderDayStrip() {
    const el = document.getElementById('input-day-strip');
    if (!el) return;
    const month = monthOf(currentDate);
    const [y, m] = month.split('-').map(Number);
    const dim = daysInMonth(y, m);
    const today = todayStr();
    const has = daysWithEntry(month, selectedStaffId());
    const cells = [];
    let elapsed = 0;
    for (let d = 1; d <= dim; d++) {
        const date = ymd(y, m, d);
        const future = date > today;
        if (!future) elapsed++;
        const cls = ['day-dot'];
        if (has.has(date)) cls.push('has');
        if (date === currentDate) cls.push('selected');
        if (date === today) cls.push('today');
        if (future) cls.push('future');
        if ([0, 6].includes(dowIndex(date))) cls.push('weekend');
        cells.push(`<div class="${cls.join(' ')}" data-date="${date}" role="listitem" ${future ? 'aria-disabled="true"' : 'tabindex="0"'} title="${m}/${d}${has.has(date) ? ' 入力済み' : ''}"><span>${d}</span><small>${dowJa(date)}</small></div>`);
    }
    el.innerHTML = cells.join('');
    setText('input-strip-title', `${y}年${m}月の入力状況${selectedStaffName() ? `（${selectedStaffName()}）` : ''}`);
    setText('input-strip-summary', `入力済み ${has.size} / ${elapsed}日`);
}

// SalonOneの店舗来店数（入力の目安）
function renderContextHint() {
    const el = document.getElementById('input-context-hint');
    if (!el) return;
    const month = monthOf(currentDate);
    const row = (monthSummary(month)?.by_day || []).find(d => d.date === currentDate);
    if (!row) { el.textContent = ''; return; }
    const visits = (row.new_visit_count || 0) + (row.repeat_visit_count || 0);
    const [, m, d] = currentDate.split('-').map(Number);
    el.textContent = `${shopName(currentShopId())} ${m}/${d} の来店 ${num(visits)}名（新規 ${num(row.new_visit_count || 0)} / 再来 ${num(row.repeat_visit_count || 0)}・SalonOne）`;
}

// 月間合計テーブル
function renderSummary() {
    const month = monthOf(currentDate);
    const [y, m] = month.split('-').map(Number);
    const totals = monthlyTotalsByStaff(month);
    const body = document.getElementById('input-summary-body');
    if (!body) return;
    setText('input-summary-month', `${y}年${m}月の入力合計`);
    const today = todayStr();
    const elapsed = month === today.slice(0, 7) ? Number(today.slice(8)) : (month < today.slice(0, 7) ? daysInMonth(y, m) : 0);
    const staffs = inputStaffs();
    const grouped = currentShopId() === 'all' && state.masters.shops.length > 1 && !isStaffLocked();
    const rowHtml = s => {
        const t = totals[String(s.id)] || emptyTotals();
        const done = t.blog >= BLOG_TARGET;
        const v = visitsOf(month, s.id);
        const rate = v.visits > 0 ? (t.nextNew + t.nextRepeat) / v.visits * 100 : null;
        const rateCls = rate === null ? 'text-surface-400' : rate >= 70 ? 'text-sage-600 font-semibold' : rate >= 50 ? 'text-primary-600 font-semibold' : 'text-rose-500 font-semibold';
        const sel = String(s.id) === selectedStaffId();
        return `<tr class="border-b border-surface-100 dark:border-accent-800 ${sel ? 'bg-primary-50/60 dark:bg-primary-900/10' : ''}">
            <td class="py-2 px-3 font-medium">${esc(s.name)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${t.days === 0 && elapsed > 0 ? 'text-rose-500' : ''}">${num(t.days)}<span class="text-[10px] text-surface-400"> /${num(elapsed)}日</span></td>
            <td class="py-2 px-3 text-right tabular-nums text-primary-600 font-semibold">${num(t.nextNew)}<span class="text-[10px] text-surface-400"> /${num(v.newV)}</span></td>
            <td class="py-2 px-3 text-right tabular-nums">${num(t.nextRepeat)}<span class="text-[10px] text-surface-400"> /${num(v.repV)}</span></td>
            <td class="py-2 px-3 text-right tabular-nums ${rateCls}">${rate === null ? '—' : rate.toFixed(0) + '%'}<span class="text-[10px] text-surface-400 font-normal"> /${num(v.visits)}名</span></td>
            <td class="py-2 px-3 text-right tabular-nums ${done ? 'text-sage-600 font-semibold' : ''}">${num(t.blog)} / ${BLOG_TARGET}${done ? ' ✅' : ''}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(t.sns)}</td>
            <td class="py-2 px-3 text-right tabular-nums">${num(t.reviews)}</td>
        </tr>`;
    };
    let html = '';
    if (grouped) {
        for (const shop of state.masters.shops) {
            const list = staffs.filter(s => String(s.shop_id) === String(shop.id));
            if (!list.length) continue;
            html += `<tr class="bg-surface-50 dark:bg-gray-800/40"><td colspan="8" class="py-1.5 px-3 text-[11px] font-bold text-surface-500 uppercase tracking-wider">${esc(shop.name)}</td></tr>` + list.map(rowHtml).join('');
        }
    } else {
        html = staffs.map(rowHtml).join('');
    }
    body.innerHTML = html || '<tr><td colspan="8" class="py-6 text-center text-surface-500">スタッフがいません</td></tr>';

    // 本日の入力状況（管理者・店長向け）
    const statusEl = document.getElementById('input-today-status');
    if (statusEl) {
        if (isStaffLocked() || staffs.length === 0) {
            statusEl.classList.add('hidden');
        } else {
            const doneSet = staffWithEntryOn(today);
            const missing = staffs.filter(s => !doneSet.has(String(s.id)));
            const doneCount = staffs.length - missing.length;
            statusEl.classList.remove('hidden');
            const [, tm, td] = today.split('-').map(Number);
            statusEl.innerHTML = `<span class="font-semibold text-accent-800 dark:text-gray-100">本日（${tm}/${td}）の日報:</span>
                <span class="ml-1 ${doneCount === staffs.length ? 'text-sage-600 font-semibold' : 'text-surface-600'}">${doneCount} / ${staffs.length}名 入力済み</span>
                ${missing.length ? `<span class="ml-2 text-surface-500">未入力:</span> ${missing.map(s => `<button type="button" class="chip chip-rose" data-pick-staff="${s.id}">${esc(s.name)}</button>`).join(' ')}` : '<span class="ml-2 text-sage-600">全員入力済み ✅</span>'}`;
            statusEl.querySelectorAll('[data-pick-staff]').forEach(btn => btn.addEventListener('click', () => {
                if (!confirmDiscard()) return;
                const sel = document.getElementById('input-staff');
                if (sel) sel.value = btn.dataset.pickStaff;
                currentDate = today;
                dirty = false;
                refresh();
                document.getElementById('input-fields')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }));
        }
    }
}

// 月次: 物販売上（管理者・店長）
function renderMonthlySection() {
    const section = document.getElementById('input-monthly-section');
    if (!section) return;
    const show = isAdminLike() || isStoreLocked();
    section.classList.toggle('hidden', !show);
    if (!show) return;
    const month = monthOf(currentDate);
    const [y, m] = month.split('-').map(Number);
    setText('input-monthly-month', `${y}年${m}月`);
    const data = getManual(month);
    const body = document.getElementById('input-monthly-body');
    const byStaff = monthSummary(month)?.by_staff || [];
    body.innerHTML = inputStaffs().map(s => {
        const cur = data.monthly[String(s.id)]?.productSales;
        const api = byStaff.find(x => String(x.staff_id) === String(s.id))?.product_sales;
        const hasApi = api !== undefined && api !== null && api > 0;
        return `<tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(s.name)}</td>
            <td class="py-2 px-3 text-right tabular-nums ${hasApi ? 'text-sage-600' : 'text-surface-400'}">${hasApi ? yen(api) : '—'}</td>
            <td class="py-2 px-3 text-right">
                <input type="number" inputmode="numeric" min="0" step="1000" value="${cur ?? ''}" placeholder="${hasApi ? '自動取得済み' : '未入力'}"
                    data-staff-id="${s.id}" aria-label="${esc(s.name)}の物販売上"
                    class="w-32 text-right px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
            </td>
        </tr>`;
    }).join('');
}

// 広告費（オーナー・マネージャー）
function renderAdCostSection() {
    const section = document.getElementById('input-adcost-section');
    if (!section) return;
    section.classList.toggle('hidden', !isAdminLike());
    if (!isAdminLike()) return;
    const month = monthOf(currentDate);
    const [y, m] = month.split('-').map(Number);
    setText('input-adcost-month', `${y}年${m}月`);
    const data = getManual(month);
    const body = document.getElementById('input-adcost-body');
    const sources = [...state.masters.visitSources, { id: 'other', name: 'その他' }];
    body.innerHTML = sources.map(src => `
        <tr class="border-b border-surface-100 dark:border-accent-800">
            <td class="py-2 px-3 font-medium">${esc(src.name || '未設定')}${src.platform_type ? ' <span class="text-[10px] text-surface-400">(API広告費あり・上書き不要)</span>' : ''}</td>
            <td class="py-2 px-3 text-right">
                <input type="number" inputmode="numeric" min="0" step="1000" value="${data.adCosts[String(src.id)] ?? ''}" placeholder="未入力"
                    data-source-id="${src.id}" aria-label="${esc(src.name || '媒体')}の広告費"
                    class="w-32 text-right px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
            </td>
        </tr>`).join('');
}

// ---- 保存 ----
function readEntry() {
    const entry = {};
    for (const [f, id] of Object.entries(FIELD_IDS)) entry[f] = numValue(id);
    return entry;
}

async function saveDaily() {
    const staffId = selectedStaffId();
    if (!staffId) { toast('スタッフを選択してください', 'warn'); return; }
    const entry = readEntry();
    if (DAILY_FIELDS.every(f => entry[f] === null)) {
        toast('数値を入力してください（すべて空のときは「クリア」で削除できます）', 'warn');
        return;
    }
    // 入力の目安: 次回予約数が店舗の来店数を超えていたら確認
    const dayRow = (monthSummary(monthOf(currentDate))?.by_day || []).find(d => d.date === currentDate);
    const visits = dayRow ? (dayRow.new_visit_count || 0) + (dayRow.repeat_visit_count || 0) : null;
    const next = (entry.nextNew || 0) + (entry.nextRepeat || 0);
    if (visits !== null && visits > 0 && next > visits) {
        if (!window.confirm(`次回予約数（${next}）が店舗の来店数（${visits}名）を超えています。このまま保存しますか？`)) return;
    }
    const btn = document.getElementById('input-daily-save');
    btn.disabled = true;
    btn.textContent = '保存中…';
    try {
        const res = await saveManualPatch(monthOf(currentDate), {
            daily: { [`${currentDate}:${staffId}`]: entry },
        });
        dirty = false;
        fillDailyForm();
        toast(res.storage === 'local' ? '保存しました（この端末のみ）' : `${selectedStaffName()} ${fmtDateShort(currentDate)} の日報を保存しました`, 'success');
    } catch (e) {
        console.error('daily save', e);
        toast(e?.body?.detail || '保存に失敗しました。時間をおいて再度お試しください', 'error');
    } finally {
        btn.disabled = false;
        btn.textContent = '保存';
    }
}

async function clearDaily() {
    const staffId = selectedStaffId();
    if (!staffId) return;
    const entry = getDailyEntry(currentDate, staffId);
    if (!hasValues(entry) && !dirty) { for (const id of Object.values(FIELD_IDS)) setValue(id, null); return; }
    if (!window.confirm(`${fmtDateShort(currentDate)} の入力を削除しますか？`)) return;
    try {
        await saveManualPatch(monthOf(currentDate), { daily: { [`${currentDate}:${staffId}`]: null } });
        dirty = false;
        fillDailyForm();
        toast('削除しました');
    } catch (e) {
        console.error('daily clear', e);
        toast('削除に失敗しました', 'error');
    }
}

async function saveMonthly(input) {
    const v = input.value === '' ? null : Number(input.value);
    if (v !== null && (!isFinite(v) || v < 0)) { toast('0以上の金額を入力してください', 'warn'); return; }
    try {
        await saveManualPatch(monthOf(currentDate), {
            monthly: { [input.dataset.staffId]: v === null ? null : { productSales: v } },
        });
        toast('物販売上を保存しました', 'success');
    } catch (e) {
        console.error('monthly save', e);
        toast(e?.body?.detail || '保存に失敗しました', 'error');
    }
}

async function saveAdCost(input) {
    const v = input.value === '' ? null : Number(input.value);
    if (v !== null && (!isFinite(v) || v < 0)) { toast('0以上の金額を入力してください', 'warn'); return; }
    try {
        await saveManualPatch(monthOf(currentDate), { adCosts: { [input.dataset.sourceId]: v } });
        toast('広告費を保存しました', 'success');
    } catch (e) {
        console.error('adcost save', e);
        toast(e?.body?.detail || '保存に失敗しました', 'error');
    }
}

// ---- util ----
function fmtDateShort(date) {
    const [, m, d] = date.split('-').map(Number);
    return `${m}/${d}`;
}
function setValue(id, v) {
    const el = document.getElementById(id);
    if (el) el.value = v ?? '';
}
function numValue(id) {
    const el = document.getElementById(id);
    const v = el?.value;
    if (v === '' || v === undefined) return null;
    const n = Number(v);
    return isFinite(n) && n >= 0 ? Math.round(n) : null;
}
function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

export { BLOG_TARGET };
