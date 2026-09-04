// シフトタブ: 希望休の申請（スタッフ）と 分配・承認（店長/マネージャー/オーナー)
// ルール: 月{offDays}日休み・うち土日は{weekendOffDays}日・自動割当は同日{maxSameDayOff}人まで
// （ルールは設定タブから変更可能）

import { state, on, isAdminLike, isStaffLocked, isStoreLocked, currentShopId, shopName, staffsOfShop } from '../core/state.js';
import { esc, todayJst, ymd, daysInMonth, dowIndex } from '../core/format.js';
import { loadShift, submitShiftRequest, saveAssignments, approveShift, getShift, autoDistribute } from '../data/shift.js';
import { toast } from '../core/engage.js';
import { ApiError } from '../core/api.js';
import { renderShopPick } from '../ui/shoppick.js';

let shiftAnchor = null;        // {y, m} 表示月（既定: 来月）
let selection = new Set();     // スタッフモード: 希望日の選択
let selectionLoaded = false;
let editingStaffId = null;     // 管理モード: 割当編集中のスタッフ
let editingDays = new Set();
let lastWarnings = [];

export function init() {
    on('tab:shown', id => { if (id === 'shift') refresh(); });
    on('data:shift', render);
    // ヘッダーで店舗を切り替えたら（ガード画面のクイック選択を含む）その場で描き直す
    on('filters', () => {
        if (state.ui.activeTab !== 'shift') return;
        editingStaffId = null;
        lastWarnings = [];
        render();
    });
    document.getElementById('shift-prev')?.addEventListener('click', () => shiftMonthBy(-1));
    document.getElementById('shift-next')?.addEventListener('click', () => shiftMonthBy(1));
    document.getElementById('shift-submit-btn')?.addEventListener('click', submitRequest);
    document.getElementById('shift-auto-btn')?.addEventListener('click', runAutoDistribute);
    document.getElementById('shift-approve-all-btn')?.addEventListener('click', approveAll);

    document.addEventListener('click', ev => {
        const cell = ev.target.closest('[data-shift-day]');
        if (cell) { onDayTap(cell.dataset.shiftDay, cell.closest('#shift-request-cal') ? 'request' : 'edit'); return; }
        const editBtn = ev.target.closest('[data-shift-edit]');
        if (editBtn) { startEditing(editBtn.dataset.shiftEdit); return; }
        const approveBtn = ev.target.closest('[data-shift-approve]');
        if (approveBtn) { approveOne(approveBtn.dataset.shiftApprove); return; }
        if (ev.target.closest('#shift-edit-save')) { saveEditing(); return; }
        if (ev.target.closest('#shift-edit-cancel')) { editingStaffId = null; render(); }
    });
}

function ensureAnchor() {
    if (shiftAnchor) return;
    const t = todayJst();
    let y = t.y, m = t.m + 1;
    if (m > 12) { m = 1; y++; }
    shiftAnchor = { y, m };
}

function monthStr() {
    ensureAnchor();
    return `${shiftAnchor.y}-${String(shiftAnchor.m).padStart(2, '0')}`;
}

function shiftMonthBy(diff) {
    ensureAnchor();
    let { y, m } = shiftAnchor;
    m += diff;
    while (m < 1) { m += 12; y--; }
    while (m > 12) { m -= 12; y++; }
    shiftAnchor = { y, m };
    selectionLoaded = false;
    editingStaffId = null;
    lastWarnings = [];
    refresh();
}

function activeShopId() {
    const id = currentShopId();
    return id === 'all' ? null : String(id);
}

function shopData() {
    const shopId = activeShopId();
    const data = getShift(monthStr());
    return (shopId && data?.shops?.[shopId]) || { requests: {}, assigned: {}, status: {} };
}

function config() {
    return getShift(monthStr())?.config || { offDays: 8, weekendOffDays: 1, maxSameDayOff: 2 };
}

async function refresh() {
    ensureAnchor();
    setText('shift-month-label', `${shiftAnchor.y}年${shiftAnchor.m}月`);
    try {
        await loadShift(monthStr());
    } catch (e) {
        console.warn('shift load', e);
        if (e instanceof ApiError && e.status === 401) toast('再ログインが必要です', 'error');
    }
    render();
}

function render() {
    const data = getShift(monthStr());
    const cfg = config();
    setText('shift-month-label', `${shiftAnchor.y}年${shiftAnchor.m}月`);
    setText('shift-rules-label', `ルール: 月${cfg.offDays}日休み ・ 土日休み${cfg.weekendOffDays}日 ・ 自動割当は同日${cfg.maxSameDayOff}人まで`);

    // サーバー保存が無いとシフトは使えない
    const notice = document.getElementById('shift-storage-notice');
    const usable = data?.storage !== 'none';
    if (notice) notice.classList.toggle('hidden', usable);

    const staffMode = isStaffLocked();
    document.getElementById('shift-staff-section')?.classList.toggle('hidden', !staffMode || !usable);
    document.getElementById('shift-manage-section')?.classList.toggle('hidden', staffMode || !usable);
    if (!usable) return;

    if (staffMode) renderStaffMode();
    else renderManageMode();
}

// ================= スタッフモード =================
function renderStaffMode() {
    const staffId = String(state.session.staffId);
    const shop = shopData();
    const cfg = config();
    const myStatus = shop.status[staffId];
    const myAssigned = shop.assigned[staffId] || [];
    const myRequest = shop.requests[staffId]?.days || [];

    if (!selectionLoaded) {
        selection = new Set(myRequest);
        selectionLoaded = true;
    }

    renderCalendar('shift-request-cal', {
        marked: selection,
        secondary: new Set(),
        tappable: myStatus !== 'approved',
    });

    const weekendUsed = [...selection].filter(d => [0, 6].includes(dowIndex(d))).length;
    setText('shift-counter', `選択中: ${selection.size} / ${cfg.offDays}日（土日 ${weekendUsed} / ${cfg.weekendOffDays}）`);

    const statusEl = document.getElementById('shift-my-status');
    if (statusEl) {
        statusEl.innerHTML = myStatus
            ? `<span class="px-3 py-1 rounded-full text-xs font-semibold ${statusBadgeClass(myStatus)}">${statusLabel(myStatus)}</span>`
            : '<span class="text-xs text-surface-500">未申請</span>';
    }
    const submitBtn = document.getElementById('shift-submit-btn');
    if (submitBtn) submitBtn.disabled = myStatus === 'approved';

    // 承認/調整結果の表示
    const resultCard = document.getElementById('shift-result-card');
    if (resultCard) {
        const show = myAssigned.length > 0;
        resultCard.classList.toggle('hidden', !show);
        if (show) {
            setText('shift-result-title', myStatus === 'approved' ? '✅ 承認されたお休み' : '調整中のお休み（承認待ち）');
            renderCalendar('shift-result-cal', { marked: new Set(myAssigned), secondary: new Set(myRequest), tappable: false });
        }
    }
}

function onDayTap(day, mode) {
    if (mode === 'request') {
        if (!isStaffLocked()) return;
        const cfg = config();
        if (selection.has(day)) selection.delete(day);
        else {
            if (selection.size >= cfg.offDays) { toast(`希望は月${cfg.offDays}日までです`, 'warn'); return; }
            const isW = [0, 6].includes(dowIndex(day));
            const weekendUsed = [...selection].filter(d => [0, 6].includes(dowIndex(d))).length;
            if (isW && weekendUsed >= cfg.weekendOffDays) { toast(`土日の休みは${cfg.weekendOffDays}日までです`, 'warn'); return; }
            selection.add(day);
        }
        renderStaffMode();
    } else if (editingStaffId) {
        if (editingDays.has(day)) editingDays.delete(day);
        else editingDays.add(day);
        renderManageMode();
    }
}

async function submitRequest() {
    const btn = document.getElementById('shift-submit-btn');
    btn.disabled = true;
    try {
        await submitShiftRequest(monthStr(), [...selection].sort());
        toast('希望休を申請しました');
    } catch (e) {
        console.error('shift request', e);
        toast(e instanceof ApiError && e.body?.detail ? e.body.detail : '申請に失敗しました', 'error');
    } finally {
        btn.disabled = false;
    }
}

// ================= 管理モード（店長/マネージャー/オーナー） =================
function renderManageMode() {
    const shopId = activeShopId();
    const guard = document.getElementById('shift-shop-guard');
    const bodyWrap = document.getElementById('shift-manage-body');
    if (guard && bodyWrap) {
        guard.classList.toggle('hidden', !!shopId);
        bodyWrap.classList.toggle('hidden', !shopId);
    }
    if (!shopId) { renderShopPick('shift-shop-pick'); return; }
    setText('shift-shop-label', shopName(shopId));

    const shop = shopData();
    const staffs = staffsOfShop(shopId);
    const cfg = config();

    // スタッフ一覧
    const body = document.getElementById('shift-staff-body');
    if (body) {
        body.innerHTML = staffs.map(st => {
            const id = String(st.id);
            const req = shop.requests[id];
            const assigned = shop.assigned[id] || [];
            const status = shop.status[id];
            return `
            <tr class="border-b border-surface-100 dark:border-accent-800 ${editingStaffId === id ? 'bg-primary-50 dark:bg-primary-900/20' : ''}">
                <td class="py-2 px-3 font-medium">${esc(st.name)}</td>
                <td class="py-2 px-3 text-right tabular-nums">${req ? `${req.days.length}日` : '<span class="text-surface-400">未申請</span>'}</td>
                <td class="py-2 px-3 text-right tabular-nums">${assigned.length ? `${assigned.length}日` : '—'}</td>
                <td class="py-2 px-3 text-center"><span class="px-2 py-0.5 rounded-full text-[11px] font-semibold ${statusBadgeClass(status)}">${statusLabel(status)}</span></td>
                <td class="py-2 px-3 text-right whitespace-nowrap">
                    <button data-shift-edit="${id}" class="btn-secondary py-1 px-3 text-xs">${assigned.length ? '編集' : '割当'}</button>
                    ${assigned.length && status !== 'approved' ? `<button data-shift-approve="${id}" class="btn-sage py-1 px-3 text-xs ml-1">承認</button>` : ''}
                </td>
            </tr>`;
        }).join('') || '<tr><td colspan="5" class="py-6 text-center text-surface-500">スタッフがいません</td></tr>';
    }

    // 混雑状況（日別の休み人数）
    renderCoverage(shop, staffs, cfg);

    // 警告
    const warnEl = document.getElementById('shift-warnings');
    if (warnEl) {
        warnEl.classList.toggle('hidden', lastWarnings.length === 0);
        if (lastWarnings.length) {
            const nameOf = id => staffs.find(s => String(s.id) === String(id))?.name || id;
            warnEl.innerHTML = '<p class="font-semibold mb-1">⚠ 自動分配の調整内容:</p>' +
                lastWarnings.map(w => `<p>・${esc(nameOf(w.staffId))} ${Number(w.day.slice(8))}日 — ${esc(w.reason)}</p>`).join('');
        }
    }

    // 編集パネル
    const editPanel = document.getElementById('shift-edit-panel');
    if (editPanel) {
        editPanel.classList.toggle('hidden', !editingStaffId);
        if (editingStaffId) {
            const st = staffs.find(s => String(s.id) === editingStaffId);
            setText('shift-edit-name', `${st?.name || editingStaffId} の休み（タップで変更）`);
            const requested = new Set(shop.requests[editingStaffId]?.days || []);
            renderCalendar('shift-edit-cal', { marked: editingDays, secondary: requested, tappable: true });
            const weekendUsed = [...editingDays].filter(d => [0, 6].includes(dowIndex(d))).length;
            setText('shift-edit-counter', `${editingDays.size} / ${cfg.offDays}日（土日 ${weekendUsed} / ${cfg.weekendOffDays}）◆=本人の希望日`);
        }
    }
}

function renderCoverage(shop, staffs, cfg) {
    const el = document.getElementById('shift-coverage');
    if (!el) return;
    const count = {};
    for (const days of Object.values(shop.assigned)) {
        for (const d of days || []) count[d] = (count[d] || 0) + 1;
    }
    const dim = daysInMonth(shiftAnchor.y, shiftAnchor.m);
    const cells = [];
    for (let d = 1; d <= dim; d++) {
        const date = ymd(shiftAnchor.y, shiftAnchor.m, d);
        const c = count[date] || 0;
        const over = c > cfg.maxSameDayOff;
        const isW = [0, 6].includes(dowIndex(date));
        cells.push(`<div class="text-center rounded px-0.5 py-1 ${over ? 'bg-red-100 dark:bg-red-900/30 text-red-600 font-bold' : c > 0 ? 'bg-surface-100 dark:bg-gray-700/50' : ''}">
            <div class="text-[9px] ${isW ? 'text-rose-400' : 'text-surface-400'}">${d}</div>
            <div class="text-[11px] tabular-nums ${c ? '' : 'text-surface-300 dark:text-gray-600'}">${c || '·'}</div>
        </div>`);
    }
    el.innerHTML = cells.join('');
}

async function runAutoDistribute() {
    const shopId = activeShopId();
    if (!shopId) return;
    const btn = document.getElementById('shift-auto-btn');
    btn.disabled = true;
    try {
        const shop = shopData();
        const staffs = staffsOfShop(shopId);
        // 承認済みは固定し、それ以外を分配対象にする
        const targets = staffs.map(s => String(s.id)).filter(id => shop.status[id] !== 'approved');
        const fixed = {};
        for (const [id, days] of Object.entries(shop.assigned)) {
            if (shop.status[id] === 'approved') fixed[id] = days;
        }
        const result = autoDistribute({
            month: monthStr(),
            staffIds: targets,
            requests: shop.requests,
            existingAssigned: fixed,
            config: config(),
        });
        lastWarnings = result.warnings;
        await saveAssignments(monthStr(), shopId, result.assigned);
        toast(`自動分配しました（${targets.length}名）`);
    } catch (e) {
        console.error('auto distribute', e);
        toast('自動分配に失敗しました', 'error');
    } finally {
        btn.disabled = false;
    }
}

function startEditing(staffId) {
    const shop = shopData();
    editingStaffId = String(staffId);
    editingDays = new Set(shop.assigned[editingStaffId] || shop.requests[editingStaffId]?.days || []);
    render();
    document.getElementById('shift-edit-panel')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function saveEditing() {
    const shopId = activeShopId();
    if (!shopId || !editingStaffId) return;
    try {
        await saveAssignments(monthStr(), shopId, { [editingStaffId]: [...editingDays].sort() });
        toast('割当を保存しました');
        editingStaffId = null;
    } catch (e) {
        console.error('shift edit save', e);
        toast('保存に失敗しました', 'error');
    }
}

async function approveOne(staffId) {
    const shopId = activeShopId();
    if (!shopId) return;
    try {
        await approveShift(monthStr(), shopId, [staffId]);
        toast('承認しました');
    } catch (e) {
        console.error('approve', e);
        toast('承認に失敗しました', 'error');
    }
}

async function approveAll() {
    const shopId = activeShopId();
    if (!shopId) return;
    const shop = shopData();
    const ids = Object.keys(shop.assigned).filter(id => shop.status[id] !== 'approved');
    if (ids.length === 0) { toast('承認対象がありません', 'warn'); return; }
    try {
        await approveShift(monthStr(), shopId, ids);
        toast(`${ids.length}名を承認しました`);
    } catch (e) {
        console.error('approve all', e);
        toast('承認に失敗しました', 'error');
    }
}

// ================= カレンダー描画（共通） =================
// marked=休み（●） secondary=本人の希望日（◆マーク）
function renderCalendar(containerId, { marked, secondary, tappable }) {
    const el = document.getElementById(containerId);
    if (!el) return;
    const { y, m } = shiftAnchor;
    const dim = daysInMonth(y, m);
    const firstDow = dowIndex(ymd(y, m, 1));
    const cells = [];
    for (let i = 0; i < firstDow; i++) cells.push('<div></div>');
    for (let d = 1; d <= dim; d++) {
        const date = ymd(y, m, d);
        const isW = [0, 6].includes(dowIndex(date));
        const isMarked = marked.has(date);
        const isReq = secondary.has(date);
        cells.push(`
        <div ${tappable ? `data-shift-day="${date}" role="button" tabindex="0"` : ''}
            class="rounded-lg p-1.5 min-h-[52px] border text-center select-none transition-colors
                ${isMarked ? 'bg-primary-500 border-primary-500 text-white font-bold' : `border-surface-100 dark:border-gray-700/60 ${tappable ? 'cursor-pointer hover:bg-surface-100 dark:hover:bg-gray-700/50' : ''}`}">
            <div class="text-[11px] ${isMarked ? '' : isW ? 'text-rose-400' : 'text-surface-500'}">${d}</div>
            <div class="text-[13px] leading-4">${isMarked ? '休' : ''}${!isMarked && isReq ? '<span class="text-primary-400">◆</span>' : ''}</div>
        </div>`);
    }
    el.innerHTML = `
        <div class="grid grid-cols-7 gap-1 text-center text-[10px] text-surface-500 mb-1">
            <span class="text-rose-500">日</span><span>月</span><span>火</span><span>水</span><span>木</span><span>金</span><span class="text-accent-500">土</span>
        </div>
        <div class="grid grid-cols-7 gap-1">${cells.join('')}</div>`;
}

function statusLabel(status) {
    return status === 'approved' ? '承認済み' : status === 'proposed' ? '調整中' : status === 'requested' ? '申請中' : '未申請';
}

function statusBadgeClass(status) {
    return status === 'approved' ? 'bg-green-100 text-green-700 dark:bg-green-900/40 dark:text-green-300'
        : status === 'proposed' ? 'bg-amber-100 text-amber-700 dark:bg-amber-900/40 dark:text-amber-300'
        : status === 'requested' ? 'bg-primary-100 text-primary-700 dark:bg-primary-900/40 dark:text-primary-300'
        : 'bg-surface-100 text-surface-500 dark:bg-gray-700 dark:text-gray-400';
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
