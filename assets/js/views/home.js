// ホーム: 役割ごとの「今日やること」+ 今日の実績（スナップショット）+ 今月の進捗 + SalonOneで確認すること
// スマホで開いて最初に見る画面。数字を眺める前に「今やるべき作業」が分かることを優先する。

import {
    state, on, isAdmin, isAdminLike, isStaffLocked, isStoreLocked,
    currentShopId, currentStaffId, shopName, staffsOfShop,
} from '../core/state.js';
import { yen, num, esc, todayJst, todayStr, ymd, daysInMonth } from '../core/format.js';
import { apiGetCached } from '../core/api.js';
import { kpisOf, scopedRow, monthToDate, salesOf } from '../data/salonone.js';
import { getGoal, getGoalRaw, monthKey } from '../data/goals.js';
import { renderRings } from '../core/engage.js';
import { loadManual, getManual, getDailyEntry, hasValues, monthlyTotalsByStaff, emptyTotals, monthKeyOf, monthOf, reconDayStatus } from '../data/manual.js';
import { loadShift, getShift } from '../data/shift.js';
import { switchTab, setHomeBadge } from '../ui/nav.js';
import { shouldShowInstallHint, dismissInstallHint } from '../ui/shell.js';
import { presetInput } from './input.js';
import { presetRecon } from './recon.js';
import { getInsights, loadInsights } from '../data/insights.js';

const DEFAULT_DEADLINE = 20;

// ホーム専用のデータ（今月・昨日の集計。フィルタの対象月とは独立して「今」を見る）
const home = {
    yesterday: null,    // 昨日1日分の sales/summary（スタッフ別の来店有無 = 日報が必要な人の判定）
    mkStaff: null,      // 今月の marketing/by-staff
    channels: null,     // 今月の marketing/by-channel（スタッフ以外）
    shopMonth: {},      // 全店舗表示の管理者向け: 店舗ごとの今月 sales/summary（入金突合の判定）
    loadedFor: null,
};
let todoActions = [];
let todosExpanded = false;
const TODO_VISIBLE = 5;

export function init() {
    on('data:core', render);
    on('data:manual', render);
    on('data:goals', render);
    on('data:shift', render);
    on('data:insights', render);
    on('meta', render);
    on('filters', render);

    document.getElementById('home-todos')?.addEventListener('click', ev => {
        if (ev.target.closest('#home-todo-more')) { todosExpanded = !todosExpanded; renderTodos(); return; }
        const btn = ev.target.closest('button[data-todo]');
        if (!btn) return;
        const act = todoActions[Number(btn.dataset.todo)];
        if (act) act();
    });
    document.getElementById('home-shortcuts')?.addEventListener('click', ev => {
        const btn = ev.target.closest('button[data-goto]');
        if (btn) switchTab(btn.dataset.goto);
    });
    document.getElementById('home-month-more')?.addEventListener('click', () => {
        switchTab(isStaffLocked() || currentStaffId() !== 'all' ? 'staff-dashboard' : 'overview');
    });
    document.getElementById('home-install-dismiss')?.addEventListener('click', () => {
        dismissInstallHint();
        document.getElementById('home-install-hint')?.classList.add('hidden');
    });
}

function active() {
    return state.ui.activeTab === 'home';
}

// ---- 日付ユーティリティ ----
function addDays(date, n) {
    const d = new Date(`${date}T12:00:00Z`);
    d.setUTCDate(d.getUTCDate() + n);
    return ymd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
}
function md(date) {
    const [, m, d] = date.split('-').map(Number);
    return `${m}/${d}`;
}
function jstHour() {
    return new Date(Date.now() + 9 * 3600 * 1000).getUTCHours();
}
function monthShift(y, m, diff) {
    let mm = m + diff, yy = y;
    while (mm < 1) { mm += 12; yy--; }
    while (mm > 12) { mm -= 12; yy++; }
    return { y: yy, m: mm };
}

// ---- 対象範囲 ----
function scopeShopIds() {
    if (isStaffLocked() || isStoreLocked()) return [String(state.session.shopId)];
    const id = currentShopId();
    return id === 'all' ? state.masters.shops.map(s => String(s.id)) : [String(id)];
}
function multiShop() {
    return scopeShopIds().length > 1;
}
function staffFilter() {
    return currentShopId() === 'all' ? {} : { shop_id: currentShopId() };
}

// 店舗の今月サマリ（入金突合の判定に店舗単位の支払い内訳が必要）
function shopMonthSummary(shopId) {
    if (!multiShop()) return state.data.nowMonth;
    return home.shopMonth[shopId] || null;
}

function visitsOf(summary, staffId) {
    const r = (summary?.by_staff || []).find(x => String(x.staff_id) === String(staffId));
    return r ? (r.new_visit_count || 0) + (r.repeat_visit_count || 0) : 0;
}

// ---- 読み込み ----
export async function loadHomeData({ force = false } = {}) {
    const t = todayJst();
    const today = todayStr();
    const y = addDays(today, -1);
    const from = ymd(t.y, t.m, 1);
    const to = ymd(t.y, t.m, daysInMonth(t.y, t.m));
    const next = monthShift(t.y, t.m, 1);
    const months = new Set([monthKeyOf(t.y, t.m), monthOf(y)]);
    if (t.d <= 5) { const p = monthShift(t.y, t.m, -1); months.add(monthKeyOf(p.y, p.m)); }
    const key = `${currentShopId()}:${currentStaffId()}:${today}`;
    if (force || home.loadedFor !== key) home.shopMonth = {};
    home.loadedFor = key;

    const f = staffFilter();
    const tasks = [
        ...[...months].map(m => loadManual(m).catch(() => null)),
        loadShift(monthKeyOf(next.y, next.m)).catch(() => null),
        apiGetCached('sales/summary', { from: y, to: y, ...f }, 300000).then(r => { home.yesterday = r; }).catch(() => null),
        apiGetCached('marketing/by-staff', { from, to, ...f }, 300000).then(r => { home.mkStaff = r.data || []; }).catch(() => null),
        loadInsights({ from, to, force }).catch(() => null),
    ];
    if (!isStaffLocked()) {
        tasks.push(apiGetCached('marketing/by-channel', { from, to, ...f }, 300000).then(r => { home.channels = r.data || []; }).catch(() => null));
    }
    if (isAdminLike() && multiShop()) {
        for (const shop of state.masters.shops) {
            tasks.push(apiGetCached('sales/summary', { from, to, shop_id: shop.id }, 300000)
                .then(r => { home.shopMonth[String(shop.id)] = r; }).catch(() => null));
        }
    }
    await Promise.all(tasks);
    render();
}

// ---- 画面遷移（やることのボタン）----
function openInput(date, staffId) {
    presetInput({ date, staffId });
    switchTab('input');
}
function selectShop(shopId) {
    // 全店舗表示の管理者は、対象店舗を切り替えてから開く（入金突合・シフトは店舗単位）
    if (!isAdminLike() || String(currentShopId()) === String(shopId)) return;
    const sel = document.getElementById('store-selector');
    if (!sel) return;
    sel.value = String(shopId);
    sel.dispatchEvent(new Event('change'));
}
function openRecon(date, shopId) {
    selectShop(shopId);
    presetRecon(date);
    switchTab('recon');
}
function openShift(shopId) {
    if (shopId) selectShop(shopId);
    switchTab('shift');
}

// ---- やること ----
// level: urgent（今日中）/ warn（早めに）/ info（予定）/ done（完了）
function buildTodos() {
    const items = [];
    const t = todayJst();
    const today = todayStr();
    const y = addDays(today, -1);
    const hour = jstHour();
    const shops = scopeShopIds();
    const nameList = list => list.slice(0, 4).map(s => s.name).join('、') + (list.length > 4 ? ` ほか${list.length - 4}名` : '');
    // 複数店舗は「千葉店 3名・大和店 1名」のように店舗ごとの人数で要約
    const whoList = list => {
        if (!multiShop()) return nameList(list);
        const by = new Map();
        for (const st of list) by.set(String(st.shop_id), (by.get(String(st.shop_id)) || 0) + 1);
        return [...by].map(([id, n]) => `${shopName(id)} ${n}名`).join('・');
    };

    // ---- 日報（次回予約など SalonOne にない数字）----
    if (isStaffLocked()) {
        const sid = String(state.session.staffId);
        const vY = visitsOf(home.yesterday, sid);
        const tRow = (state.data.today?.by_staff || []).find(r => String(r.staff_id) === sid);
        const vT = tRow ? (tRow.new_visit_count || 0) + (tRow.repeat_visit_count || 0) : 0;
        if (vY > 0 && !hasValues(getDailyEntry(y, sid))) {
            items.push({ level: 'urgent', icon: 'notebook-pen', title: `昨日（${md(y)}）の日報が未入力です`, desc: `来店 ${vY}名。次回予約の数を入力してください`, label: '入力', run: () => openInput(y, sid) });
        }
        if (hasValues(getDailyEntry(today, sid))) {
            items.push({ level: 'done', icon: 'check', title: '今日の日報は入力済みです', desc: 'お疲れさまでした', label: '確認', run: () => openInput(today, sid) });
        } else if (vT > 0 || hour >= 15) {
            items.push({
                level: hour >= 18 ? 'warn' : 'info', icon: 'notebook-pen', title: '今日の日報',
                desc: vT > 0 ? `今日の来店 ${vT}名（新規 ${tRow.new_visit_count || 0}）。退勤前に次回予約数を入力` : '退勤前に次回予約・ブログ・SNSを入力',
                label: '入力', run: () => openInput(today, sid),
            });
        }
    } else {
        const staffs = shops.flatMap(id => staffsOfShop(id));
        const missY = staffs.filter(s => visitsOf(home.yesterday, s.id) > 0 && !hasValues(getDailyEntry(y, s.id)));
        if (missY.length) {
            items.push({ level: 'urgent', icon: 'notebook-pen', title: `昨日（${md(y)}）の日報 未入力 ${missY.length}名`, desc: `${whoList(missY)}（昨日来店を担当）`, label: '確認', run: () => openInput(y, missY[0].id) });
        }
        if (hour >= 18) {
            const missT = staffs.filter(s => visitsOf(state.data.today, s.id) > 0 && !hasValues(getDailyEntry(today, s.id)));
            if (missT.length) items.push({ level: 'warn', icon: 'notebook-pen', title: `今日の日報 未入力 ${missT.length}名`, desc: whoList(missT), label: '確認', run: () => openInput(today, missT[0].id) });
        }
    }

    // ---- 入金突合（店長・マネージャー・オーナー）----
    if (!isStaffLocked()) {
        const monthKeyNow = monthKeyOf(t.y, t.m);
        const missing = [], diffs = [], pendings = [];
        for (const shopId of shops) {
            const summary = shopMonthSummary(shopId);
            if (!summary) continue;
            const recon = getManual(monthOf(y)).recon || {};
            const rowY = (summary.by_day || []).find(d => d.date === y);
            const entryY = recon[`${y}:${shopId}`] || {};
            const stY = reconDayStatus(rowY, entryY);
            if (stY.state === 'empty' || stY.state === 'partial') missing.push({ shopId, st: stY });
            else if (stY.state === 'diff' && !entryY.memo) diffs.push({ shopId, st: stY });
            // 今月の未突合（昨日より前）。突合を始めた日より前の日は数えない（導入月に過去分まで一気に出さない）
            const monthRecon = getManual(monthKeyNow).recon || {};
            const started = Object.keys(monthRecon).filter(k => k.endsWith(`:${shopId}`)).map(k => k.split(':')[0]).sort()[0];
            if (!started) continue;
            const pending = (summary.by_day || []).filter(d => d.date >= started && d.date < y && d.date.startsWith(monthKeyNow))
                .filter(d => ['empty', 'partial'].includes(reconDayStatus(d, monthRecon[`${d.date}:${shopId}`] || {}).state));
            if (pending.length) pendings.push({ shopId, days: pending.map(d => d.date) });
        }
        if (missing.length === 1 && !multiShop()) {
            const { shopId, st } = missing[0];
            items.push({ level: 'urgent', icon: 'scale', title: `昨日（${md(y)}）の入金突合が${st.state === 'partial' ? '途中です' : '未入力です'}`, desc: `SalonOneの記録 ${yen(st.rec)}。レジ実査・端末集計の金額を入力`, label: '突合', run: () => openRecon(y, shopId) });
        } else if (missing.length) {
            items.push({ level: 'urgent', icon: 'scale', title: `昨日（${md(y)}）の入金突合 未入力 ${missing.length}店舗`, desc: missing.map(m => shopName(m.shopId)).join('・'), label: '突合', run: () => openRecon(y, missing[0].shopId) });
        }
        if (diffs.length) {
            items.push({
                level: 'warn', icon: 'scale',
                title: diffs.length === 1 ? `${multiShop() ? `${shopName(diffs[0].shopId)}: ` : ''}昨日の入金に差異 ${diffs[0].st.diff > 0 ? '+' : ''}${yen(diffs[0].st.diff)}` : `昨日の入金に差異 ${diffs.length}店舗`,
                desc: '原因を確認してメモを残してください（釣銭・返金・支払方法の選び間違い等）',
                label: '確認', run: () => openRecon(y, diffs[0].shopId),
            });
        }
        if (pendings.length) {
            const total = pendings.reduce((a, p) => a + p.days.length, 0);
            items.push({
                level: 'warn', icon: 'calendar-x',
                title: `今月の未突合 ${total}日`,
                desc: multiShop() ? pendings.map(p => `${shopName(p.shopId)} ${p.days.length}日`).join('・') : `最も古い日: ${md(pendings[0].days[0])}`,
                label: '突合', run: () => openRecon(pendings[0].days[0], pendings[0].shopId),
            });
        }
    }

    // ---- シフト（翌月の希望休）----
    const next = monthShift(t.y, t.m, 1);
    const nextKey = monthKeyOf(next.y, next.m);
    const shift = getShift(nextKey);
    if (shift && shift.storage !== 'none') {
        const deadline = shift.config?.requestDeadline ?? DEFAULT_DEADLINE;
        const left = deadline ? deadline - t.d : null;
        if (isStaffLocked()) {
            const sid = String(state.session.staffId);
            const shop = shift.shops?.[String(state.session.shopId)] || {};
            const status = shop.status?.[sid];
            const requested = !!shop.requests?.[sid];
            if (status === 'approved') {
                items.push({ level: 'done', icon: 'calendar-check', title: `${next.m}月の休みが確定しました`, desc: `${(shop.assigned?.[sid] || []).length}日`, label: '見る', run: () => openShift() });
            } else if (!requested) {
                if (!deadline || left >= 0) {
                    items.push({ level: left !== null && left <= 3 ? 'urgent' : 'info', icon: 'calendar-heart', title: `${next.m}月の希望休を申請`, desc: deadline ? `締切 ${t.m}/${deadline}（${left === 0 ? '今日まで' : `あと${left}日`}）` : '休みたい日を選んで申請', label: '申請', run: () => openShift() });
                } else {
                    items.push({ level: 'warn', icon: 'calendar-heart', title: `${next.m}月の希望休が未申請です`, desc: `締切（${t.m}/${deadline}）を過ぎています。店長に相談してください`, label: '申請', run: () => openShift() });
                }
            } else {
                items.push({ level: 'done', icon: 'calendar-check', title: `${next.m}月の希望休を申請済み`, desc: status === 'proposed' ? '店長が調整中です' : '店長の承認待ちです', label: '見る', run: () => openShift() });
            }
        } else {
            const rows = [];
            for (const shopId of shops) {
                const staffs = staffsOfShop(shopId);
                if (!staffs.length) continue;
                const shop = shift.shops?.[shopId] || {};
                const notRequested = staffs.filter(s => !shop.requests?.[String(s.id)]);
                const approved = staffs.filter(s => shop.status?.[String(s.id)] === 'approved');
                if (approved.length === staffs.length) continue;
                rows.push({ shopId, staffs, notRequested, approved });
            }
            if (rows.length) {
                const after = deadline && left < 0;
                if (after || !deadline || left <= 7) {
                    const one = rows.length === 1 && !multiShop();
                    const r0 = rows[0];
                    items.push({
                        level: after ? 'warn' : 'info', icon: 'calendar-clock',
                        title: after ? `${next.m}月のシフトを分配・承認` : `${next.m}月の希望休の申請状況`,
                        desc: one
                            ? (after
                                ? `承認 ${r0.approved.length}/${r0.staffs.length}名${r0.notRequested.length ? `・未申請 ${nameList(r0.notRequested)}` : ''}`
                                : (r0.notRequested.length ? `申請 ${r0.staffs.length - r0.notRequested.length}/${r0.staffs.length}名・未申請 ${nameList(r0.notRequested)}${deadline ? `（締切 ${t.m}/${deadline}）` : ''}` : '全員申請済み。分配・承認できます'))
                            : rows.map(r => `${shopName(r.shopId)} ${after ? `承認${r.approved.length}` : `申請${r.staffs.length - r.notRequested.length}`}/${r.staffs.length}`).join('・') + (after || !deadline ? '' : `（締切 ${t.m}/${deadline}）`),
                        label: 'シフト', run: () => openShift(r0.shopId),
                    });
                }
            }
        }
    }

    // ---- 目標（店長・マネージャー・オーナー）----
    if (isAdminLike() || isStoreLocked()) {
        const mk = monthKey({ y: t.y, m: t.m });
        const noGoal = shops.filter(id => !getGoalRaw(mk, `shop:${id}`) && !staffsOfShop(id).some(s => getGoalRaw(mk, `staff:${s.id}`)));
        if (noGoal.length) {
            items.push({ level: 'warn', icon: 'target', title: `今月の売上目標が未設定${multiShop() ? `（${noGoal.map(shopName).join('、')}）` : ''}`, desc: '目標を入れると、進捗リング・着地予測との比較が使えます', label: '設定', run: () => switchTab('goal') });
        }
        if (t.d >= daysInMonth(t.y, t.m) - 4) {
            const nk = monthKey(next);
            const noNext = shops.filter(id => !getGoalRaw(nk, `shop:${id}`) && !staffsOfShop(id).some(s => getGoalRaw(nk, `staff:${s.id}`)));
            if (noNext.length) items.push({ level: 'info', icon: 'target', title: `${next.m}月の目標を設定`, desc: '月初からスタッフのリングに反映されます', label: '設定', run: () => switchTab('goal') });
        }
    }

    // ---- 管理（オーナー）----
    if (isAdmin() && state.meta) {
        if (state.meta.passwords && !state.meta.passwords.admin) {
            items.push({ level: 'urgent', icon: 'lock', title: 'オーナー画面にパスワードが未設定です', desc: 'URLを知っていれば誰でも全店舗の数字を見られます。Vercelの環境変数 ADMIN_PASSWORD を設定してください', label: '確認', run: () => switchTab('settings') });
        }
        if (state.meta.manualStorage === false) {
            items.push({ level: 'warn', icon: 'database', title: 'サーバー保存（Supabase）が未設定です', desc: '日報・シフト・目標が端末ごとに分かれ、共有されません', label: '確認', run: () => switchTab('settings') });
        }
        if (state.meta.storage?.warning) {
            items.push({ level: 'urgent', icon: 'database', title: 'サーバー保存のキーに問題があります', desc: state.meta.storage.warning, label: '確認', run: () => switchTab('settings') });
        }
        if (t.d <= 5) {
            items.push({ level: 'info', icon: 'coins', title: '先月の締め（インセンティブ・物販の確認）', desc: '対象月を先月にしてインセンティブタブで確認', label: '開く', run: () => switchTab('incentive') });
        }
    }

    const order = { urgent: 0, warn: 1, info: 2, done: 3 };
    return items.sort((a, b) => order[a.level] - order[b.level]);
}

// ---- SalonOneで確認すること（入力漏れの検知）----
function buildQuality() {
    if (isStaffLocked()) return [];
    const items = [];
    const mk = home.mkStaff || [];
    const unassigned = mk.filter(r => !r.is_total && (r.staff_id === null || r.staff_id === undefined))
        .reduce((a, r) => a + (r.new_booking_count || 0), 0);
    if (unassigned > 0) {
        items.push({ icon: 'user-x', title: `担当者が未設定の新規予約 ${num(unassigned)}件`, desc: '予約の担当スタッフを設定すると、スタッフ別の新規数・入会率に反映されます' });
    }
    const noSource = (home.channels || []).filter(c => c.visit_source_id === null || c.visit_source_id === undefined)
        .reduce((a, c) => a + (c.booking_count || 0), 0);
    if (noSource > 0) {
        items.push({ icon: 'megaphone', title: `流入元が未設定の新規客 ${num(noSource)}人`, desc: '顧客の流入元（ホットペッパー・Instagram等）を登録すると、媒体別の集客効果が正しく出ます' });
    }
    const sum = state.data.nowMonth;
    const noStaffSales = (sum?.by_staff || []).filter(r => r.staff_id === null || r.staff_id === undefined).reduce((a, r) => a + salesOf(r), 0);
    if (noStaffSales > 0) {
        items.push({ icon: 'receipt', title: `担当者なしの売上 ${yen(noStaffSales)}`, desc: '会計の担当スタッフを設定すると、スタッフ別売上・インセンティブに反映されます' });
    }
    const unknownPay = (sum?.by_day || []).flatMap(d => d.payment_breakdown || [])
        .filter(p => p.is_sales !== false && !p.name && (p.amount || 0) > 0).reduce((a, p) => a + p.amount, 0);
    if (unknownPay > 0) {
        items.push({ icon: 'wallet', title: `支払い方法が未設定の会計 ${yen(unknownPay)}`, desc: '入金突合ができないため、会計時に支払い方法を選んでください' });
    }
    // 予約明細（β）: 過去日なのに会計・キャンセル処理がされていない予約
    const t = todayJst();
    const ins = getInsights(monthKeyOf(t.y, t.m));
    if (ins?.unsettled?.count > 0) {
        const dates = Object.keys(ins.unsettled.byDate || {}).sort();
        items.push({ icon: 'calendar-x', title: `会計が完了していない過去の予約 ${num(ins.unsettled.count)}件`, desc: `${dates.length ? `最も古い日: ${md(dates[0])}。` : ''}会計漏れ・キャンセル処理漏れがないか確認してください（予約データから自動検知・β）` });
    }
    return items;
}

// ---- 描画 ----
function render() {
    if (!state.data.summary) return;
    renderInstallHint();
    renderTodos();
    renderQuality();
    renderMonth();
    renderShortcuts();
}

function renderInstallHint() {
    document.getElementById('home-install-hint')?.classList.toggle('hidden', !shouldShowInstallHint());
}

function renderTodos() {
    const list = document.getElementById('home-todos');
    if (!list) return;
    const items = buildTodos();
    const open = items.filter(i => i.level === 'urgent' || i.level === 'warn');
    setHomeBadge(open.length);
    const countEl = document.getElementById('home-todo-count');
    if (countEl) {
        countEl.textContent = `${open.length}件`;
        countEl.classList.toggle('hidden', open.length === 0);
    }
    todoActions = items.map(i => i.run);
    if (!items.length) {
        list.innerHTML = '<li class="todo-empty"><i data-lucide="check-check" class="w-4 h-4"></i>いまやることはありません</li>';
    } else {
        const shown = todosExpanded ? items : items.slice(0, TODO_VISIBLE);
        const rest = items.length - shown.length;
        list.innerHTML = shown.map((i, idx) => `
            <li class="todo-item ${i.level}">
                <span class="todo-dot"><i data-lucide="${i.icon}"></i></span>
                <div class="todo-body">
                    <p class="todo-title">${esc(i.title)}</p>
                    ${i.desc ? `<p class="todo-desc">${esc(i.desc)}</p>` : ''}
                </div>
                ${i.run ? `<button type="button" class="todo-action" data-todo="${idx}">${esc(i.label || '開く')}</button>` : ''}
            </li>`).join('')
            + (rest > 0 || todosExpanded && items.length > TODO_VISIBLE
                ? `<li><button type="button" id="home-todo-more" class="todo-more">${rest > 0 ? `ほか ${rest}件を表示` : '閉じる'}</button></li>`
                : '');
    }
    if (window.lucide) lucide.createIcons();
}

function renderQuality() {
    const card = document.getElementById('home-quality-card');
    const list = document.getElementById('home-quality');
    if (!card || !list) return;
    const items = buildQuality();
    card.classList.toggle('hidden', items.length === 0);
    list.innerHTML = items.map(i => `
        <li class="todo-item warn">
            <span class="todo-dot"><i data-lucide="${i.icon}"></i></span>
            <div class="todo-body">
                <p class="todo-title">${esc(i.title)}</p>
                <p class="todo-desc">${esc(i.desc)}</p>
            </div>
        </li>`).join('');
    if (window.lucide) lucide.createIcons();
}

// 今月の進捗リング: 売上・新規来店・次回予約率・入会
function renderMonth() {
    const t = todayJst();
    const staffScope = currentStaffId() !== 'all';
    const now = state.data.nowMonth;
    if (!now) return;
    const row = staffScope ? scopedRow(now) : monthToDate(now);
    const k = kpisOf(row);
    const goal = getGoal(monthKey({ y: t.y, m: t.m }), currentShopId(), currentStaffId()) || {};

    // 次回予約率（日報）: 対象スタッフの手入力合計 ÷ SalonOneの来店数
    const totals = monthlyTotalsByStaff(monthKeyOf(t.y, t.m));
    const staffIds = staffScope ? [String(currentStaffId())] : scopeShopIds().flatMap(id => staffsOfShop(id)).map(s => String(s.id));
    let next = 0, visits = 0, days = 0;
    for (const id of staffIds) {
        const tt = totals[id] || emptyTotals();
        next += tt.nextNew + tt.nextRepeat;
        days += tt.days;
        visits += visitsOf(now, id);
    }
    const nextRate = visits > 0 && next > 0 ? next / visits * 100 : null;

    const mkRow = staffScope
        ? (home.mkStaff || []).find(r => String(r.staff_id) === String(currentStaffId()))
        : (home.mkStaff || []).find(r => r.is_total);
    const joins = mkRow?.purchase_in_period_count || 0;

    const rings = [
        { label: '売上', color: '#b8956a', pct: goal.sales > 0 ? k.sales / goal.sales * 100 : 0, value: yen(k.sales), sub: goal.sales > 0 ? `目標 ${yen(goal.sales)}` : '目標未設定' },
        { label: '新規来店', color: '#739977', pct: goal.newVisits > 0 ? k.newVisits / goal.newVisits * 100 : 0, value: `${num(k.newVisits)}名`, sub: goal.newVisits > 0 ? `目標 ${num(goal.newVisits)}名` : '目標未設定' },
        { label: '次回予約率', color: '#566882', pct: nextRate ?? 0, value: nextRate === null ? '—' : `${num(next)}/${num(visits)}名`, sub: nextRate === null ? '日報から集計' : '日報の入力から' },
        { label: '入会', color: '#c9a96e', pct: goal.joins > 0 ? joins / goal.joins * 100 : 0, value: `${num(joins)}名`, sub: goal.joins > 0 ? `目標 ${num(goal.joins)}名` : '期間内の契約' },
    ];
    renderRings('home-rings', rings);
    document.getElementById('home-rings')?.classList.add('cols-4');

    const title = document.getElementById('home-month-title');
    if (title) title.textContent = isStaffLocked() ? `${t.m}月のあなたの進捗` : `${t.m}月の進捗`;

    // 補足の数字
    const stats = [];
    stats.push({ label: '客単価', value: yen(k.unitPrice), sub: `${num(k.visits)}名来店` });
    if (staffScope) {
        const sorted = [...(now.by_staff || [])].sort((a, b) => salesOf(b) - salesOf(a));
        const rank = sorted.findIndex(r => String(r.staff_id) === String(currentStaffId()));
        stats.push({ label: '店内順位（売上）', value: rank >= 0 ? `${rank + 1}位` : '—', sub: `${sorted.length}人中` });
        stats.push({ label: '日報の入力', value: `${num(days)}日`, sub: `${t.d}日経過` });
    } else {
        stats.push({ label: 'キャンセル率', value: `${k.cancelRate.toFixed(1)}%`, sub: `無断 ${num(k.noShows)}件` });
        const expected = staffIds.length * t.d;
        stats.push({ label: '日報の入力率', value: expected > 0 ? `${Math.round(days / expected * 100)}%` : '—', sub: `${num(days)}/${num(expected)}人日` });
    }
    const statsEl = document.getElementById('home-month-stats');
    if (statsEl) {
        statsEl.innerHTML = stats.map(s => `
            <div class="home-stat">
                <p class="home-stat-label">${s.label}</p>
                <p class="home-stat-value">${s.value}</p>
                <p class="home-stat-sub">${s.sub}</p>
            </div>`).join('');
    }
}

function renderShortcuts() {
    const el = document.getElementById('home-shortcuts');
    if (!el) return;
    const r = state.session?.role;
    const sets = {
        staff: [['input', 'notebook-pen', '日報入力'], ['staff-dashboard', 'user-round', 'マイ成績'], ['shift', 'calendar-heart', '希望休'], ['guide', 'book-open', '使い方']],
        store: [['input', 'notebook-pen', '日報入力'], ['recon', 'scale', '入金突合'], ['shift', 'calendar-clock', 'シフト'], ['goal', 'target', '目標']],
        manager: [['overview', 'layout-dashboard', 'サマリー'], ['marketing', 'megaphone', 'マーケ'], ['shift', 'calendar-clock', 'シフト'], ['goal', 'target', '目標']],
        admin: [['overview', 'layout-dashboard', 'サマリー'], ['marketing', 'megaphone', 'マーケ'], ['incentive', 'coins', '歩合'], ['settings', 'settings', '設定']],
    };
    el.innerHTML = (sets[r] || sets.admin).map(([tab, icon, label]) => `
        <button type="button" class="home-shortcut" data-goto="${tab}"><i data-lucide="${icon}"></i><span>${label}</span></button>`).join('');
    if (window.lucide) lucide.createIcons();
}

// テスト用
export const _internal = { buildTodos, buildQuality, home };
export { active as homeActive };
