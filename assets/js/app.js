// エントリポイント: 認証 → マスタ取得 → 初期表示 → 自動更新
import { ensureAuthenticated } from './core/auth.js';
import { state, emit, on, isLocked, isStaffLocked, currentShopId } from './core/state.js';
import { apiGet, clearApiCache, ApiError } from './core/api.js';
import { loadMasters, loadCore, loadMarketing, loadRetention, loadAgeDist } from './data/salonone.js';
import { initChartDefaults, refreshChartsTheme } from './core/charts.js';
import { setLiveIndicator, toast } from './core/engage.js';
import { initNav, renderNav, switchTab, defaultTab } from './ui/nav.js';
import { todayJst, esc } from './core/format.js';

import * as overview from './views/overview.js';
import * as staffView from './views/staff.js';
import * as salesView from './views/sales.js';
import * as marketingView from './views/marketing.js';
import * as customersView from './views/customers.js';
import * as calendarView from './views/calendar.js';
import * as incentiveView from './views/incentive.js';
import * as goalView from './views/goal.js';
import * as settingsView from './views/settings.js';
import * as inputView from './views/input.js';
import * as reconView from './views/recon.js';
import * as shiftView from './views/shift.js';
import { loadManual, monthKeyOf } from './data/manual.js';
import { loadGoals } from './data/goals.js';

const VIEWS = [overview, staffView, salesView, marketingView, customersView, calendarView, incentiveView, goalView, settingsView, inputView, reconView, shiftView];

const REFRESH_INTERVAL = 5 * 60 * 1000;

// ---- ダークモード ----
function initDarkMode() {
    const saved = localStorage.getItem('darkMode');
    const prefer = matchMedia('(prefers-color-scheme: dark)').matches;
    const dark = saved === null ? prefer : saved === 'true';
    document.documentElement.classList.toggle('dark', dark);
    updateDarkToggle(dark);
}

function updateDarkToggle(dark) {
    const icon = document.getElementById('dark-mode-icon');
    const label = document.getElementById('dark-mode-label');
    if (icon) icon.setAttribute('data-lucide', dark ? 'sun' : 'moon');
    if (label) label.textContent = dark ? 'ライト' : 'ダーク';
    if (window.lucide) lucide.createIcons();
}

function toggleDarkMode() {
    const dark = !document.documentElement.classList.contains('dark');
    document.documentElement.classList.toggle('dark', dark);
    localStorage.setItem('darkMode', String(dark));
    updateDarkToggle(dark);
    refreshChartsTheme();
    emit('theme');
}

// ---- フィルタUI ----
function populateSelectors() {
    const shopSel = document.getElementById('store-selector');
    if (shopSel) {
        shopSel.innerHTML = '<option value="all">全店舗</option>' +
            state.masters.shops.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
        shopSel.value = String(state.filters.shopId);
    }
    populateStaffSelector();
    populateDateSelector();
}

function populateStaffSelector() {
    const sel = document.getElementById('staff-selector');
    if (!sel) return;
    if (isStaffLocked()) {
        sel.innerHTML = `<option value="${state.session.staffId}">自分の成績</option><option value="all">店舗全体</option>`;
        sel.value = state.filters.staffId === 'all' ? 'all' : String(state.session.staffId);
        return;
    }
    const shopId = currentShopId();
    const staffs = shopId === 'all'
        ? state.masters.staffs
        : state.masters.staffs.filter(s => String(s.shop_id) === String(shopId));
    sel.innerHTML = '<option value="all">店舗合計</option>' +
        staffs.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
    sel.value = staffs.some(s => String(s.id) === String(state.filters.staffId)) ? String(state.filters.staffId) : 'all';
    if (sel.value === 'all') state.filters.staffId = 'all';
}

function populateDateSelector() {
    const sel = document.getElementById('date-selector');
    if (!sel) return;
    const t = todayJst();
    const options = [];
    // 過去24ヶ月〜今月
    for (let i = 0; i < 24; i++) {
        let y = t.y, m = t.m - i;
        while (m < 1) { m += 12; y -= 1; }
        options.push({ y, m });
    }
    sel.innerHTML = options.map(o => `<option value="${o.y}-${o.m}">${o.y}年${o.m}月</option>`).join('');
    sel.value = `${state.filters.anchor.y}-${state.filters.anchor.m}`;
}

function applyRoleUi() {
    const session = state.session;
    const badge = document.getElementById('staff-mode-badge');
    const shopSel = document.getElementById('store-selector');

    if (session.role === 'staff') {
        badge?.classList.remove('hidden');
        if (badge) badge.textContent = `${session.staffName} 専用`;
        shopSel?.closest('div')?.classList.add('hidden');
        // スタッフは「自分 / 店舗全体」の2択で切り替え可能
        state.filters.shopId = session.shopId;
        state.filters.staffId = session.staffId;
    } else if (session.role === 'store') {
        badge?.classList.remove('hidden');
        if (badge) badge.textContent = `${session.shopName} 店長`;
        shopSel?.closest('div')?.classList.add('hidden');
        state.filters.shopId = session.shopId;
    } else if (session.role === 'manager') {
        badge?.classList.remove('hidden');
        if (badge) badge.textContent = 'マネージャー';
    }

    if (state.demo) {
        document.getElementById('demo-badge')?.classList.remove('hidden');
    }
}

// ---- データロード ----
let loading = false;

async function refreshAll({ silent = false } = {}) {
    if (loading) return;
    loading = true;
    setLiveIndicator('syncing');
    if (!silent) document.body.classList.add('is-loading');
    try {
        await loadCore();
        state.ui.loadError = null;
        setLiveIndicator('ok', Date.now());
        // 表示中のタブが遅延データを使うなら追加ロード
        await loadTabData(state.ui.activeTab, { force: true });
    } catch (e) {
        console.error('データ取得エラー', e);
        state.ui.loadError = e;
        setLiveIndicator('error');
        if (e instanceof ApiError && e.status === 401) {
            // セッション期限切れ（24時間）: 再読み込みで再認証する
            toast('セッションの有効期限が切れました。再読み込みします', 'warn');
            setTimeout(() => location.reload(), 1500);
        } else if (e instanceof ApiError && e.code === 'rate_limited') {
            toast(`アクセスが集中しています。${e.body.retryAfter || 30}秒後に自動再試行します`, 'warn');
            setTimeout(() => refreshAll({ silent: true }), (e.body.retryAfter || 30) * 1000);
        } else if (e instanceof ApiError && e.code === 'upstream_auth') {
            toast('SalonOne APIキーが無効です。設定を確認してください', 'error');
        } else if (!silent) {
            toast('データの取得に失敗しました。時間をおいて再度お試しください', 'error');
        }
    } finally {
        loading = false;
        document.body.classList.remove('is-loading');
    }
}

// 遅延ロードの失敗処理: レート制限は自動再試行、それ以外は警告のみ
const lazyRetryTimers = {};
function lazyCatch(name, retryFn) {
    return (e) => {
        if (e instanceof ApiError && e.code === 'rate_limited') {
            const wait = (e.body.retryAfter || 30) * 1000;
            toast(`アクセスが集中しています。${Math.round(wait / 1000)}秒後に再取得します`, 'warn');
            clearTimeout(lazyRetryTimers[name]);
            lazyRetryTimers[name] = setTimeout(() => retryFn().catch(err => console.warn(`${name} retry`, err)), wait);
        } else if (e instanceof ApiError && e.status === 403) {
            // 役割制限で見えないデータ（正常系）
        } else {
            console.warn(`${name} load`, e);
        }
    };
}

// 手入力データを読む月: 今月 + 表示中の対象月（単月表示のとき）
function manualMonths() {
    const t = todayJst();
    const months = new Set([monthKeyOf(t.y, t.m)]);
    if (state.filters.periodKind === 'month') months.add(monthKeyOf(state.filters.anchor.y, state.filters.anchor.m));
    return [...months];
}

async function loadTabData(tabId, { force = false } = {}) {
    const need = [];
    if (['marketing', 'staff-dashboard', 'overview'].includes(tabId)) {
        if (force || !state.data.mkStaff) need.push(loadMarketing().catch(lazyCatch('marketing', loadMarketing)));
    }
    if (['marketing', 'customers', 'staff-dashboard'].includes(tabId)) {
        if (force || !state.data.retention) need.push(loadRetention().catch(lazyCatch('retention', loadRetention)));
    }
    if (tabId === 'customers') {
        if (force || !state.data.ageDist) need.push(loadAgeDist().catch(lazyCatch('agedist', loadAgeDist)));
    }
    // 手入力データ: サマリー/マイダッシュボードは今月+対象月、インセンティブ/マーケ/入金突合は対象月
    if (['overview', 'staff-dashboard'].includes(tabId)) {
        for (const m of manualMonths()) need.push(loadManual(m).catch(e => console.warn('manual load', e)));
    }
    if (['incentive', 'marketing', 'recon'].includes(tabId)) {
        need.push(loadManual(monthKeyOf(state.filters.anchor.y, state.filters.anchor.m)).catch(e => console.warn('manual load', e)));
    }
    await Promise.all(need);
}

// ---- フィルタ変更 ----
function bindFilterEvents() {
    document.getElementById('store-selector')?.addEventListener('change', ev => {
        state.filters.shopId = ev.target.value === 'all' ? 'all' : Number(ev.target.value);
        state.filters.staffId = 'all';
        populateStaffSelector();
        renderNav();
        emit('filters');
        onFiltersChanged();
    });
    document.getElementById('staff-selector')?.addEventListener('change', ev => {
        state.filters.staffId = ev.target.value === 'all' ? 'all' : Number(ev.target.value);
        renderNav();
        if (isStaffLocked()) {
            // 自分 → マイダッシュボード、店舗全体 → サマリー
            switchTab(state.filters.staffId === 'all' ? 'overview' : 'staff-dashboard');
        } else if (state.filters.staffId !== 'all') {
            // スタッフを選んだらマイダッシュボードへ誘導（作業用タブを開いている間はそのまま）
            if (!['input', 'shift', 'recon', 'goal', 'settings', 'incentive'].includes(state.ui.activeTab)) switchTab('staff-dashboard');
        } else if (state.ui.activeTab === 'staff-dashboard') {
            switchTab('overview');
        }
        emit('filters');
        emit('data:core');
        emit('data:marketing');
        emit('data:manual');
        emit('data:goals');
    });
    document.querySelectorAll('.period-btn[data-period]').forEach(btn => {
        btn.addEventListener('click', () => {
            state.filters.periodKind = btn.dataset.period;
            document.querySelectorAll('.period-btn[data-period]').forEach(b => {
                const active = b === btn;
                b.classList.toggle('active', active);
                b.classList.toggle('text-surface-600', !active);
            });
            onFiltersChanged();
        });
    });
    document.getElementById('date-selector')?.addEventListener('change', ev => {
        const [y, m] = ev.target.value.split('-').map(Number);
        state.filters.anchor = { y, m };
        onFiltersChanged();
    });
    document.getElementById('refresh-data-btn')?.addEventListener('click', () => {
        clearApiCache();
        refreshAll();
    });
    document.getElementById('dark-mode-toggle')?.addEventListener('click', toggleDarkMode);
}

function onFiltersChanged() {
    state.data.channels = null;
    state.data.mkStaff = null;
    state.data.retention = null;
    emit('filters');
    refreshAll();
}

// ---- 自動更新（LIVE）----
function startLiveRefresh() {
    setInterval(() => {
        if (document.visibilityState === 'visible') refreshAll({ silent: true });
    }, REFRESH_INTERVAL);
    document.addEventListener('visibilitychange', () => {
        if (document.visibilityState === 'visible' && state.ui.lastFetchAt && Date.now() - state.ui.lastFetchAt > REFRESH_INTERVAL) {
            refreshAll({ silent: true });
        }
    });
    // オフライン/オンラインの切り替えを知らせる
    window.addEventListener('offline', () => { setLiveIndicator('error'); toast('オフラインです。接続が戻ると自動で更新します', 'warn'); });
    window.addEventListener('online', () => refreshAll({ silent: true }));
}

// ---- スプラッシュ ----
function removeSplash() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.classList.add('opacity-0');
        setTimeout(() => overlay.remove(), 450);
    }
    document.body.classList.remove('is-booting');
}

function setSplashText(text) {
    const el = document.getElementById('loading-text');
    if (el) el.textContent = text;
    const bar = document.getElementById('splash-progress-bar');
    if (bar) bar.style.width = `${Math.min((parseFloat(bar.style.width) || 10) + 22, 90)}%`;
}

// ---- 起動 ----
async function boot() {
    initDarkMode();
    initChartDefaults();

    setSplashText('認証を確認しています…');
    await ensureAuthenticated();

    setSplashText('マスタデータを取得しています…');
    try {
        await loadMasters();
    } catch (e) {
        console.error('マスタ取得失敗', e);
        toast('店舗情報の取得に失敗しました', 'error');
    }

    applyRoleUi();
    populateSelectors();
    bindFilterEvents();
    for (const v of VIEWS) {
        try { v.init(); } catch (e) { console.error('view init error', e); }
    }
    initNav();
    on('tab:shown', id => loadTabData(id));

    setSplashText('目標を読み込んでいます…');
    await loadGoals().catch(e => console.warn('goals load', e));

    setSplashText('データを読み込んでいます…');
    await refreshAll();
    removeSplash();
    startLiveRefresh();

    // メタ情報（デモ/AI可否）
    apiGet('meta').then(meta => {
        state.aiAvailable = !!meta.aiAvailable;
        state.demo = !!meta.demo;
        state.meta = meta;
        if (state.demo) document.getElementById('demo-badge')?.classList.remove('hidden');
        emit('meta');
    }).catch(() => {});
}

// PWA: Service Worker登録 + 新バージョン検知
function registerServiceWorker() {
    if (!('serviceWorker' in navigator) || location.protocol !== 'https:') return;
    navigator.serviceWorker.register('./sw.js').then(reg => {
        const notify = () => toast('新しいバージョンがあります', 'info', {
            duration: 0,
            action: { label: '再読み込み', onClick: () => location.reload() },
        });
        // 既に新しいSWが待機中
        if (reg.waiting && navigator.serviceWorker.controller) notify();
        reg.addEventListener('updatefound', () => {
            const nw = reg.installing;
            nw?.addEventListener('statechange', () => {
                if (nw.state === 'installed' && navigator.serviceWorker.controller) notify();
            });
        });
        // 1時間ごとに更新を確認（開きっぱなしの店内端末向け）
        setInterval(() => reg.update().catch(() => {}), 60 * 60 * 1000);
    }).catch(e => console.warn('SW登録失敗', e));
}

document.addEventListener('DOMContentLoaded', () => {
    registerServiceWorker();
    boot().catch(e => {
        console.error('boot error', e);
    });
});
