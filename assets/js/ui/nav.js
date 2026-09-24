// タブナビゲーション（デスクトップサイドバー / モバイル横タブ / ボトムナビ+シート）
// 役割ごとの表示タブを一元定義する。

import { state, emit, isStaffLocked } from '../core/state.js';

// 役割別の表示タブ（インセンティブ=給与はオーナーのみ、設定はオーナーのみ）
// short はスマホ下部ナビ用の短いラベル
export const TABS = [
    { id: 'home',            label: 'ホーム',           short: 'ホーム',   icon: 'house',            roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'overview',        label: '店舗サマリー',     short: 'サマリー', icon: 'layout-dashboard', roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'staff-dashboard', label: 'マイ成績',         short: 'マイ成績', icon: 'user-round',       roles: ['staff', 'admin-staff-selected'] },
    { id: 'input',           label: '日報入力',        short: '日報',     icon: 'notebook-pen',     roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'sales',           label: '売上詳細',        short: '売上',     icon: 'receipt',          roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'recon',           label: '入金突合',        short: '入金',     icon: 'scale',            roles: ['admin', 'manager', 'store'] },
    { id: 'cashbook',        label: '出納帳',          short: '出納帳',   icon: 'wallet',           roles: ['admin', 'manager', 'store'] },
    { id: 'shift',           label: 'シフト',          short: 'シフト',   icon: 'calendar-clock',   roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'marketing',       label: 'マーケティング',   short: 'マーケ',   icon: 'megaphone',        roles: ['admin', 'manager', 'store'] },
    { id: 'customers',       label: '顧客分析',        short: '顧客',     icon: 'pie-chart',        roles: ['admin', 'manager', 'store'] },
    { id: 'calendar',        label: 'カレンダー',      short: '暦',       icon: 'calendar',         roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'incentive',       label: 'インセンティブ',   short: '歩合',     icon: 'coins',            roles: ['admin'] },
    { id: 'goal',            label: '売上目標設定',     short: '目標',     icon: 'target',           roles: ['admin', 'manager', 'store'] },
    { id: 'guide',           label: '使い方・業務分担', short: '使い方',   icon: 'book-open',        roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'settings',        label: '設定',           short: '設定',     icon: 'settings',         roles: ['admin'] },
];

export const SIDEBAR_GROUPS = [
    { label: 'きょう', tabs: ['home'] },
    { label: '経営', tabs: ['overview', 'sales', 'recon', 'cashbook', 'calendar', 'incentive', 'goal'] },
    { label: '現場', tabs: ['staff-dashboard', 'input', 'shift'] },
    { label: 'マーケ', tabs: ['marketing', 'customers'] },
    { label: '管理', tabs: ['guide', 'settings'] },
];

// スマホ下部ナビに常時出す4つ（役割ごとに「毎日使うもの」を優先）+ その他
const BOTTOM_BY_ROLE = {
    staff: ['home', 'input', 'staff-dashboard', 'overview'],
    store: ['home', 'overview', 'cashbook', 'recon'],
    manager: ['home', 'overview', 'sales', 'marketing'],
    admin: ['home', 'overview', 'sales', 'marketing'],
};

// 「その他」シートのグループ
const MORE_GROUPS = [
    { label: '毎日の作業', tabs: ['input', 'cashbook', 'recon', 'shift'] },
    { label: '数字を見る', tabs: ['overview', 'staff-dashboard', 'sales', 'calendar', 'marketing', 'customers'] },
    { label: '計画・管理', tabs: ['goal', 'incentive', 'guide', 'settings'] },
];

// 「作業用」タブ: KPIカードを隠して入力画面をすぐ表示する
export const TOOL_TABS = new Set(['input', 'shift', 'recon', 'cashbook', 'settings', 'goal', 'incentive', 'guide', 'home']);

function role() {
    return state.session?.role || 'admin';
}

export function tabLabel(id) {
    const t = TABS.find(x => x.id === id);
    if (!t) return '';
    // スタッフにとってのサマリーは「店舗の成績」
    if (id === 'overview' && role() === 'staff') return '店舗の成績';
    return t.label;
}

function shortLabel(id) {
    const t = TABS.find(x => x.id === id);
    if (id === 'overview' && role() === 'staff') return '店舗';
    return t?.short || t?.label || '';
}

export function visibleTabs() {
    const r = role();
    return TABS.filter(t => {
        if (t.roles.includes(r)) return true;
        // 管理者/店舗ビューでスタッフを選択中はマイ成績も見せる
        if (t.roles.includes('admin-staff-selected') && r !== 'staff' && state.filters.staffId !== 'all') return true;
        return false;
    });
}

export function defaultTab() {
    return 'home';
}

// 未完了のやること件数（ホームのバッジ）
let homeBadge = 0;
export function setHomeBadge(n) {
    homeBadge = n;
    for (const el of document.querySelectorAll('[data-tab="home"] .nav-badge')) {
        el.textContent = n > 9 ? '9+' : String(n);
        el.classList.toggle('hidden', !n);
    }
}

function badgeHtml(id) {
    if (id !== 'home') return '';
    return `<span class="nav-badge${homeBadge ? '' : ' hidden'}">${homeBadge > 9 ? '9+' : homeBadge}</span>`;
}

export function renderNav() {
    const tabs = visibleTabs();
    const ids = new Set(tabs.map(t => t.id));
    const active = state.ui.activeTab;

    // サイドバー
    const sidebar = document.getElementById('sidebar-nav');
    if (sidebar) {
        sidebar.innerHTML = SIDEBAR_GROUPS.map(g => {
            const items = g.tabs.filter(id => ids.has(id));
            if (!items.length) return '';
            return `<p class="sidebar-group-label">${g.label}</p>` + items.map(id => {
                const t = TABS.find(x => x.id === id);
                return `<button class="sidebar-item${active === id ? ' active' : ''}" data-tab="${id}">
                    <i data-lucide="${t.icon}"></i><span>${tabLabel(id)}</span>${badgeHtml(id)}</button>`;
            }).join('');
        }).join('');
    }

    // タブレット横タブ（768〜1023px）
    const mainTabs = document.getElementById('main-tabs');
    if (mainTabs) {
        mainTabs.innerHTML = tabs.map(t => `
            <button data-tab="${t.id}" class="tab-btn${active === t.id ? ' active' : ''} py-2 px-3 md:py-3 md:px-4 font-medium text-sm ${active === t.id ? 'text-accent-800' : 'text-surface-500'} flex items-center justify-center gap-2" title="${tabLabel(t.id)}">
                <i data-lucide="${t.icon}" class="w-6 h-6 md:w-4 md:h-4"></i>
                <span class="hidden md:inline">${tabLabel(t.id)}</span>
            </button>`).join('');
    }

    // ボトムナビ
    const bottomNav = document.getElementById('bottom-nav-items');
    if (bottomNav) {
        const main = (BOTTOM_BY_ROLE[role()] || BOTTOM_BY_ROLE.admin).filter(id => ids.has(id)).slice(0, 4);
        const rest = tabs.filter(t => !main.includes(t.id));
        bottomNav.innerHTML = main.map(id => {
            const t = TABS.find(x => x.id === id);
            return `<button data-tab="${id}" class="bottom-nav-item${active === id ? ' active' : ''}" aria-label="${tabLabel(id)}">
                <span class="bottom-nav-icon"><i data-lucide="${t.icon}"></i>${badgeHtml(id)}</span><span>${shortLabel(id)}</span></button>`;
        }).join('') + (rest.length ? `
            <button id="bottom-nav-more" class="bottom-nav-item${rest.some(t => t.id === active) ? ' active' : ''}" aria-label="その他のメニュー">
                <span class="bottom-nav-icon"><i data-lucide="layout-grid"></i></span><span>その他</span></button>` : '');

        const sheet = document.getElementById('more-sheet-items');
        if (sheet) {
            const restIds = new Set(rest.map(t => t.id));
            sheet.innerHTML = MORE_GROUPS.map(g => {
                const items = g.tabs.filter(id => restIds.has(id));
                if (!items.length) return '';
                return `<p class="more-sheet-group-title">${g.label}</p><div class="more-sheet-grid">` + items.map(id => {
                    const t = TABS.find(x => x.id === id);
                    return `<button data-tab="${id}" class="more-sheet-item${active === id ? ' active' : ''}">
                        <i data-lucide="${t.icon}"></i><span>${tabLabel(id)}</span></button>`;
                }).join('') + '</div>';
            }).join('');
        }
    }

    if (window.lucide) lucide.createIcons();
}

// スタッフは「マイ成績/ホーム = 自分」「店舗の成績 = 店舗全体」をタブで切り替える
// （ヘッダーの切替セレクタとも同期）
function syncStaffScope(tabId) {
    if (!isStaffLocked()) return false;
    const want = tabId === 'overview' ? 'all' : (['home', 'staff-dashboard'].includes(tabId) ? state.session.staffId : null);
    if (want === null || String(state.filters.staffId) === String(want)) return false;
    state.filters.staffId = want;
    const sel = document.getElementById('staff-selector');
    if (sel) sel.value = String(want);
    return true;
}

export function switchTab(id) {
    if (!visibleTabs().some(t => t.id === id)) return;
    const scopeChanged = syncStaffScope(id);
    state.ui.activeTab = id;
    document.body.dataset.activeTab = id; // body自体に data-tab を付けるとクリック判定に巻き込まれるため別名
    document.body.classList.toggle('tab-tool', TOOL_TABS.has(id));
    for (const section of document.querySelectorAll('.tab-content')) {
        section.classList.toggle('hidden', section.id !== `content-${id}`);
    }
    for (const btn of document.querySelectorAll('[data-tab]')) {
        const on = btn.dataset.tab === id;
        btn.classList.toggle('active', on);
        if (btn.classList.contains('tab-btn')) {
            btn.classList.toggle('text-accent-800', on);
            btn.classList.toggle('text-surface-500', !on);
        }
    }
    document.getElementById('bottom-nav-more')?.classList.toggle('active',
        !document.querySelector(`.bottom-nav-item[data-tab="${id}"]`));
    closeMoreSheet();
    if (scopeChanged) {
        // 表示スコープ（自分/店舗）が変わったので各ビューを描き直す
        emit('filters');
        for (const ev of ['data:core', 'data:marketing', 'data:manual', 'data:goals']) emit(ev);
    }
    emit('tab:shown', id);
    window.scrollTo({ top: 0, behavior: 'instant' });
}

function setMoreSheet(open) {
    document.getElementById('more-sheet')?.classList.toggle('open', open);
    document.getElementById('more-sheet-backdrop')?.classList.toggle('open', open);
    document.getElementById('more-sheet')?.setAttribute('aria-hidden', String(!open));
}

function closeMoreSheet() { setMoreSheet(false); }

export function initNav() {
    document.addEventListener('click', ev => {
        if (ev.target.closest('#bottom-nav-more')) { setMoreSheet(true); return; }
        const tabBtn = ev.target.closest('[data-tab]');
        if (tabBtn && !tabBtn.closest('#filter-panel')) { switchTab(tabBtn.dataset.tab); return; }
        if (ev.target.closest('#more-sheet-backdrop') || ev.target.closest('#more-sheet-close')) {
            closeMoreSheet();
        }
    });
    document.addEventListener('keydown', ev => {
        if (ev.key === 'Escape') closeMoreSheet();
    });

    renderNav();
    switchTab(defaultTab());
    return { renderNav };
}
