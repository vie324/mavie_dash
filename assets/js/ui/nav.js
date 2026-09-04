// タブナビゲーション（デスクトップサイドバー / モバイル横タブ / ボトムナビ+シート）
// 役割ごとの表示タブを一元定義する。

import { state, emit, isStaffLocked, isStoreLocked } from '../core/state.js';

// 役割別に見えるタブ（従来ツールの権限モデルを踏襲）
// 役割別の表示タブ（インセンティブ=給与はオーナーのみ、設定はオーナーのみ）
export const TABS = [
    { id: 'overview',        label: '店舗サマリー',     icon: 'layout-dashboard', roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'staff-dashboard', label: 'マイダッシュボード', icon: 'user',            roles: ['staff', 'admin-staff-selected'] },
    { id: 'input',           label: '日報入力',        icon: 'notebook-pen',     roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'sales',           label: '売上詳細',        icon: 'receipt',          roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'recon',           label: '入金突合',        icon: 'scale',            roles: ['admin', 'manager', 'store'] },
    { id: 'shift',           label: 'シフト',          icon: 'calendar-clock',   roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'marketing',       label: 'マーケティング',   icon: 'megaphone',        roles: ['admin', 'manager', 'store'] },
    { id: 'customers',       label: '顧客分析',        icon: 'pie-chart',        roles: ['admin', 'manager', 'store'] },
    { id: 'calendar',        label: 'カレンダー',      icon: 'calendar',         roles: ['admin', 'manager', 'store', 'staff'] },
    { id: 'incentive',       label: 'インセンティブ',   icon: 'coins',            roles: ['admin'] },
    { id: 'goal',            label: '売上目標設定',     icon: 'target',           roles: ['admin', 'manager', 'store'] },
    { id: 'settings',        label: '設定',           icon: 'settings',         roles: ['admin'] },
];

export const SIDEBAR_GROUPS = [
    { label: '経営', tabs: ['overview', 'sales', 'recon', 'calendar', 'incentive', 'goal'] },
    { label: '現場', tabs: ['staff-dashboard', 'input', 'shift'] },
    { label: 'マーケ', tabs: ['marketing', 'customers'] },
    { label: '管理', tabs: ['settings'] },
];

// ボトムナビに常時出す4つ + その他
const BOTTOM_MAIN = ['overview', 'staff-dashboard', 'input', 'sales'];

// 「作業用」タブ: 本日スナップショット・KPIカードを隠して入力画面をすぐ表示する
export const TOOL_TABS = new Set(['input', 'shift', 'recon', 'settings', 'goal', 'incentive']);

function role() {
    return state.session?.role || 'admin';
}

export function visibleTabs() {
    const r = role();
    return TABS.filter(t => {
        if (t.roles.includes(r)) {
            // スタッフが「店舗全体」を選んでいる間はマイダッシュボードを隠す（サマリーに切替）
            if (t.id === 'staff-dashboard' && r === 'staff' && state.filters.staffId === 'all') return false;
            return true;
        }
        // 管理者/店舗ビューでスタッフを選択中はマイダッシュボードも見せる
        if (t.roles.includes('admin-staff-selected') && r !== 'staff' && state.filters.staffId !== 'all') return true;
        return false;
    });
}

export function defaultTab() {
    if (isStaffLocked()) return 'staff-dashboard';
    return 'overview';
}

export function renderNav() {
    const tabs = visibleTabs();
    const ids = new Set(tabs.map(t => t.id));

    // サイドバー
    const sidebar = document.getElementById('sidebar-nav');
    if (sidebar) {
        sidebar.innerHTML = SIDEBAR_GROUPS.map(g => {
            const items = g.tabs.filter(id => ids.has(id));
            if (!items.length) return '';
            return `<p class="sidebar-group-label">${g.label}</p>` + items.map(id => {
                const t = TABS.find(x => x.id === id);
                return `<button class="sidebar-item${state.ui.activeTab === id ? ' active' : ''}" data-tab="${id}">
                    <i data-lucide="${t.icon}"></i><span>${t.label}</span></button>`;
            }).join('');
        }).join('');
    }

    // モバイル横タブ
    const mainTabs = document.getElementById('main-tabs');
    if (mainTabs) {
        mainTabs.innerHTML = tabs.map(t => `
            <button data-tab="${t.id}" class="tab-btn${state.ui.activeTab === t.id ? ' active' : ''} py-2 px-3 md:py-3 md:px-4 font-medium text-sm ${state.ui.activeTab === t.id ? 'text-accent-800' : 'text-surface-500'} flex items-center justify-center gap-2" title="${t.label}">
                <i data-lucide="${t.icon}" class="w-6 h-6 md:w-4 md:h-4"></i>
                <span class="hidden md:inline">${t.label}</span>
            </button>`).join('');
    }

    // ボトムナビ
    const bottomNav = document.getElementById('bottom-nav-items');
    if (bottomNav) {
        const main = BOTTOM_MAIN.filter(id => ids.has(id)).slice(0, 4);
        const rest = tabs.filter(t => !main.includes(t.id));
        bottomNav.innerHTML = main.map(id => {
            const t = TABS.find(x => x.id === id);
            return `<button data-tab="${id}" class="bottom-nav-item${state.ui.activeTab === id ? ' active' : ''}">
                <i data-lucide="${t.icon}"></i><span>${t.label}</span></button>`;
        }).join('') + (rest.length ? `
            <button id="bottom-nav-more" class="bottom-nav-item"><i data-lucide="more-horizontal"></i><span>その他</span></button>` : '');

        const sheet = document.getElementById('more-sheet-items');
        if (sheet) {
            sheet.innerHTML = rest.map(t => `
                <button data-tab="${t.id}" class="more-sheet-item">
                    <i data-lucide="${t.icon}"></i><span>${t.label}</span></button>`).join('');
        }
    }

    if (window.lucide) lucide.createIcons();
}

export function switchTab(id) {
    if (!visibleTabs().some(t => t.id === id)) return;
    state.ui.activeTab = id;
    document.body.classList.toggle('tab-tool', TOOL_TABS.has(id));
    for (const section of document.querySelectorAll('.tab-content')) {
        section.classList.toggle('hidden', section.id !== `content-${id}`);
    }
    for (const btn of document.querySelectorAll('[data-tab]')) {
        const active = btn.dataset.tab === id;
        btn.classList.toggle('active', active);
        if (btn.classList.contains('tab-btn')) {
            btn.classList.toggle('text-accent-800', active);
            btn.classList.toggle('text-surface-500', !active);
        }
    }
    closeMoreSheet();
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
        if (tabBtn) { switchTab(tabBtn.dataset.tab); return; }
        if (ev.target.closest('#more-sheet-backdrop') || ev.target.closest('#more-sheet-close')) {
            closeMoreSheet();
        }
    });

    // サマリー/マイダッシュボードなどの本体セクションの初期表示
    renderNav();
    const start = defaultTab();
    switchTab(start);

    // スタッフ選択の変化でナビを組み直す（マイダッシュボードタブの出現）
    return { renderNav };
}
