// スマホ向けアプリシェル: 絞り込みシート / 表示条件サマリー / 引っ張って更新 / ホーム画面追加の案内
// PCではヘッダー内に横並びの絞り込み、スマホ（<768px）では下から出るシートに切り替える。

import { state, on, isStaffLocked, isStoreLocked, currentShopId, currentStaffId, shopName, staffName } from '../core/state.js';
import { todayJst, todayStr, dowJa } from '../core/format.js';
import { tabLabel } from './nav.js';

const mq = window.matchMedia('(max-width: 767px)');

// ---- 絞り込みシート ----
// ヘッダーは backdrop-filter を持つため、その中の position:fixed は画面基準にならない。
// スマホ時はパネル要素ごと body 直下へ移し、PC時はヘッダー内の元の位置に戻す。
let panel, panelHome, panelNext;

function placePanel() {
    if (!panel) return;
    if (mq.matches) {
        if (panel.parentElement !== document.body) document.body.appendChild(panel);
    } else if (panel.parentElement !== panelHome) {
        panelHome.insertBefore(panel, panelNext);
        setSheet(false);
    }
}

function setSheet(open) {
    if (!panel) return;
    const on = open && mq.matches;
    panel.classList.toggle('open', on);
    document.getElementById('filter-backdrop')?.classList.toggle('open', on);
    document.body.classList.toggle('sheet-open', on);
    panel.setAttribute('aria-modal', String(on));
    // セレクトに自動フォーカスするとiPhoneで選択ホイールが勝手に開くため、閉じるボタンに合わせる
    if (on) document.getElementById('filter-sheet-close')?.focus({ preventScroll: true });
}

export function closeFilterSheet() { setSheet(false); }

function periodLabel() {
    const { y, m } = state.filters.anchor;
    const kind = state.filters.periodKind;
    if (kind === 'month') return `${y}年${m}月`;
    const len = { '3months': '3ヶ月', '6months': '6ヶ月', year: '1年' }[kind] || '';
    return `${y}年${m}月まで${len}`;
}

function scopeLabel() {
    if (isStaffLocked()) {
        return currentStaffId() === 'all' ? `${shopName(currentShopId())} 全体` : `${state.session.staffName || '自分'}`;
    }
    const shop = shopName(currentShopId());
    const staff = currentStaffId() === 'all' ? '' : ` · ${staffName(currentStaffId())}`;
    return `${shop}${staff}`;
}

export function updateFilterSummary() {
    const tabEl = document.getElementById('filter-summary-tab');
    const scopeEl = document.getElementById('filter-summary-scope');
    if (!tabEl || !scopeEl) return;
    const tab = state.ui.activeTab;
    tabEl.textContent = tabLabel(tab) || '—';
    if (tab === 'home') {
        const t = todayJst();
        scopeEl.textContent = `今日 ${t.m}/${t.d}（${dowJa(todayStr())}）· ${scopeLabel()}`;
    } else if (['input', 'shift', 'recon', 'goal', 'settings', 'guide', 'incentive'].includes(tab)) {
        // 作業用タブは画面内で日付・月を選ぶため、ここでは対象範囲だけ示す
        scopeEl.textContent = scopeLabel();
    } else {
        scopeEl.textContent = `${scopeLabel()} · ${periodLabel()}`;
    }
    // スタッフ/店長は店舗が固定なので、切替シートの店舗欄は出さない
    document.getElementById('store-selector-wrap')?.classList.toggle('hidden', isStaffLocked() || isStoreLocked());
    const staffLabel = document.getElementById('staff-selector-label');
    if (staffLabel) staffLabel.textContent = isStaffLocked() ? '表示する成績' : 'スタッフ';
}

// ---- 引っ張って更新（ホーム画面から起動したPWAのみ。ブラウザには標準の更新があるため）----
function isStandalone() {
    return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
}

function initPullToRefresh(onRefresh) {
    if (!isStandalone()) return;
    const indicator = document.createElement('div');
    indicator.className = 'ptr-indicator';
    indicator.innerHTML = '<i data-lucide="arrow-down"></i>';
    document.body.appendChild(indicator);
    if (window.lucide) lucide.createIcons();

    const THRESHOLD = 72;
    let startY = null, dy = 0, busy = false;
    window.addEventListener('touchstart', ev => {
        if (busy || window.scrollY > 0 || document.body.classList.contains('sheet-open') || ev.touches.length !== 1) { startY = null; return; }
        startY = ev.touches[0].clientY;
        dy = 0;
    }, { passive: true });
    window.addEventListener('touchmove', ev => {
        if (startY === null) return;
        dy = Math.max(0, ev.touches[0].clientY - startY);
        if (dy <= 0) return;
        const pull = Math.min(dy, THRESHOLD * 1.6);
        indicator.style.transform = `translate(-50%, ${pull * 0.6}px) rotate(${Math.min(dy / THRESHOLD, 1) * 180}deg)`;
        indicator.classList.toggle('ready', dy >= THRESHOLD);
        indicator.classList.add('visible');
    }, { passive: true });
    window.addEventListener('touchend', async () => {
        if (startY === null) return;
        const fire = dy >= THRESHOLD;
        startY = null;
        if (!fire) { indicator.classList.remove('visible', 'ready'); indicator.style.transform = ''; return; }
        busy = true;
        indicator.classList.add('spinning');
        try { await onRefresh(); } finally {
            busy = false;
            indicator.classList.remove('visible', 'ready', 'spinning');
            indicator.style.transform = '';
        }
    });
}

// ---- iPhoneの「ホーム画面に追加」案内（Safariで開いていて未追加のときだけ）----
const INSTALL_KEY = 'vie_install_hint_dismissed';
export function shouldShowInstallHint() {
    const ua = navigator.userAgent || '';
    const ios = /iPhone|iPad|iPod/.test(ua);
    if (!ios || isStandalone()) return false;
    try { return localStorage.getItem(INSTALL_KEY) !== '1'; } catch (_) { return false; }
}
export function dismissInstallHint() {
    try { localStorage.setItem(INSTALL_KEY, '1'); } catch (_) { /* ignore */ }
}

export function initShell({ onRefresh }) {
    panel = document.getElementById('filter-panel');
    if (panel) {
        panelHome = panel.parentElement;
        panelNext = panel.nextElementSibling;
    }
    placePanel();
    mq.addEventListener?.('change', placePanel);

    document.addEventListener('click', ev => {
        if (ev.target.closest('#filter-open-btn') || ev.target.closest('#filter-summary')) { setSheet(true); return; }
        if (ev.target.closest('#filter-sheet-close') || ev.target.closest('#filter-sheet-done') || ev.target.closest('#filter-backdrop')) {
            setSheet(false);
        }
    });
    document.addEventListener('keydown', ev => { if (ev.key === 'Escape') setSheet(false); });

    on('filters', updateFilterSummary);
    on('tab:shown', updateFilterSummary);
    on('masters', updateFilterSummary);
    updateFilterSummary();

    initPullToRefresh(onRefresh);
}
