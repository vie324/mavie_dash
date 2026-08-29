// 設定タブ（管理者専用）: 連携状態・スタッフ専用URL発行・数値の定義

import { state, on } from '../core/state.js';
import { esc } from '../core/format.js';
import { apiGet } from '../core/api.js';
import { toast } from '../core/engage.js';

export function init() {
    on('masters', renderUrlSelectors);
    on('meta', renderStatus);
    document.getElementById('url-shop-selector')?.addEventListener('change', renderUrlStaffOptions);
    document.getElementById('url-generate-btn')?.addEventListener('click', generateUrl);
    document.getElementById('url-copy-btn')?.addEventListener('click', copyUrl);
    renderStatus();
}

async function renderStatus() {
    const wrap = document.getElementById('settings-status');
    if (!wrap) return;
    let meta = null;
    try { meta = await apiGet('meta'); } catch (_) { /* 下で未接続表示 */ }
    const row = (label, value, ok) => `
        <div class="flex items-center justify-between bg-surface-50 dark:bg-gray-700/40 rounded-lg px-4 py-3">
            <span class="text-surface-600 dark:text-gray-300">${label}</span>
            <span class="font-semibold ${ok === true ? 'text-sage-600' : ok === false ? 'text-rose-500' : 'text-accent-900 dark:text-gray-100'}">${value}</span>
        </div>`;
    if (!meta) {
        wrap.innerHTML = row('接続状態', 'サーバーに接続できません', false);
        return;
    }
    wrap.innerHTML = [
        row('接続状態', meta.demo ? 'デモモード（APIキー未設定）' : '接続済み', !meta.demo),
        row('ブランド', esc(meta.brand?.name || '—')),
        row('スキーマバージョン', esc(meta.schemaVersion || '—')),
        row('個人情報の取得', meta.piiIncluded === true ? '含む（キー設定）' : '含まない', meta.piiIncluded === true ? false : true),
        row('AIアドバイス', meta.aiAvailable ? '利用可能' : '未設定（GEMINI_API_KEY）', meta.aiAvailable ? true : undefined),
    ].join('');
}

function renderUrlSelectors() {
    const shopSel = document.getElementById('url-shop-selector');
    if (!shopSel) return;
    shopSel.innerHTML = state.masters.shops.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
    renderUrlStaffOptions();
}

function renderUrlStaffOptions() {
    const shopSel = document.getElementById('url-shop-selector');
    const staffSel = document.getElementById('url-staff-selector');
    if (!shopSel || !staffSel) return;
    const staffs = state.masters.staffs.filter(s => String(s.shop_id) === String(shopSel.value));
    staffSel.innerHTML = '<option value="">店舗ビュー（スタッフ指定なし）</option>' +
        staffs.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
}

function generateUrl() {
    const shopSel = document.getElementById('url-shop-selector');
    const staffSel = document.getElementById('url-staff-selector');
    const wrap = document.getElementById('url-output-wrap');
    const output = document.getElementById('url-output');
    if (!shopSel || !output) return;
    const params = new URLSearchParams();
    params.set('store', shopSel.value);
    if (staffSel?.value) params.set('staff', staffSel.value);
    output.value = `${location.origin}${location.pathname}?${params.toString()}`;
    wrap?.classList.remove('hidden');
}

async function copyUrl() {
    const output = document.getElementById('url-output');
    if (!output?.value) return;
    try {
        await navigator.clipboard.writeText(output.value);
        toast('URLをコピーしました');
    } catch (_) {
        output.select();
        document.execCommand('copy');
        toast('URLをコピーしました');
    }
}
