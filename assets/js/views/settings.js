// 設定タブ（管理者専用）: 連携状態・スタッフ専用URL発行・数値の定義

import { state, on } from '../core/state.js';
import { esc, todayJst } from '../core/format.js';
import { apiGet } from '../core/api.js';
import { toast } from '../core/engage.js';
import { loadShift, saveShiftConfig } from '../data/shift.js';
import { salesBasis, setSalesBasis } from '../data/salonone.js';
import { emit } from '../core/state.js';

export function init() {
    on('masters', renderUrlSelectors);
    on('meta', renderStatus);
    document.getElementById('url-shop-selector')?.addEventListener('change', renderUrlStaffOptions);
    document.getElementById('url-role-selector')?.addEventListener('change', updateUrlRoleUi);
    document.getElementById('url-generate-btn')?.addEventListener('click', generateUrl);
    document.getElementById('url-copy-btn')?.addEventListener('click', copyUrl);
    document.getElementById('shift-cfg-save')?.addEventListener('click', saveShiftRules);
    const basisSel = document.getElementById('sales-basis-selector');
    if (basisSel) {
        basisSel.value = salesBasis();
        basisSel.addEventListener('change', () => {
            setSalesBasis(basisSel.value);
            toast('売上の基準を変更しました');
            emit('data:core');
            emit('data:marketing');
            emit('data:manual');
        });
    }
    on('tab:shown', id => { if (id === 'settings') { loadShiftRules(); loadAccounts(); } });
    on('masters', renderAccounts);
    document.getElementById('accounts-body')?.addEventListener('click', onAccountsClick);
    updateUrlRoleUi();
    renderStatus();
}

// ---- スタッフ/店長アカウント（画面から発行） ----
let accountsState = { storage: null, accounts: {} };

async function accountsRequest(options) {
    const res = await fetch('/api/accounts', { credentials: 'same-origin', cache: 'no-store', ...options });
    let json = {};
    try { json = await res.json(); } catch (_) { /* 空 */ }
    if (!res.ok) throw Object.assign(new Error(json.error || 'unknown'), { status: res.status, body: json });
    return json;
}

async function loadAccounts() {
    try {
        accountsState = await accountsRequest({});
    } catch (e) {
        console.warn('accounts load', e);
        accountsState = { storage: 'error', accounts: {} };
    }
    renderAccounts();
}

function accountUrl(kind, shopId, staffId) {
    const params = new URLSearchParams();
    params.set('store', shopId);
    if (kind === 'staff') params.set('staff', staffId);
    return `${location.origin}${location.pathname}?${params.toString()}`;
}

function renderAccounts() {
    const body = document.getElementById('accounts-body');
    if (!body) return;
    const usable = accountsState.storage === 'kv';
    document.getElementById('accounts-storage-notice')?.classList.toggle('hidden', usable);
    const accounts = accountsState.accounts || {};
    const rows = [];
    for (const shop of state.masters.shops) {
        const storeKey = `store:${shop.id}`;
        rows.push(accountRow('store', shop.id, shop.id, `<span class="font-semibold">${esc(shop.name)}</span> <span class="text-[10px] text-surface-400">店長</span>`, accounts[storeKey], usable));
        for (const st of state.masters.staffs.filter(s => String(s.shop_id) === String(shop.id))) {
            rows.push(accountRow('staff', st.id, shop.id, `<span class="text-surface-400 mr-1">└</span>${esc(st.name)}`, accounts[`staff:${st.id}`], usable));
        }
    }
    body.innerHTML = rows.join('') || '<tr><td colspan="4" class="py-6 text-center text-surface-500">店舗・スタッフ情報がありません</td></tr>';
}

function accountRow(kind, id, shopId, label, acc, usable) {
    const set = !!acc;
    return `
    <tr class="border-b border-surface-100 dark:border-accent-800" data-kind="${kind}" data-id="${id}" data-shop="${shopId}">
        <td class="py-2 px-3">${label}</td>
        <td class="py-2 px-3">${set ? '<span class="text-sage-600 font-semibold">設定済み</span>' : '<span class="text-surface-400">未設定（URLのみで閲覧可）</span>'}</td>
        <td class="py-2 px-3">
            <input type="password" autocomplete="new-password" placeholder="${set ? '変更する場合は入力' : '4文字以上'}" ${usable ? '' : 'disabled'}
                class="acc-pass w-40 px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white disabled:opacity-50">
        </td>
        <td class="py-2 px-3 text-right whitespace-nowrap">
            <button data-acc-action="set" class="btn-primary py-1 px-3 text-xs" ${usable ? '' : 'disabled'}>${set ? '変更' : '発行'}</button>
            ${set ? '<button data-acc-action="delete" class="btn-secondary py-1 px-3 text-xs ml-1">解除</button>' : ''}
            <button data-acc-action="copy" class="btn-secondary py-1 px-3 text-xs ml-1">URLコピー</button>
        </td>
    </tr>`;
}

async function onAccountsClick(ev) {
    const btn = ev.target.closest('button[data-acc-action]');
    if (!btn) return;
    const tr = btn.closest('tr');
    const { kind, id, shop } = tr.dataset;
    const action = btn.dataset.accAction;
    if (action === 'copy') {
        const url = accountUrl(kind, shop, id);
        try { await navigator.clipboard.writeText(url); toast('URLをコピーしました'); }
        catch (_) { toast(url); }
        return;
    }
    btn.disabled = true;
    try {
        if (action === 'set') {
            const password = tr.querySelector('.acc-pass')?.value || '';
            if (password.length < 4) { toast('パスワードは4文字以上で入力してください', 'warn'); return; }
            const res = await accountsRequest({ method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'set', kind, id, password }) });
            accountsState.accounts = res.accounts;
            toast('パスワードを設定しました。URLと一緒にスタッフへお渡しください');
        } else if (action === 'delete') {
            const res = await accountsRequest({ method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'delete', kind, id }) });
            accountsState.accounts = res.accounts;
            toast('パスワードを解除しました');
        }
        renderAccounts();
    } catch (e) {
        console.error('accounts', e);
        toast(e.body?.detail || '操作に失敗しました', 'error');
    } finally {
        btn.disabled = false;
    }
}

// ---- シフトルール（設定から変更可能） ----
async function loadShiftRules() {
    try {
        const t = todayJst();
        const res = await loadShift(`${t.y}-${String(t.m).padStart(2, '0')}`);
        fillShiftRules(res.config);
    } catch (e) {
        console.warn('shift config load', e);
    }
}

function fillShiftRules(cfg) {
    if (!cfg) return;
    setValue('shift-cfg-offdays', cfg.offDays);
    setValue('shift-cfg-weekend', cfg.weekendOffDays);
    setValue('shift-cfg-sameday', cfg.maxSameDayOff);
    const cur = document.getElementById('shift-cfg-current');
    if (cur) cur.textContent = `月${cfg.offDays}日休み ・ 土日${cfg.weekendOffDays}日 ・ 同日${cfg.maxSameDayOff}人まで`;
}

async function saveShiftRules() {
    const btn = document.getElementById('shift-cfg-save');
    btn.disabled = true;
    try {
        const res = await saveShiftConfig({
            offDays: Number(document.getElementById('shift-cfg-offdays')?.value),
            weekendOffDays: Number(document.getElementById('shift-cfg-weekend')?.value),
            maxSameDayOff: Number(document.getElementById('shift-cfg-sameday')?.value),
        });
        fillShiftRules(res.config);
        toast('シフトルールを保存しました');
    } catch (e) {
        console.error('shift config save', e);
        toast(e?.body?.detail || 'ルールの保存に失敗しました（サーバー保存が必要です）', 'error');
    } finally {
        btn.disabled = false;
    }
}

function setValue(id, v) {
    const el = document.getElementById(id);
    if (el) el.value = v ?? '';
}

// 役割に応じて店舗/スタッフ選択の表示を切り替え
function updateUrlRoleUi() {
    const role = document.getElementById('url-role-selector')?.value || 'staff';
    document.getElementById('url-shop-wrap')?.classList.toggle('hidden', role === 'manager');
    document.getElementById('url-staff-wrap')?.classList.toggle('hidden', role !== 'staff');
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
    const rows = [
        row('接続状態', meta.demo ? 'デモモード（APIキー未設定）' : '接続済み', !meta.demo),
        row('ブランド', esc(meta.brand?.name || '—')),
        row('スキーマバージョン', esc(meta.schemaVersion || '—')),
        row('個人情報の取得', meta.piiIncluded === true ? '含む（キー設定）' : '含まない', meta.piiIncluded === true ? false : true),
        row('AIアドバイス', meta.aiAvailable ? '利用可能' : '未設定（GEMINI_API_KEY）', meta.aiAvailable ? true : undefined),
        row('サーバー保存', meta.manualStorage ? `${esc(meta.storage?.label || 'サーバー')}（全端末共有）` : 'この端末のみ（未設定: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY）', meta.manualStorage ? true : undefined),
    ];
    if (meta.storage?.warning) rows.push(row('サーバー保存の警告', '⚠ ' + esc(meta.storage.warning), false));
    // パスワード設定状況の警告（オーナーセッションのみ返る）
    if (meta.passwords) {
        const p = meta.passwords;
        rows.push(
            row('オーナーパスワード', p.admin ? '設定済み' : '⚠ 未設定（URLを知っていれば誰でも閲覧可）', p.admin ? true : false),
            row('マネージャーパスワード', p.manager ? '設定済み' : '未設定（MANAGER_PASSWORD）', p.manager ? true : undefined),
            row('店長パスワード', p.storeCount > 0 ? `${p.storeCount}件 設定済み` : '未設定（STORE_PASSWORDS）', p.storeCount > 0 ? true : undefined),
            row('スタッフパスワード', (p.staffCount > 0 || p.staffAccounts > 0) ? `${(p.staffAccounts || 0) + (p.staffCount || 0)}件 設定済み` : '未設定（下の「スタッフアカウントの発行」から設定）', (p.staffCount > 0 || p.staffAccounts > 0) ? true : undefined),
        );
    }
    wrap.innerHTML = rows.join('');
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
    const role = document.getElementById('url-role-selector')?.value || 'staff';
    const shopSel = document.getElementById('url-shop-selector');
    const staffSel = document.getElementById('url-staff-selector');
    const wrap = document.getElementById('url-output-wrap');
    const output = document.getElementById('url-output');
    if (!output) return;
    const params = new URLSearchParams();
    if (role === 'manager') {
        params.set('mode', 'manager');
    } else {
        if (!shopSel?.value) return;
        params.set('store', shopSel.value);
        if (role === 'staff' && staffSel?.value) params.set('staff', staffSel.value);
    }
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
