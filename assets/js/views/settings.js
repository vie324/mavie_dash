// 設定タブ（管理者専用）: 連携状態・スタッフ専用URL発行・数値の定義

import { state, on } from '../core/state.js';
import { esc, todayJst } from '../core/format.js';
import { apiGet } from '../core/api.js';
import { toast } from '../core/engage.js';
import { loadShift, saveShiftConfig } from '../data/shift.js';

export function init() {
    on('masters', renderUrlSelectors);
    on('meta', renderStatus);
    document.getElementById('url-shop-selector')?.addEventListener('change', renderUrlStaffOptions);
    document.getElementById('url-role-selector')?.addEventListener('change', updateUrlRoleUi);
    document.getElementById('url-generate-btn')?.addEventListener('click', generateUrl);
    document.getElementById('url-copy-btn')?.addEventListener('click', copyUrl);
    document.getElementById('shift-cfg-save')?.addEventListener('click', saveShiftRules);
    on('tab:shown', id => { if (id === 'settings') loadShiftRules(); });
    updateUrlRoleUi();
    renderStatus();
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
        row('手入力データの保存', meta.manualStorage ? 'サーバー保存（全端末共有）' : 'この端末のみ（Upstash未設定）', meta.manualStorage ? true : undefined),
    ];
    // パスワード設定状況の警告（オーナーセッションのみ返る）
    if (meta.passwords) {
        const p = meta.passwords;
        rows.push(
            row('オーナーパスワード', p.admin ? '設定済み' : '⚠ 未設定（URLを知っていれば誰でも閲覧可）', p.admin ? true : false),
            row('マネージャーパスワード', p.manager ? '設定済み' : '未設定（MANAGER_PASSWORD）', p.manager ? true : undefined),
            row('店長パスワード', p.storeCount > 0 ? `${p.storeCount}件 設定済み` : '未設定（STORE_PASSWORDS）', p.storeCount > 0 ? true : undefined),
            row('スタッフパスワード', p.staffCount > 0 ? `${p.staffCount}件 設定済み` : '未設定（STAFF_PASSWORDS）', p.staffCount > 0 ? true : undefined),
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
