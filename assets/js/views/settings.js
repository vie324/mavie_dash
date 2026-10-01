// 設定タブ（管理者専用）: 連携状態・スタッフ専用URL発行・数値の定義

import { state, on, shopName } from '../core/state.js';
import { getInsights, REASON_LABELS } from '../data/insights.js';
import { monthlyTotalsByStaff } from '../data/manual.js';
import { esc, todayJst } from '../core/format.js';
import { apiGet } from '../core/api.js';
import { toast } from '../core/engage.js';
import { loadShift, saveShiftConfig } from '../data/shift.js';
import { salesBasis, setSalesBasis } from '../data/salonone.js';
import { emit } from '../core/state.js';

export function init() {
    on('masters', renderUrlSelectors);
    on('meta', renderStatus);
    on('data:insights', renderInsightsDiag);
    on('data:manual', renderInsightsDiag);
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
    document.getElementById('accounts-bulk-shop')?.addEventListener('change', renderBulkButton);
    document.getElementById('accounts-bulk-btn')?.addEventListener('click', bulkIssue);
    document.getElementById('accounts-issued-list')?.addEventListener('click', onIssuedClick);
    document.getElementById('accounts-issued-copyall')?.addEventListener('click', copyAllIssued);
    document.getElementById('accounts-issued-close')?.addEventListener('click', closeIssued);
    updateUrlRoleUi();
    renderStatus();
}

// ---- スタッフ/店長アカウント（画面から発行） ----
let accountsState = { storage: null, accounts: {} };
// この画面で発行したアカウント（平文のパスワードはここにだけ一時的に持つ。再読み込みで消える）
let issued = [];

// 自動発行のパスワード: 見間違えやすい文字（0/o/1/l/i）を除いた英小文字+数字 8文字（サーバーの一括発行と同じ規則）
const PASS_CHARS = 'abcdefghjkmnpqrstuvwxyz23456789';
function randomPassword() {
    const out = [];
    const buf = new Uint8Array(16);
    while (out.length < 8) {
        crypto.getRandomValues(buf);
        // 偏りが出ないよう 31の倍数（248）未満だけ使う
        for (const b of buf) if (b < 248 && out.length < 8) out.push(PASS_CHARS[b % PASS_CHARS.length]);
    }
    return out.join('');
}

async function copyText(text) {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch (_) {
        // http・古いブラウザ向けのフォールバック
        const ta = document.createElement('textarea');
        ta.value = text;
        ta.setAttribute('readonly', '');
        ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0;';
        document.body.appendChild(ta);
        ta.select();
        let ok = false;
        try { ok = document.execCommand('copy'); } catch (_) { ok = false; }
        ta.remove();
        return ok;
    }
}

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
    renderBulkShops();
    renderBulkButton();
}

function accountRow(kind, id, shopId, label, acc, usable) {
    const set = !!acc;
    return `
    <tr class="border-b border-surface-100 dark:border-accent-800" data-kind="${kind}" data-id="${id}" data-shop="${shopId}">
        <td class="py-2 px-3">${label}</td>
        <td class="py-2 px-3">${set ? '<span class="text-sage-600 font-semibold">設定済み</span>' : '<span class="text-surface-400">未設定（URLのみで閲覧可）</span>'}</td>
        <td class="py-2 px-3">
            <div class="flex items-center gap-1">
                <input type="password" autocomplete="new-password" placeholder="${set ? '変更する場合は入力' : '4文字以上'}" ${usable ? '' : 'disabled'}
                    class="acc-pass w-40 px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white disabled:opacity-50">
                <button type="button" data-acc-action="gen" class="btn-secondary py-1 px-2 text-xs" title="パスワードを自動で作る" ${usable ? '' : 'disabled'}>自動</button>
            </div>
        </td>
        <td class="py-2 px-3 text-right whitespace-nowrap">
            <button data-acc-action="set" class="btn-primary py-1 px-3 text-xs" ${usable ? '' : 'disabled'}>${set ? '再発行' : '発行'}</button>
            ${set ? '<button data-acc-action="delete" class="btn-secondary py-1 px-3 text-xs ml-1">解除</button>' : ''}
            <button data-acc-action="copy" class="btn-secondary py-1 px-3 text-xs ml-1">URLコピー</button>
        </td>
    </tr>`;
}

// ---- 一括発行 ----
function bulkTargets(shopId) {
    const accounts = accountsState.accounts || {};
    return state.masters.staffs.filter(s => (shopId === 'all' || String(s.shop_id) === String(shopId)) && !accounts[`staff:${s.id}`]);
}

function renderBulkShops() {
    const sel = document.getElementById('accounts-bulk-shop');
    if (!sel) return;
    const prev = sel.value || 'all';
    sel.innerHTML = '<option value="all">全店舗</option>' + state.masters.shops.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
    sel.value = [...sel.options].some(o => o.value === prev) ? prev : 'all';
}

function renderBulkButton() {
    const btn = document.getElementById('accounts-bulk-btn');
    if (!btn) return;
    const usable = accountsState.storage === 'kv';
    const n = bulkTargets(document.getElementById('accounts-bulk-shop')?.value || 'all').length;
    btn.textContent = n > 0 ? `${n}名に一括発行` : '未発行のスタッフはいません';
    btn.disabled = !usable || n === 0;
}

async function bulkIssue() {
    const btn = document.getElementById('accounts-bulk-btn');
    const shopId = document.getElementById('accounts-bulk-shop')?.value || 'all';
    const targets = bulkTargets(shopId);
    if (!targets.length) return;
    const names = targets.slice(0, 6).map(s => s.name).join('、') + (targets.length > 6 ? ` ほか${targets.length - 6}名` : '');
    if (!window.confirm(`${shopId === 'all' ? '全店舗' : shopName(shopId)}のアカウント未発行のスタッフ ${targets.length}名（${names}）にパスワードを自動で発行します。よろしいですか？`)) return;
    btn.disabled = true;
    try {
        const res = await accountsRequest({ method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'bulk', shopId }) });
        accountsState.accounts = res.accounts;
        addIssued(res.issued || []);
        toast(res.issued?.length ? `${res.issued.length}名に発行しました。LINEの文面をコピーして送ってください` : '発行するスタッフはいませんでした', res.issued?.length ? 'success' : 'info');
        renderAccounts();
    } catch (e) {
        console.error('accounts bulk', e);
        toast(e.body?.detail || '一括発行に失敗しました', 'error');
        renderBulkButton();
    }
}

// ---- 発行したアカウント（URL・パスワード・LINE用の文面）----
function addIssued(items) {
    for (const it of items) {
        issued = issued.filter(x => !(x.kind === it.kind && x.id === String(it.id)));
        const name = it.kind === 'store' ? `${shopName(it.shopId)} 店長` : (it.name || state.masters.staffs.find(s => String(s.id) === String(it.id))?.name || '');
        issued.push({ kind: it.kind, id: String(it.id), shopId: String(it.shopId), name, password: it.password, copied: false });
    }
    renderIssued();
    document.getElementById('accounts-issued')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function lineMessage(it) {
    const url = accountUrl(it.kind, it.shopId, it.id);
    if (it.kind === 'store') {
        return [
            `${shopName(it.shopId)}の店長用アカウントです。`,
            '店舗の数字・スタッフの日報・出納帳・入金突合・シフトの承認ができます。',
            '',
            '▼URL（開いたらホーム画面に追加しておくと便利です）',
            url,
            '',
            '▼パスワード',
            it.password,
            '',
            '※パスワードは他の人に教えないでください',
        ].join('\n');
    }
    return [
        `${it.name}さん、お疲れさまです！`,
        `vieダッシュボードのログイン情報です（${it.name}さん専用）。`,
        '',
        '▼URL（開いたらホーム画面に追加しておくと便利です）',
        url,
        '',
        '▼パスワード',
        it.password,
        '',
        '毎日、退勤前に「日報」から次回予約（媒体別・新規/2回目以降）を入力してください。',
        '※パスワードは他の人に教えないでください',
    ].join('\n');
}

function renderIssued() {
    const panel = document.getElementById('accounts-issued');
    const list = document.getElementById('accounts-issued-list');
    if (!panel || !list) return;
    panel.classList.toggle('hidden', issued.length === 0);
    const count = document.getElementById('accounts-issued-count');
    if (count) count.textContent = `${issued.length}件`;
    list.innerHTML = issued.map((it, i) => `
        <li class="issued-item${it.copied ? ' copied' : ''}">
            <div class="issued-who">
                <span class="issued-shop">${esc(shopName(it.shopId))}</span>
                <b>${esc(it.name)}</b>
            </div>
            <div class="issued-cred">
                <span class="issued-label">URL</span><code class="issued-url">${esc(accountUrl(it.kind, it.shopId, it.id))}</code>
                <span class="issued-label">パスワード</span><code class="issued-pass">${esc(it.password)}</code>
            </div>
            <button type="button" class="btn-primary py-1.5 px-3 text-xs" data-issued-copy="${i}">${it.copied ? '✓ コピー済み（もう一度）' : 'LINEの文面をコピー'}</button>
        </li>`).join('');
}

async function onIssuedClick(ev) {
    const btn = ev.target.closest('button[data-issued-copy]');
    if (!btn) return;
    const it = issued[Number(btn.dataset.issuedCopy)];
    if (!it) return;
    if (await copyText(lineMessage(it))) {
        it.copied = true;
        renderIssued();
        toast(`${it.name}さん用の文面をコピーしました。LINEに貼り付けて送ってください`, 'success');
    } else {
        toast('コピーできませんでした。URLとパスワードを長押しでコピーしてください', 'warn');
    }
}

async function copyAllIssued() {
    const text = issued.map(it => `${shopName(it.shopId)} ${it.name}\nURL: ${accountUrl(it.kind, it.shopId, it.id)}\nパスワード: ${it.password}`).join('\n\n');
    toast(await copyText(text) ? '一覧をコピーしました（パスワード入り。共有先に注意してください）' : 'コピーできませんでした', 'info');
}

function closeIssued() {
    const notCopied = issued.filter(it => !it.copied).length;
    if (notCopied && !window.confirm(`まだ文面をコピーしていないアカウントが${notCopied}件あります。閉じるとパスワードは二度と表示されません（再発行はできます）。閉じますか？`)) return;
    issued = [];
    renderIssued();
}

async function onAccountsClick(ev) {
    const btn = ev.target.closest('button[data-acc-action]');
    if (!btn) return;
    const tr = btn.closest('tr');
    const { kind, id, shop } = tr.dataset;
    const action = btn.dataset.accAction;
    if (action === 'copy') {
        const url = accountUrl(kind, shop, id);
        toast(await copyText(url) ? 'URLをコピーしました' : url);
        return;
    }
    if (action === 'gen') {
        // 自動で作ったパスワードは見えるようにしておく（発行後に文面へ入る）
        const input = tr.querySelector('.acc-pass');
        if (input) { input.type = 'text'; input.value = randomPassword(); input.focus(); }
        return;
    }
    btn.disabled = true;
    try {
        if (action === 'set') {
            const password = tr.querySelector('.acc-pass')?.value || '';
            if (password.length < 4) { toast('パスワードは4文字以上で入力してください（「自動」で作れます）', 'warn'); return; }
            const res = await accountsRequest({ method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'set', kind, id, password }) });
            accountsState.accounts = res.accounts;
            addIssued([{ kind, id, shopId: shop, password }]);
            toast('パスワードを設定しました。「LINEの文面をコピー」で本人に送ってください', 'success');
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
    setValue('shift-cfg-deadline', cfg.requestDeadline ?? 20);
    const cur = document.getElementById('shift-cfg-current');
    if (cur) cur.textContent = `月${cfg.offDays}日休み ・ 土日${cfg.weekendOffDays}日 ・ 同日${cfg.maxSameDayOff}人まで ・ ${cfg.requestDeadline ? `申請締切 毎月${cfg.requestDeadline}日` : '申請締切なし'}`;
}

async function saveShiftRules() {
    const btn = document.getElementById('shift-cfg-save');
    btn.disabled = true;
    try {
        const res = await saveShiftConfig({
            offDays: Number(document.getElementById('shift-cfg-offdays')?.value),
            weekendOffDays: Number(document.getElementById('shift-cfg-weekend')?.value),
            maxSameDayOff: Number(document.getElementById('shift-cfg-sameday')?.value),
            requestDeadline: Number(document.getElementById('shift-cfg-deadline')?.value || 0),
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
        row('日報・目標・シフトの保存', meta.manualStorage ? `${esc(meta.storage?.label || 'サーバー')}に保存（全端末で共有）` : '⚠ この端末のみ（未設定: SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY）', meta.manualStorage ? true : false),
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

// ---- 予約データ連携（β）の診断 ----
function renderInsightsDiag() {
    const el = document.getElementById('insights-diag');
    if (!el) return;
    const now = new Date(Date.now() + 9 * 3600 * 1000);
    const month = `${now.getUTCFullYear()}-${String(now.getUTCMonth() + 1).padStart(2, '0')}`;
    const ins = getInsights(month);
    if (!ins) { el.textContent = '読み込み中…（取得できない場合はSalonOneのAPIキーの権限を確認してください）'; return; }
    const d = ins.diagnostics || {};
    const ok = ins.reliable;
    const row = (label, value) => `<div class="flex justify-between gap-3 py-1.5 border-b border-surface-100 dark:border-accent-800"><span class="text-surface-500">${label}</span><span class="font-semibold text-right break-all">${value}</span></div>`;
    // 日報（手入力）との比較: 今月の合計
    const totals = monthlyTotalsByStaff(month);
    let manualNext = 0;
    for (const t of Object.values(totals)) manualNext += (t.nextNew || 0) + (t.nextRepeat || 0);
    const estRate = ins.total?.visits > 0 ? `${Math.round(ins.total.withNext / ins.total.visits * 100)}%（${ins.total.withNext} / ${ins.total.visits}名）` : '—';
    el.innerHTML = `
        <div class="mb-3">${ok
            ? '<span class="chip chip-sage">✓ 推定に使えます</span>'
            : `<span class="chip chip-rose">推定に使えません</span> <span class="text-xs text-surface-500">${esc(REASON_LABELS[ins.reason] || ins.reason || '')}</span>`}</div>
        ${row('取得した予約（今月分の判定用）', `${(d.fetched ?? 0).toLocaleString('ja-JP')}件${d.truncated ? '（上限で打ち切り）' : ''}`)}
        ${row('開始日時 / 作成日時の項目', `${esc(d.startKey || '—')} / ${esc(d.createdKey || '—')}`)}
        ${row('新規の判定', esc({ appointment: '予約の新規フラグ', customer_first_visit: '顧客の初回来店日', visit_source: '流入元が入っている来店（新規来店数と一致）' }[d.newSource] || '判定できず（新規/既存の区別なし）'))}
        ${row('状態の内訳（判定結果）', `会計済み ${d.kindCounts?.done ?? 0}・予約中 ${d.kindCounts?.open ?? 0}・キャンセル ${d.kindCounts?.canceled ?? 0}・無断 ${d.kindCounts?.no_show ?? 0}・来店以外 ${d.kindCounts?.non_visit ?? 0}`)}
        ${row('売上サマリとの突き合わせ（今月・今日まで）', d.calibration
            ? `来店数 ${d.calibration.summaryVisits ?? '—'}名 / 予約データの会計済み ${d.calibration.done}件${d.calibration.summaryNew !== null && d.calibration.summaryNew !== undefined ? `（新規来店 ${d.calibration.summaryNew}名 / 流入元つき ${d.calibration.doneWithSource}件）` : ''}`
            : '—')}
        ${row('SalonOneのステータス値', esc(Object.entries(d.statusCounts || {}).map(([k, v]) => `${k}: ${v}`).join('、') || '—'))}
        ${row('推定の次回予約率（今月）', estRate)}
        ${row('日報に入力された次回予約（今月）', `${manualNext.toLocaleString('ja-JP')}名`)}
        ${row('会計未処理の過去予約（今月）', `${ins.unsettled?.count ?? 0}件`)}
        <details class="mt-3 text-xs text-surface-500"><summary class="cursor-pointer">予約明細の項目一覧・値の分布</summary>
            <p class="mt-2 break-all">${esc((d.fields || []).join(', ') || '—')}</p>
            ${Object.entries(d.valueCounts || {}).map(([k, v]) => `<p class="mt-1 break-all"><b>${esc(k)}</b>: ${esc(Object.entries(v).map(([a, n]) => `${a}=${n}`).join('、'))}</p>`).join('')}
        </details>`;
}
