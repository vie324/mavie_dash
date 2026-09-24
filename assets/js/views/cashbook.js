// 出納帳タブ: 現金の入出金の記録・レジの実査・日次締め・SalonOneの現金売上との突合・操作ログ
// オーナー・マネージャー・店長が利用（スタッフ不可）。レジ（店舗）ごとにつける。
// 帳簿残高と過不足はサーバーが計算する（api/_lib/cashbook.js）。この画面は表示と入力のみ。

import { state, on, isAdminLike, isStoreLocked, currentShopId, shopName, staffsOfShop } from '../core/state.js';
import { yen, esc, todayStr, dowJa } from '../core/format.js';
import { ApiError } from '../core/api.js';
import { toast } from '../core/engage.js';
import { renderShopPick } from '../ui/shoppick.js';
import { loadCashbook, getCashbook, cashbookAction } from '../data/cashbook.js';

const DENOMS = [10000, 5000, 2000, 1000, 500, 100, 50, 10, 5, 1];
const DENOM_LABEL = { 10000: '1万円', 5000: '5千円', 2000: '2千円', 1000: '千円', 500: '500円', 100: '100円', 50: '50円', 10: '10円', 5: '5円', 1: '1円' };

const STATUS = {
    idle: { label: '動きなし', cls: 'cb-chip-muted' },
    open: { label: '未締め', cls: 'cb-chip-warn' },
    counted: { label: '実査済み・未締め', cls: 'cb-chip-warn' },
    closed: { label: '締め済み', cls: 'cb-chip-ok' },
    closed_diff: { label: '締め済み（過不足あり）', cls: 'cb-chip-diff' },
};

const ERROR_TEXT = {
    day_closed: 'この日は締め済みです。修正するには締めを取り消してください',
    count_required: '先にレジの現金を数えて保存してください',
    reason_required: '理由を入力してください',
    salon_unavailable: 'SalonOneの売上を取得できませんでした。時間をおいて再度お試しください',
    storage_unconfigured: '出納帳にはサーバー保存（Supabase）の設定が必要です',
    setup_required: '先に出納帳の開始日と残高を設定してください',
    future_date: '未来の日付には記録できません',
    forbidden: 'この操作の権限がありません',
};

const LOG_FILTERS = {
    all: () => true,
    entry: l => l.action === 'entry.add' || l.action === 'entry.edit',
    close: l => l.action === 'count.save' || l.action === 'day.close',
    void: l => l.action === 'entry.void' || l.action === 'day.reopen',
    settings: l => l.action.startsWith('settings.'),
};

let selectedDate = todayStr();
let form = null;          // { mode: 'add'|'edit', type, entryId }
let countDraft = null;    // { key: 'shop:date', denoms: {d: n}, dirty }
let voidTarget = null;    // 取消の理由入力中の entryId
let reopenOpen = false;
let logFilter = 'all';
let busy = false;
let loadError = null;

// ホームの「やること」などから日付を指定して開く
export function presetCashbook(date) {
    if (date) selectedDate = date;
}

// 更新ボタン・引っ張って更新
export function reload() {
    return refresh();
}

function activeShopId() {
    const id = currentShopId();
    return id === 'all' ? null : String(id);
}
function monthSel() {
    return selectedDate.slice(0, 7);
}
function currentView() {
    const shopId = activeShopId();
    return shopId ? getCashbook(shopId, monthSel()) : null;
}
function addDays(date, n) {
    const d = new Date(`${date}T00:00:00Z`);
    d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
}
function md(date) {
    return `${Number(date.slice(5, 7))}/${Number(date.slice(8, 10))}`;
}
function hm(iso) {
    if (!iso) return '';
    const d = new Date(Date.parse(iso) + 9 * 3600e3);
    return `${d.getUTCMonth() + 1}/${d.getUTCDate()} ${String(d.getUTCHours()).padStart(2, '0')}:${String(d.getUTCMinutes()).padStart(2, '0')}`;
}
function fullTime(iso) {
    if (!iso) return '';
    return new Date(Date.parse(iso) + 9 * 3600e3).toISOString().slice(0, 19).replace('T', ' ');
}
function signed(n) {
    if (n === null || n === undefined) return '—';
    if (n === 0) return '±0';
    return (n > 0 ? '+' : '−') + '¥' + Math.abs(n).toLocaleString('ja-JP');
}
function diffCls(n) {
    if (n === null || n === undefined) return 'text-surface-400';
    return n === 0 ? 'cb-ok' : 'cb-ng';
}

// ---- 担当者（実際に作業したスタッフ。操作ログに残る）----
function opStoreKey(shopId) { return `vie_cb_op_${shopId}`; }
function getOp(shopId) { try { return localStorage.getItem(opStoreKey(shopId)) || ''; } catch (_) { return ''; } }
function setOp(shopId, v) { try { localStorage.setItem(opStoreKey(shopId), v); } catch (_) { /* ignore */ } }

function catName(v, id) {
    return v?.settings?.categories?.find(c => c.id === id)?.name || id;
}

export function init() {
    on('tab:shown', id => { if (id === 'cashbook') refresh(); });
    on('filters', () => {
        // 店舗が変わったら入力途中の状態は捨てる
        form = null; voidTarget = null; reopenOpen = false; countDraft = null;
        if (state.ui.activeTab === 'cashbook') refresh();
    });
    on('data:cashbook', () => { if (state.ui.activeTab === 'cashbook') render(); });
    on('masters', () => { if (state.ui.activeTab === 'cashbook') render(); });

    const root = document.getElementById('content-cashbook');
    if (!root) return;

    const dateInput = document.getElementById('cb-date');
    dateInput?.addEventListener('click', () => { try { dateInput.showPicker?.(); } catch (_) { /* 標準動作 */ } });
    dateInput?.addEventListener('change', ev => setDate(ev.target.value || todayStr()));
    document.getElementById('cb-prev-day')?.addEventListener('click', () => setDate(addDays(selectedDate, -1)));
    document.getElementById('cb-next-day')?.addEventListener('click', () => setDate(addDays(selectedDate, 1)));
    document.getElementById('cb-today')?.addEventListener('click', () => setDate(todayStr()));
    document.getElementById('cb-operator')?.addEventListener('change', ev => {
        const shopId = activeShopId();
        if (shopId) setOp(shopId, ev.target.value);
    });
    document.getElementById('cb-add-in')?.addEventListener('click', () => openForm('add', 'in'));
    document.getElementById('cb-add-out')?.addEventListener('click', () => openForm('add', 'out'));
    document.getElementById('cb-csv')?.addEventListener('click', exportBookCsv);
    document.getElementById('cb-log-csv')?.addEventListener('click', exportLogCsv);
    document.getElementById('cb-log-filter')?.addEventListener('change', ev => { logFilter = ev.target.value; renderLog(); });
    document.getElementById('cb-log-card')?.addEventListener('toggle', ev => {
        if (ev.target.open) loadLog();
    });

    root.addEventListener('click', onClick);
    root.addEventListener('input', onInput);
    root.addEventListener('change', onChange);
    root.addEventListener('submit', ev => ev.preventDefault());
}

function setDate(date) {
    const today = todayStr();
    if (date > today) date = today;
    if (countDraft?.dirty && !confirm('数えた枚数が保存されていません。破棄して日付を切り替えますか？')) {
        const el = document.getElementById('cb-date');
        if (el) el.value = selectedDate;
        return;
    }
    const monthChanged = date.slice(0, 7) !== monthSel();
    selectedDate = date;
    form = null; voidTarget = null; reopenOpen = false; countDraft = null;
    if (monthChanged) refresh();
    else render();
}

async function refresh() {
    const shopId = activeShopId();
    loadError = null;
    render();
    if (!shopId) return;
    const logOpen = document.getElementById('cb-log-card')?.open;
    try {
        await loadCashbook(shopId, monthSel(), { log: !!logOpen });
    } catch (e) {
        loadError = e;
        render();
    }
}

async function loadLog() {
    const shopId = activeShopId();
    if (!shopId) return;
    try {
        await loadCashbook(shopId, monthSel(), { log: true });
    } catch (e) {
        console.warn('cashbook log', e);
    }
}

// ---- 保存系 ----
async function act(body, okMessage) {
    const shopId = activeShopId();
    if (!shopId || busy) return null;
    busy = true;
    document.getElementById('content-cashbook')?.classList.add('cb-busy');
    try {
        const op = getOp(shopId);
        const res = await cashbookAction({ shopId: Number(shopId), ...(op ? { opId: Number(op) } : {}), ...body });
        if (okMessage) toast(okMessage);
        if (document.getElementById('cb-log-card')?.open) loadLog();
        return res;
    } catch (e) {
        const msg = e instanceof ApiError ? (e.body?.detail || ERROR_TEXT[e.code] || '保存に失敗しました') : '保存に失敗しました（通信エラー）';
        toast(msg, 'error');
        return null;
    } finally {
        busy = false;
        document.getElementById('content-cashbook')?.classList.remove('cb-busy');
    }
}

function onClick(ev) {
    const t = ev.target;
    const btn = t.closest('button');
    if (!btn) return;
    if (btn.dataset.cbAction === 'setup') return saveSetup();
    if (btn.dataset.cbAction === 'form-cancel') { form = null; renderEntries(); return; }
    if (btn.dataset.cbAction === 'form-save') return saveForm();
    if (btn.dataset.cbAction === 'form-type') {
        // 入力途中の値は残して区分だけ切り替える（科目は区分が違うので選び直し）
        form.values = { ...form.values, ...readFormValues(), cat: '' };
        form.type = btn.dataset.type;
        renderEntries();
        return;
    }
    if (btn.dataset.edit) { openForm('edit', null, btn.dataset.edit); return; }
    if (btn.dataset.void) { voidTarget = btn.dataset.void; renderEntries(); document.getElementById('cb-void-reason')?.focus(); return; }
    if (btn.dataset.cbAction === 'void-cancel') { voidTarget = null; renderEntries(); return; }
    if (btn.dataset.cbAction === 'void-confirm') return confirmVoid();
    if (btn.dataset.step && btn.dataset.denom) { stepDenom(Number(btn.dataset.denom), Number(btn.dataset.step)); return; }
    if (btn.dataset.cbAction === 'count-save') return saveCount();
    if (btn.dataset.cbAction === 'count-reset') { countDraft = null; renderCount(); renderClose(); return; }
    if (btn.dataset.cbAction === 'close') return closeDay();
    if (btn.dataset.cbAction === 'reopen-open') { reopenOpen = true; renderClose(); document.getElementById('cb-reopen-reason')?.focus(); return; }
    if (btn.dataset.cbAction === 'reopen-cancel') { reopenOpen = false; renderClose(); return; }
    if (btn.dataset.cbAction === 'reopen') return reopenDay();
    if (btn.dataset.gotoDate) { setDate(btn.dataset.gotoDate); document.getElementById('cb-day-card')?.scrollIntoView({ behavior: 'smooth', block: 'start' }); return; }
    if (btn.dataset.cbAction === 'deposit') { recordDeposit(Number(btn.dataset.amount) || ''); return; }
    if (btn.dataset.cbAction === 'save-shop') return saveShopSettings();
    if (btn.dataset.cbAction === 'save-brand') return saveBrandSettings();
    if (btn.dataset.cbAction === 'add-cat') { addCategoryRow(); return; }
}

// 預け入れは締めた日ではなく今日の出金として記録する（銀行へ持っていくのは翌日以降のため）
function recordDeposit(amount) {
    const today = todayStr();
    const open = () => openForm('add', 'out', null, { cat: 'out_deposit', amount });
    if (selectedDate === today) { open(); return; }
    const monthChanged = today.slice(0, 7) !== monthSel();
    selectedDate = today;
    form = null; voidTarget = null; reopenOpen = false; countDraft = null;
    if (monthChanged) refresh().then(open);
    else { render(); open(); }
}

function onInput(ev) {
    const t = ev.target;
    if (t.matches('input[data-denom]')) {
        const d = Number(t.dataset.denom);
        const n = Math.max(0, Math.min(99999, Math.floor(Number(t.value) || 0)));
        ensureDraft();
        countDraft.denoms[d] = n;
        countDraft.dirty = true;
        updateCountTotals();
    }
    if (t.id === 'cb-close-reason') updateCloseButton();
}

function onChange(ev) {
    const t = ev.target;
    if (t.id === 'cb-cash-auto') {
        document.getElementById('cb-cash-methods')?.classList.toggle('hidden', t.checked);
    }
}

// ---- 表示 ----
function render() {
    if (!isAdminLike() && !isStoreLocked()) return;
    const shopId = activeShopId();
    const guard = document.getElementById('cb-shop-guard');
    const body = document.getElementById('cb-body');
    const setup = document.getElementById('cb-setup');
    const notice = document.getElementById('cb-notice');
    guard?.classList.toggle('hidden', !!shopId);
    if (!shopId) {
        body?.classList.add('hidden');
        setup?.classList.add('hidden');
        notice?.classList.add('hidden');
        renderShopPick('cb-shop-pick');
        return;
    }
    const v = currentView();
    if (loadError || v?.storage === 'none') {
        body?.classList.add('hidden');
        setup?.classList.add('hidden');
        notice.classList.remove('hidden');
        notice.innerHTML = `<div class="premium-card p-6 text-center">
            <i data-lucide="${v?.storage === 'none' ? 'database' : 'wifi-off'}" class="w-10 h-10 mx-auto text-surface-400 mb-3"></i>
            <p class="font-semibold text-accent-800 dark:text-gray-200 mb-1">${v?.storage === 'none' ? '出納帳にはサーバー保存の設定が必要です' : '出納帳を読み込めませんでした'}</p>
            <p class="text-sm text-surface-500">${v?.storage === 'none'
                ? '記録と操作ログを全端末で共有・保管するため、Supabase の設定後に利用できます（設定タブの「連携状態」を参照）。'
                : esc(loadError?.body?.detail || '通信状態を確認して、右上の更新ボタンで再読み込みしてください。')}</p>
        </div>`;
        if (window.lucide) lucide.createIcons({ nodes: [...notice.querySelectorAll('[data-lucide]')] });
        return;
    }
    notice?.classList.add('hidden');
    if (!v) { body?.classList.add('hidden'); return; } // 読み込み中
    if (v.setupRequired) {
        body?.classList.add('hidden');
        setup.classList.remove('hidden');
        renderSetup(v);
        return;
    }
    setup?.classList.add('hidden');
    body?.classList.remove('hidden');
    setText('cb-shop-label', shopName(shopId));
    renderDateBar(v);
    renderOperator(shopId);
    renderAlerts(v);
    renderBalance(v);
    renderEntries();
    renderCount();
    renderClose();
    renderMonth(v);
    renderLog();
    renderSettings(v);
    if (window.lucide) lucide.createIcons({ nodes: [...document.querySelectorAll('#content-cashbook [data-lucide]')] });
}

function rowOf(v, date = selectedDate) {
    return v?.rows?.find(r => r.date === date) || null;
}

function renderSetup(v) {
    const el = document.getElementById('cb-setup');
    const today = todayStr();
    el.innerHTML = `
        <div class="flex items-center gap-3 mb-3">
            <div class="w-10 h-10 rounded-lg gradient-sage flex items-center justify-center"><i data-lucide="wallet" class="w-5 h-5 text-white"></i></div>
            <div><h3 class="text-lg font-display font-bold text-accent-900">出納帳を始める</h3>
            <p class="text-xs text-surface-500">${esc(shopName(v.shopId))}</p></div>
        </div>
        <ol class="cb-steps">
            <li>開始日の<b>開店前にレジにある現金</b>を数えて「開始時の残高」に入力します</li>
            <li>SalonOneで<b>現金払いの会計は自動で入金</b>になります。経費の支払い・銀行への預け入れ・釣り銭の補充などを記録します</li>
            <li>閉店後にレジの現金を数えて<b>締め</b>ると、SalonOneの現金売上との過不足が確定し、操作ログに残ります</li>
        </ol>
        <div class="cb-form-grid mt-4">
            <label class="cb-field"><span>開始日</span>
                <input type="date" id="cb-setup-start" value="${today}" max="${today}" class="cb-input"></label>
            <label class="cb-field"><span>開始時の残高（レジの現金）</span>
                <input type="number" id="cb-setup-balance" inputmode="numeric" min="0" step="1" placeholder="例: 50000" class="cb-input tabular-nums"></label>
            <label class="cb-field"><span>釣り銭準備金（任意）</span>
                <input type="number" id="cb-setup-float" inputmode="numeric" min="0" step="1" placeholder="例: 30000" class="cb-input tabular-nums">
                <small>毎日レジに残しておく額。締めのときに「預け入れの目安」を表示します</small></label>
        </div>
        <button type="button" class="btn-primary w-full sm:w-auto mt-4 py-3 px-8" data-cb-action="setup">出納帳を始める</button>`;
    if (window.lucide) lucide.createIcons({ nodes: [...el.querySelectorAll('[data-lucide]')] });
}

async function saveSetup() {
    const startDate = document.getElementById('cb-setup-start')?.value;
    const bal = document.getElementById('cb-setup-balance')?.value;
    const fl = document.getElementById('cb-setup-float')?.value;
    if (!startDate) { toast('開始日を選んでください', 'warn'); return; }
    if (bal === '' || Number(bal) < 0) { toast('開始時の残高を入力してください', 'warn'); return; }
    const res = await act({ action: 'settings.shop', startDate, initialBalance: Math.round(Number(bal)), float: fl === '' ? 0 : Math.round(Number(fl)), month: monthSel() }, '出納帳を開始しました');
    if (res) {
        if (selectedDate < startDate) selectedDate = startDate;
        refresh();
    }
}

function renderDateBar(v) {
    const today = todayStr();
    const label = document.getElementById('cb-date-label');
    if (label) label.textContent = `${Number(selectedDate.slice(5, 7))}月${Number(selectedDate.slice(8, 10))}日（${dowJa(selectedDate)}）${selectedDate === today ? '・今日' : ''}`;
    const input = document.getElementById('cb-date');
    if (input) {
        input.value = selectedDate;
        input.max = today;
        if (v.settings.shop?.startDate) input.min = v.settings.shop.startDate;
    }
    const next = document.getElementById('cb-next-day');
    if (next) next.disabled = selectedDate >= today;
    const prev = document.getElementById('cb-prev-day');
    if (prev) prev.disabled = !!v.settings.shop?.startDate && selectedDate <= v.settings.shop.startDate;
    const todayBtn = document.getElementById('cb-today');
    if (todayBtn) todayBtn.classList.toggle('invisible', selectedDate === today);
    const row = rowOf(v);
    const chip = document.getElementById('cb-status-chip');
    if (chip) {
        const st = row ? STATUS[row.status] : null;
        chip.className = `cb-chip ${st ? st.cls : 'cb-chip-muted'}`;
        chip.textContent = st ? st.label : (selectedDate < (v.settings.shop?.startDate || '') ? '開始日より前' : '—');
    }
}

function renderOperator(shopId) {
    const sel = document.getElementById('cb-operator');
    if (!sel) return;
    const staffs = staffsOfShop(shopId);
    const cur = getOp(shopId);
    sel.innerHTML = `<option value="">（選ばない）</option>` + staffs.map(s => `<option value="${s.id}"${String(s.id) === cur ? ' selected' : ''}>${esc(s.name)}</option>`).join('');
}

function renderAlerts(v) {
    const el = document.getElementById('cb-alerts');
    if (!el) return;
    const row = rowOf(v);
    const out = [];
    const alert = (level, icon, html) => out.push(`<div class="cb-alert cb-alert-${level}"><i data-lucide="${icon}" class="w-4 h-4 flex-shrink-0 mt-0.5"></i><div>${html}</div></div>`);
    if (v.salonError) alert('danger', 'cloud-off', 'SalonOneの売上を取得できないため、帳簿残高を計算できません。時間をおいて更新してください');
    if (v.openingUncertain) alert('warn', 'history', '2年以上締めていないため、繰越額が正確でない可能性があります。設定の開始日を見直してください');
    if (row?.salonChanged) {
        alert('danger', 'alert-triangle', `締めた後に SalonOne の現金売上が変わりました（締め時 ${yen(row.closed.salonCash)} → 現在 ${yen(row.salonCash)}）。SalonOneで会計の修正がなかったか確認し、必要なら締めを取り消して締め直してください`);
    }
    if (row?.openingChanged) {
        alert('warn', 'alert-triangle', `締めた後に前日までの記録が変わり、繰越が ${yen(row.closed.opening)} → ${yen(row.opening)} になっています。前日までの記録を確認してください`);
    }
    if (row && row.unset > 0) alert('warn', 'help-circle', `支払い方法が未設定の会計が ${yen(row.unset)} あります。現金払いなら SalonOne で支払い方法を「現金」に直すと、ここに反映されます`);
    if (row && row.refund > 0) alert('info', 'undo-2', `SalonOneの返金 ${yen(row.refund)}。現金で返した場合は「返金（現金）」として出金を記録してください`);
    const pending = (v.summary?.pendingDays || []).filter(d => d !== selectedDate);
    if (pending.length) {
        const chips = pending.slice(0, 6).map(d => `<button type="button" class="cb-date-chip" data-goto-date="${d}">${md(d)}</button>`).join('');
        alert('warn', 'calendar-x', `締めていない日があります ${chips}${pending.length > 6 ? ` ほか${pending.length - 6}日` : ''}`);
    }
    const changed = (v.summary?.changedDays || []).filter(d => d !== selectedDate);
    if (changed.length) {
        const chips = changed.slice(0, 6).map(d => `<button type="button" class="cb-date-chip" data-goto-date="${d}">${md(d)}</button>`).join('');
        alert('danger', 'alert-triangle', `締めた後に金額が変わった日があります ${chips}`);
    }
    if (!row && v.settings.shop?.startDate && selectedDate < v.settings.shop.startDate) {
        alert('info', 'info', `出納帳は ${md(v.settings.shop.startDate)} から始まっています`);
    }
    el.innerHTML = out.join('');
    el.classList.toggle('hidden', out.length === 0);
}

function renderBalance(v) {
    const el = document.getElementById('cb-balance');
    if (!el) return;
    const row = rowOf(v);
    if (!row) { el.innerHTML = ''; return; }
    const methods = row.cashMethods?.length > 1 ? `<small>${row.cashMethods.map(m => `${esc(m.name)} ${yen(m.amount)}`).join('・')}</small>` : '';
    const counted = row.closed ? row.closed.counted : row.count?.total;
    const diff = row.closed ? row.closed.diff : row.diff;
    el.innerHTML = `
        <div class="cb-balance">
            <div class="cb-line"><span>前日からの繰越</span><b>${row.opening === null ? '—' : yen(row.opening)}</b></div>
            <div class="cb-line cb-plus"><span>＋ 現金売上 <em>SalonOne</em>${methods}</span><b>${row.salonCash === null ? '—' : yen(row.salonCash)}</b></div>
            <div class="cb-line cb-plus"><span>＋ 入金（記録）</span><b>${yen(row.manualIn)}</b></div>
            <div class="cb-line cb-minus"><span>− 出金（記録）</span><b>${yen(row.manualOut)}</b></div>
            <div class="cb-line cb-total"><span>＝ 帳簿上の残高</span><b>${row.expected === null ? '—' : yen(row.expected)}</b></div>
            <div class="cb-line cb-count"><span>数えた現金（実査）</span><b>${counted === undefined || counted === null ? '<span class="text-surface-400 font-medium">未実施</span>' : yen(counted)}</b></div>
            <div class="cb-line cb-diff"><span>過不足</span><b class="${diffCls(diff)}">${counted === undefined || counted === null ? '—' : signed(diff)}</b></div>
        </div>`;
}

// ---- 入金・出金 ----
function openForm(mode, type, entryId, preset) {
    const v = currentView();
    const row = rowOf(v);
    if (row?.closed) { toast('この日は締め済みです。修正するには締めを取り消してください', 'warn'); return; }
    if (mode === 'edit') {
        const e = v.entries.find(x => x.id === entryId);
        if (!e) return;
        form = { mode, type: e.type, entryId, values: { ...e } };
    } else {
        form = { mode, type, entryId: null, values: preset || {} };
    }
    voidTarget = null;
    renderEntries();
    const amount = document.getElementById('cb-f-amount');
    amount?.focus();
    document.getElementById('cb-form')?.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function formHtml(v) {
    const cats = v.settings.categories.filter(c => c.type === form.type && (!c.disabled || c.id === form.values.cat));
    const val = form.values;
    const today = todayStr();
    return `
    <div class="cb-form">
        <div class="cb-seg" role="tablist">
            <button type="button" class="cb-seg-btn${form.type === 'in' ? ' active in' : ''}" data-cb-action="form-type" data-type="in">入金</button>
            <button type="button" class="cb-seg-btn${form.type === 'out' ? ' active out' : ''}" data-cb-action="form-type" data-type="out">出金</button>
        </div>
        <div class="cb-form-grid">
            <label class="cb-field"><span>金額</span>
                <input type="number" id="cb-f-amount" inputmode="numeric" min="1" step="1" value="${val.amount ?? ''}" placeholder="0" class="cb-input cb-input-amount tabular-nums"></label>
            <label class="cb-field"><span>科目</span>
                <select id="cb-f-cat" class="cb-input"><option value="">科目を選ぶ</option>${cats.map(c => `<option value="${c.id}"${c.id === val.cat ? ' selected' : ''}>${esc(c.name)}</option>`).join('')}</select></label>
            <label class="cb-field"><span>摘要（内容）</span>
                <input type="text" id="cb-f-memo" maxlength="100" value="${esc(val.memo || '')}" placeholder="${form.type === 'out' ? '例: コットン・綿棒' : '例: 両替 1万円分'}" class="cb-input"></label>
            <label class="cb-field"><span>${form.type === 'out' ? '支払先' : '入金元'}（任意）</span>
                <input type="text" id="cb-f-payee" maxlength="40" value="${esc(val.payee || '')}" placeholder="${form.type === 'out' ? '例: ◯◯ドラッグ' : '例: 本部'}" class="cb-input"></label>
            ${form.mode === 'edit' ? `<label class="cb-field"><span>日付</span>
                <input type="date" id="cb-f-date" value="${val.date}" min="${monthSel()}-01" max="${today}" class="cb-input"></label>` : ''}
            ${form.type === 'out' ? `<label class="cb-check"><input type="checkbox" id="cb-f-receipt"${val.receipt ? ' checked' : ''}> 領収書・レシートあり</label>` : ''}
        </div>
        <div class="cb-form-actions">
            <button type="button" class="input-clear-btn" data-cb-action="form-cancel">キャンセル</button>
            <button type="button" class="btn-primary cb-save-btn" data-cb-action="form-save">${form.mode === 'edit' ? '修正を保存' : `${form.type === 'in' ? '入金' : '出金'}を記録`}</button>
        </div>
    </div>`;
}

function renderEntries() {
    const v = currentView();
    const formEl = document.getElementById('cb-form');
    const list = document.getElementById('cb-entries');
    if (!v || !formEl || !list) return;
    const row = rowOf(v);
    const locked = !!row?.closed || !row;
    // 再描画で入力途中の値が消えないように、描画前に画面の値を拾っておく
    if (form && document.getElementById('cb-f-amount')) {
        const cat = document.getElementById('cb-f-cat')?.value;
        form.values = { ...form.values, ...readFormValues(), ...(cat !== undefined ? { cat } : {}) };
    }
    for (const id of ['cb-add-in', 'cb-add-out']) {
        const b = document.getElementById(id);
        if (b) b.disabled = locked;
    }
    formEl.innerHTML = form ? formHtml(v) : '';
    formEl.classList.toggle('hidden', !form);
    const entries = v.entries.filter(e => e.date === selectedDate);
    if (!entries.length) {
        list.innerHTML = `<p class="cb-empty">${row?.closed ? 'この日の記録はありません（締め済み）' : 'この日の入金・出金の記録はありません。<br>SalonOneの現金会計は自動で「現金売上」に入ります。'}</p>`;
        return;
    }
    list.innerHTML = entries.map(e => {
        const sign = e.type === 'in' ? '+' : '−';
        const meta = [
            e.memo ? esc(e.memo) : '',
            e.payee ? `${e.type === 'out' ? '支払先' : '入金元'}: ${esc(e.payee)}` : '',
            e.receipt ? '<span class="cb-tag">領収書あり</span>' : '',
            e.rev > 1 ? '<span class="cb-tag cb-tag-muted">修正済み</span>' : '',
        ].filter(Boolean).join(' · ');
        const who = `${hm(e.createdAt)} ${esc(e.createdBy || '')}${e.op ? ` · 担当 ${esc(e.op)}` : ''}`;
        const voidBox = voidTarget === e.id ? `
            <div class="cb-inline-confirm">
                <input type="text" id="cb-void-reason" maxlength="200" placeholder="取り消す理由（必須）例: 二重に記録した" class="cb-input">
                <div class="cb-form-actions">
                    <button type="button" class="input-clear-btn" data-cb-action="void-cancel">やめる</button>
                    <button type="button" class="cb-danger-btn" data-cb-action="void-confirm">取り消す</button>
                </div>
            </div>` : '';
        return `
        <div class="cb-entry${e.voided ? ' voided' : ''}">
            <div class="cb-entry-icon ${e.type}"><i data-lucide="${e.type === 'in' ? 'arrow-down-left' : 'arrow-up-right'}" class="w-4 h-4"></i></div>
            <div class="cb-entry-main">
                <div class="cb-entry-top"><span class="cb-entry-cat">${esc(catName(v, e.cat))}</span>
                    <b class="cb-entry-amount ${e.type}">${sign}${yen(e.amount)}</b></div>
                ${meta ? `<div class="cb-entry-meta">${meta}</div>` : ''}
                <div class="cb-entry-who">${who}</div>
                ${e.voided ? `<div class="cb-entry-void">取消: ${esc(e.voided.reason)}（${hm(e.voided.at)} ${esc(e.voided.by || '')}）</div>` : ''}
                ${!e.voided && !row?.closed && voidTarget !== e.id ? `<div class="cb-entry-actions">
                    <button type="button" class="cb-link-btn" data-edit="${e.id}"><i data-lucide="pencil" class="w-3.5 h-3.5"></i>修正</button>
                    <button type="button" class="cb-link-btn danger" data-void="${e.id}"><i data-lucide="x-circle" class="w-3.5 h-3.5"></i>取消</button>
                </div>` : ''}
                ${voidBox}
            </div>
        </div>`;
    }).join('');
    if (window.lucide) lucide.createIcons({ nodes: [...list.querySelectorAll('[data-lucide]'), ...formEl.querySelectorAll('[data-lucide]')] });
}

function readFormValues() {
    const g = id => document.getElementById(id);
    const out = {};
    if (g('cb-f-amount')) out.amount = g('cb-f-amount').value;
    if (g('cb-f-memo')) out.memo = g('cb-f-memo').value;
    if (g('cb-f-payee')) out.payee = g('cb-f-payee').value;
    if (g('cb-f-date')) out.date = g('cb-f-date').value;
    if (g('cb-f-receipt')) out.receipt = g('cb-f-receipt').checked;
    return out;
}

async function saveForm() {
    const v = currentView();
    if (!form || !v) return;
    const amount = Math.round(Number(document.getElementById('cb-f-amount')?.value));
    if (!amount || amount < 1) { toast('金額を入力してください', 'warn'); document.getElementById('cb-f-amount')?.focus(); return; }
    if (!document.getElementById('cb-f-cat')?.value) { toast('科目を選んでください', 'warn'); document.getElementById('cb-f-cat')?.focus(); return; }
    const fields = {
        type: form.type,
        cat: document.getElementById('cb-f-cat')?.value,
        amount,
        memo: document.getElementById('cb-f-memo')?.value || '',
        payee: document.getElementById('cb-f-payee')?.value || '',
        receipt: form.type === 'out' ? !!document.getElementById('cb-f-receipt')?.checked : false,
    };
    let res;
    if (form.mode === 'edit') {
        const date = document.getElementById('cb-f-date')?.value || selectedDate;
        res = await act({ action: 'entry.edit', date: selectedDate, entryId: form.entryId, patch: { ...fields, date } }, '修正しました');
    } else {
        res = await act({ action: 'entry.add', date: selectedDate, ...fields }, `${form.type === 'in' ? '入金' : '出金'}を記録しました`);
    }
    if (res) { form = null; render(); }
}

async function confirmVoid() {
    const reason = document.getElementById('cb-void-reason')?.value.trim();
    if (!reason) { toast('取り消す理由を入力してください', 'warn'); return; }
    const res = await act({ action: 'entry.void', date: selectedDate, entryId: voidTarget, reason }, '記録を取り消しました（ログに残ります）');
    if (res) { voidTarget = null; render(); }
}

// ---- 実査（金種ごとの枚数）----
function draftKey() {
    return `${activeShopId()}:${selectedDate}`;
}
function ensureDraft() {
    if (countDraft && countDraft.key === draftKey()) return;
    const row = rowOf(currentView());
    const saved = row?.closed?.denoms || row?.count?.denoms || {};
    countDraft = { key: draftKey(), denoms: Object.fromEntries(DENOMS.map(d => [d, Number(saved[d] || 0)])), dirty: false };
}
function draftTotal() {
    return DENOMS.reduce((a, d) => a + d * (countDraft?.denoms[d] || 0), 0);
}
function stepDenom(d, step) {
    ensureDraft();
    countDraft.denoms[d] = Math.max(0, (countDraft.denoms[d] || 0) + step);
    countDraft.dirty = true;
    const input = document.querySelector(`#cb-count input[data-denom="${d}"]`);
    if (input) input.value = countDraft.denoms[d] || '';
    updateCountTotals();
}

function renderCount() {
    const el = document.getElementById('cb-count');
    const v = currentView();
    if (!el || !v) return;
    const row = rowOf(v);
    if (!row) { el.innerHTML = ''; return; }
    ensureDraft();
    const locked = !!row.closed;
    el.innerHTML = `
        <div class="cb-denoms${locked ? ' locked' : ''}">
            ${DENOMS.map(d => `
            <div class="cb-denom">
                <span class="cb-denom-label${d >= 1000 ? ' bill' : ''}">${DENOM_LABEL[d]}</span>
                <div class="stepper">
                    <button type="button" class="stepper-btn" data-step="-1" data-denom="${d}" aria-label="${DENOM_LABEL[d]}を減らす"${locked ? ' disabled' : ''}>−</button>
                    <input type="number" class="stepper-input" inputmode="numeric" pattern="[0-9]*" min="0" max="99999" placeholder="0"
                        value="${countDraft.denoms[d] || ''}" data-denom="${d}" aria-label="${DENOM_LABEL[d]}の枚数"${locked ? ' disabled' : ''}>
                    <button type="button" class="stepper-btn" data-step="1" data-denom="${d}" aria-label="${DENOM_LABEL[d]}を増やす"${locked ? ' disabled' : ''}>＋</button>
                </div>
                <span class="cb-denom-sub tabular-nums" data-sub="${d}">${yen(d * (countDraft.denoms[d] || 0))}</span>
            </div>`).join('')}
        </div>
        <div class="cb-count-bar${locked ? ' hidden' : ''}">
            <div class="cb-count-sum">
                <span>数えた合計 <b id="cb-count-total" class="tabular-nums"></b></span>
                <span id="cb-count-diff" class="tabular-nums"></span>
                <small id="cb-count-saved"></small>
            </div>
            <div class="cb-form-actions">
                <button type="button" class="input-clear-btn" data-cb-action="count-reset">元に戻す</button>
                <button type="button" class="btn-primary cb-save-btn" data-cb-action="count-save">実査を保存</button>
            </div>
        </div>`;
    updateCountTotals();
}

function updateCountTotals() {
    const v = currentView();
    const row = rowOf(v);
    if (!row || !countDraft) return;
    for (const d of DENOMS) {
        const sub = document.querySelector(`#cb-count [data-sub="${d}"]`);
        if (sub) sub.textContent = yen(d * (countDraft.denoms[d] || 0));
    }
    const total = draftTotal();
    setText('cb-count-total', yen(total));
    const diffEl = document.getElementById('cb-count-diff');
    if (diffEl) {
        const diff = row.expected === null ? null : total - row.expected;
        diffEl.className = `tabular-nums ${diffCls(diff)}`;
        diffEl.textContent = row.expected === null ? '' : `帳簿との差 ${signed(diff)}`;
    }
    const saved = document.getElementById('cb-count-saved');
    if (saved) {
        saved.textContent = countDraft.dirty ? '未保存の変更があります'
            : row.count ? `保存済み ${hm(row.count.at)} ${row.count.by || ''}${row.count.op ? ` · 担当 ${row.count.op}` : ''}` : '';
        saved.classList.toggle('cb-unsaved', !!countDraft.dirty);
    }
}

async function saveCount() {
    ensureDraft();
    const denoms = Object.fromEntries(DENOMS.map(d => [d, countDraft.denoms[d] || 0]));
    const res = await act({ action: 'count.save', date: selectedDate, denoms }, '実査を保存しました');
    if (res) { countDraft = null; render(); }
}

// ---- 締め ----
function renderClose() {
    const el = document.getElementById('cb-close');
    const v = currentView();
    if (!el || !v) return;
    const row = rowOf(v);
    if (!row) { el.innerHTML = ''; return; }
    const float = v.settings.shop?.float || 0;
    if (row.closed) {
        const c = row.closed;
        const deposit = float > 0 && c.counted > float ? c.counted - float : 0;
        el.innerHTML = `
        <div class="cb-closed ${c.diff === 0 ? 'ok' : 'ng'}">
            <div class="cb-closed-head"><i data-lucide="${c.diff === 0 ? 'lock' : 'lock'}" class="w-4 h-4"></i>
                <b>${md(row.date)} は締め済み</b><span>${hm(c.at)} ${esc(c.by || '')}${c.op ? ` · 担当 ${esc(c.op)}` : ''}</span></div>
            <div class="cb-closed-grid">
                <div><span>帳簿上の残高</span><b>${yen(c.expected)}</b></div>
                <div><span>実査</span><b>${yen(c.counted)}</b></div>
                <div><span>過不足</span><b class="${diffCls(c.diff)}">${signed(c.diff)}</b></div>
                <div><span>現金売上（実際）</span><b>${yen(c.actualCash)}</b></div>
            </div>
            ${c.reason ? `<p class="cb-closed-reason">理由: ${esc(c.reason)}</p>` : ''}
            ${row.reopenCount ? `<p class="cb-closed-note">締め直し ${row.reopenCount}回（履歴は操作ログ）</p>` : ''}
            ${deposit ? `<div class="cb-deposit"><span>預け入れの目安 <b>${yen(deposit)}</b>（釣り銭準備金 ${yen(float)} を残す）</span>
                <button type="button" class="cb-link-btn" data-cb-action="deposit" data-amount="${deposit}">今日の出金として記録</button></div>` : ''}
        </div>
        ${reopenOpen ? `
        <div class="cb-inline-confirm">
            <input type="text" id="cb-reopen-reason" maxlength="200" placeholder="締めを取り消す理由（必須）例: 数え間違い" class="cb-input">
            <div class="cb-form-actions">
                <button type="button" class="input-clear-btn" data-cb-action="reopen-cancel">やめる</button>
                <button type="button" class="cb-danger-btn" data-cb-action="reopen">締めを取り消す</button>
            </div>
        </div>` : `<button type="button" class="cb-link-btn mt-3" data-cb-action="reopen-open"><i data-lucide="unlock" class="w-3.5 h-3.5"></i>締めを取り消す（理由が必要・ログに残ります）</button>`}`;
    } else {
        const counted = row.count?.total;
        const diff = row.diff;
        const earlier = (v.summary?.pendingDays || []).filter(d => d < row.date);
        el.innerHTML = `
        <div class="cb-close-box">
            ${counted === undefined || counted === null
                ? `<p class="cb-close-hint"><i data-lucide="info" class="w-4 h-4"></i>レジの現金を数えて「実査を保存」すると締められます</p>`
                : `<div class="cb-close-summary">
                    <div><span>帳簿上の残高</span><b>${row.expected === null ? '—' : yen(row.expected)}</b></div>
                    <div><span>実査（保存済み）</span><b>${yen(counted)}</b></div>
                    <div><span>過不足</span><b class="${diffCls(diff)}">${signed(diff)}</b></div>
                   </div>`}
            ${earlier.length ? `<p class="cb-close-warn">前日までに締めていない日があります（${earlier.slice(0, 4).map(md).join('・')}${earlier.length > 4 ? ' ほか' : ''}）。先に締めると繰越が確定します</p>` : ''}
            <label class="cb-field mt-3"><span>${diff ? '過不足の理由（必須）' : 'メモ（任意）'}</span>
                <input type="text" id="cb-close-reason" maxlength="200" placeholder="${diff ? '例: お釣りの渡し間違い・両替の記録漏れ など' : '例: 問題なし'}" class="cb-input"></label>
            <button type="button" id="cb-close-btn" class="btn-primary w-full py-3 mt-3" data-cb-action="close"${counted === undefined || counted === null || row.expected === null ? ' disabled' : ''}>
                ${md(row.date)} を締める</button>
            <p class="cb-close-foot">締めると、この日の記録は修正できなくなります（取り消しは理由つきでログに残ります）</p>
        </div>`;
        updateCloseButton();
    }
    if (window.lucide) lucide.createIcons({ nodes: [...el.querySelectorAll('[data-lucide]')] });
}

function updateCloseButton() {
    const row = rowOf(currentView());
    const btn = document.getElementById('cb-close-btn');
    if (!btn || !row || row.closed) return;
    const reason = document.getElementById('cb-close-reason')?.value.trim();
    const ready = row.count && row.expected !== null && (row.diff === 0 || !!reason);
    btn.disabled = !ready;
}

async function closeDay() {
    const row = rowOf(currentView());
    if (!row) return;
    if (countDraft?.dirty) { toast('数えた枚数が保存されていません。先に「実査を保存」してください', 'warn'); return; }
    const reason = document.getElementById('cb-close-reason')?.value.trim() || '';
    if (row.diff !== 0 && !reason) { toast('過不足の理由を入力してください', 'warn'); return; }
    const res = await act({ action: 'day.close', date: selectedDate, reason }, `${md(selectedDate)} を締めました`);
    if (res) { countDraft = null; render(); }
}

async function reopenDay() {
    const reason = document.getElementById('cb-reopen-reason')?.value.trim();
    if (!reason) { toast('締めを取り消す理由を入力してください', 'warn'); return; }
    const res = await act({ action: 'day.reopen', date: selectedDate, reason }, '締めを取り消しました（ログに残ります）');
    if (res) { reopenOpen = false; countDraft = null; render(); }
}

// ---- 月の出納帳 ----
function renderMonth(v) {
    const [y, m] = monthSel().split('-').map(Number);
    setText('cb-month-label', `${y}年${m}月 ・ ${shopName(v.shopId)}`);
    const tiles = document.getElementById('cb-month-tiles');
    const s = v.summary;
    if (tiles) {
        tiles.innerHTML = !s ? '' : [
            ['月初の繰越', s.opening === null ? '—' : yen(s.opening)],
            ['現金売上（SalonOne）', s.salonCash === null ? '—' : yen(s.salonCash)],
            ['入金（記録）', yen(s.manualIn)],
            ['出金（記録）', yen(s.manualOut)],
            [v.rows.some(r => r.date === todayStr()) ? '現在の残高' : '月末の残高', s.closing === null ? '—' : yen(s.closing)],
            ['過不足（締めた日）', `<span class="${diffCls(s.closedDays ? s.diffTotal : null)}">${s.closedDays ? signed(s.diffTotal) : '—'}</span><small>${s.closedDays}日締め${s.diffDays ? `・差異${s.diffDays}日` : ''}</small>`],
        ].map(([k, val]) => `<div class="cb-tile"><span>${k}</span><b>${val}</b></div>`).join('');
    }
    const body = document.getElementById('cb-month-body');
    if (body) {
        const rows = [...v.rows].reverse();
        body.innerHTML = rows.length ? rows.map(r => {
            const st = STATUS[r.status];
            const counted = r.closed ? r.closed.counted : r.count?.total;
            const flags = (r.salonChanged || r.openingChanged) ? ' <span class="cb-tag cb-tag-danger">締め後に変更</span>' : '';
            return `
            <tr class="${r.date === selectedDate ? 'cb-row-active' : ''}">
                <td><button type="button" class="cb-row-date" data-goto-date="${r.date}">${md(r.date)}<small>（${dowJa(r.date)}）</small></button></td>
                <td class="text-right tabular-nums" data-l="繰越">${r.opening === null ? '—' : yen(r.opening)}</td>
                <td class="text-right tabular-nums" data-l="現金売上">${r.salonCash === null ? '—' : yen(r.salonCash)}</td>
                <td class="text-right tabular-nums" data-l="入金">${r.manualIn ? yen(r.manualIn) : '—'}</td>
                <td class="text-right tabular-nums" data-l="出金">${r.manualOut ? yen(r.manualOut) : '—'}</td>
                <td class="text-right tabular-nums font-semibold" data-l="帳簿残高">${r.expected === null ? '—' : yen(r.expected)}</td>
                <td class="text-right tabular-nums" data-l="実査">${counted === undefined || counted === null ? '—' : yen(counted)}</td>
                <td class="text-right tabular-nums ${diffCls(counted === undefined || counted === null ? null : (r.closed ? r.closed.diff : r.diff))}" data-l="過不足">${counted === undefined || counted === null ? '—' : signed(r.closed ? r.closed.diff : r.diff)}</td>
                <td class="text-right"><span class="cb-chip ${st.cls}">${st.label}</span>${flags}</td>
            </tr>`;
        }).join('') : '<tr><td colspan="9" class="py-6 text-center text-surface-500">この月の記録はありません</td></tr>';
    }
    const cat = document.getElementById('cb-cat-summary');
    if (cat && s) {
        const items = Object.entries(s.byCat).map(([id, amt]) => ({ c: v.settings.categories.find(x => x.id === id) || { id, name: id, type: 'out' }, amt }))
            .sort((a, b) => (a.c.type === b.c.type ? b.amt - a.amt : a.c.type === 'in' ? -1 : 1));
        cat.innerHTML = items.length ? `<h4 class="cb-subhead">科目別の合計</h4><div class="cb-cats">${items.map(({ c, amt }) =>
            `<div class="cb-cat"><span><i class="cb-dot ${c.type}"></i>${esc(c.name)}</span><b class="tabular-nums">${c.type === 'in' ? '+' : '−'}${yen(amt)}</b></div>`).join('')}</div>` : '';
    }
}

// ---- 操作ログ ----
function describe(l, v) {
    const a = l.after || {}, b = l.before || {};
    const cat = id => esc(catName(v, id));
    const typeJa = t => (t === 'in' ? '入金' : '出金');
    switch (l.action) {
        case 'entry.add': return `${typeJa(a.type)}を記録: ${cat(a.cat)} ${yen(a.amount)}${a.memo ? `「${esc(a.memo)}」` : ''}${a.payee ? `（${esc(a.payee)}）` : ''}`;
        case 'entry.edit': {
            const parts = Object.keys(a).map(k => {
                const label = { date: '日付', type: '区分', cat: '科目', amount: '金額', memo: '摘要', payee: '支払先', receipt: '領収書' }[k] || k;
                const fmt = x => (k === 'amount' ? yen(x) : k === 'cat' ? cat(x) : k === 'type' ? typeJa(x) : k === 'receipt' ? (x ? 'あり' : 'なし') : k === 'date' ? md(x) : `「${esc(x || '')}」`);
                return `${label} ${fmt(b[k])} → ${fmt(a[k])}`;
            });
            return `記録を修正（${cat(l.cat)} ${yen(l.amount)}）: ${parts.join('、')}`;
        }
        case 'entry.void': return `記録を取消: ${typeJa(b.type)} ${cat(b.cat)} ${yen(b.amount)}${b.memo ? `「${esc(b.memo)}」` : ''}`;
        case 'count.save': return `実査を保存: ${yen(a.total)}${b.total !== undefined ? `（前回 ${yen(b.total)}）` : ''}`;
        case 'day.close': return `締め: 帳簿 ${yen(a.expected)} / 実査 ${yen(a.counted)} / 過不足 ${signed(a.diff)}（繰越 ${yen(a.opening)}・現金売上 ${yen(a.salonCash)}・入金 ${yen(a.manualIn)}・出金 ${yen(a.manualOut)}）`;
        case 'day.reopen': return `締めを取消（締め時: 実査 ${yen(b.counted)} / 過不足 ${signed(b.diff)}）`;
        case 'settings.shop': return `店舗の設定: ${b ? `開始日 ${esc(b.startDate || '—')} → ${esc(a.startDate)}、開始残高 ${yen(b.initialBalance)} → ${yen(a.initialBalance)}、準備金 ${yen(b.float)} → ${yen(a.float)}` : `開始 ${esc(a.startDate)}・開始残高 ${yen(a.initialBalance)}・準備金 ${yen(a.float)}`}`;
        case 'settings.brand': return '科目・現金とみなす支払い方法の設定を変更';
        default: return esc(l.action);
    }
}

const LOG_ICON = {
    'entry.add': 'plus-circle', 'entry.edit': 'pencil', 'entry.void': 'x-circle', 'count.save': 'coins',
    'day.close': 'lock', 'day.reopen': 'unlock', 'settings.shop': 'settings', 'settings.brand': 'settings',
};

function renderLog() {
    const v = currentView();
    const el = document.getElementById('cb-log');
    setText('cb-log-count', v ? `${v.logCount}件` : '');
    if (!el || !v) return;
    if (!v.log) { el.innerHTML = '<p class="cb-empty">開くと読み込みます</p>'; return; }
    const list = v.log.filter(LOG_FILTERS[logFilter] || LOG_FILTERS.all);
    el.innerHTML = list.length ? list.map(l => `
        <div class="cb-log-item ${l.action.replace('.', '-')}">
            <i data-lucide="${LOG_ICON[l.action] || 'dot'}" class="w-4 h-4 flex-shrink-0"></i>
            <div class="min-w-0">
                <div class="cb-log-head">${hm(l.at)} · ${esc(l.by || '')}${l.op ? ` · 担当 ${esc(l.op)}` : ''}${l.date ? ` <span class="cb-tag">対象 ${md(l.date)}</span>` : ''}</div>
                <div class="cb-log-body">${describe(l, v)}</div>
                ${l.reason ? `<div class="cb-log-reason">理由: ${esc(l.reason)}</div>` : ''}
            </div>
        </div>`).join('') : '<p class="cb-empty">該当する操作はありません</p>';
    if (window.lucide) lucide.createIcons({ nodes: [...el.querySelectorAll('[data-lucide]')] });
}

// ---- 設定 ----
function renderSettings(v) {
    const el = document.getElementById('cb-settings');
    if (!el) return;
    const sc = v.settings.shop || {};
    const canSetup = v.can?.setup;
    const canManage = v.can?.manage;
    const autoCash = !v.settings.cashMethodIds.length;
    const methods = v.methods || [];
    el.innerHTML = `
        <h4 class="cb-subhead">${esc(shopName(v.shopId))} の設定</h4>
        <div class="cb-form-grid">
            <label class="cb-field"><span>開始日</span>
                <input type="date" id="cb-s-start" value="${sc.startDate || ''}" max="${todayStr()}" class="cb-input"${canSetup ? '' : ' disabled'}></label>
            <label class="cb-field"><span>開始時の残高</span>
                <input type="number" id="cb-s-balance" inputmode="numeric" min="0" value="${sc.initialBalance ?? ''}" class="cb-input tabular-nums"${canSetup ? '' : ' disabled'}></label>
            <label class="cb-field"><span>釣り銭準備金</span>
                <input type="number" id="cb-s-float" inputmode="numeric" min="0" value="${sc.float ?? 0}" class="cb-input tabular-nums"></label>
        </div>
        ${canSetup ? '' : '<p class="cb-note">開始日・開始時の残高はオーナー・マネージャーのみ変更できます</p>'}
        <p class="cb-note">最終更新: ${sc.updatedAt ? `${hm(sc.updatedAt)} ${esc(sc.updatedBy || '')}` : '—'}</p>
        <button type="button" class="btn-secondary py-2 px-5 text-sm mt-2" data-cb-action="save-shop">店舗の設定を保存</button>
        ${canManage ? `
        <h4 class="cb-subhead mt-6">現金とみなす SalonOne の支払い方法（全店舗共通）</h4>
        <label class="cb-check"><input type="checkbox" id="cb-cash-auto"${autoCash ? ' checked' : ''}> 自動（名前に「現金」を含む支払い方法）</label>
        <div id="cb-cash-methods" class="cb-methods${autoCash ? ' hidden' : ''}">
            ${methods.length ? methods.map(m => `<label class="cb-check"><input type="checkbox" data-method="${m.id}"${v.settings.cashMethodIds.includes(m.id) || (autoCash && m.isCash) ? ' checked' : ''}> ${esc(m.name)} <small>ID ${m.id}</small></label>`).join('') : '<p class="cb-note">この月の売上に支払い方法の記録がありません</p>'}
        </div>
        <h4 class="cb-subhead mt-6">科目（全店舗共通）</h4>
        <p class="cb-note">名前の変更・使わない科目の非表示・追加ができます（過去の記録はそのまま残ります）</p>
        <div id="cb-cat-list" class="cb-cat-list">
            ${v.settings.categories.map(c => catRowHtml(c)).join('')}
        </div>
        <button type="button" class="cb-link-btn mt-2" data-cb-action="add-cat"><i data-lucide="plus" class="w-3.5 h-3.5"></i>科目を追加</button>
        <div><button type="button" class="btn-secondary py-2 px-5 text-sm mt-3" data-cb-action="save-brand">科目・支払い方法を保存</button></div>` : ''}`;
}

function catRowHtml(c) {
    return `<div class="cb-cat-row" data-cat-id="${c.id}" data-cat-type="${c.type}">
        <span class="cb-cat-type ${c.type}">${c.type === 'in' ? '入金' : '出金'}</span>
        <input type="text" maxlength="30" value="${esc(c.name)}" class="cb-input" data-cat-name aria-label="科目名">
        <label class="cb-check cb-cat-use"><input type="checkbox" data-cat-use${c.disabled ? '' : ' checked'}> 使う</label>
    </div>`;
}

function addCategoryRow() {
    const list = document.getElementById('cb-cat-list');
    if (!list) return;
    const type = confirm('入金の科目を追加しますか？\n（キャンセルで出金の科目を追加）') ? 'in' : 'out';
    list.insertAdjacentHTML('beforeend', catRowHtml({ id: 'new', type, name: '', disabled: false }));
    list.lastElementChild?.querySelector('[data-cat-name]')?.focus();
}

async function saveShopSettings() {
    const v = currentView();
    if (!v) return;
    const body = { action: 'settings.shop', month: monthSel(), float: Math.round(Number(document.getElementById('cb-s-float')?.value || 0)) };
    if (v.can?.setup) {
        body.startDate = document.getElementById('cb-s-start')?.value;
        body.initialBalance = Math.round(Number(document.getElementById('cb-s-balance')?.value || 0));
        const sc = v.settings.shop || {};
        if ((body.startDate !== sc.startDate || body.initialBalance !== sc.initialBalance)
            && !confirm('開始日・開始時の残高を変更すると、締めていない日の繰越が計算し直されます。変更しますか？')) return;
    }
    const res = await act(body, '店舗の設定を保存しました');
    if (res) render();
}

async function saveBrandSettings() {
    const categories = [...document.querySelectorAll('#cb-cat-list .cb-cat-row')].map(row => ({
        id: row.dataset.catId,
        type: row.dataset.catType,
        name: row.querySelector('[data-cat-name]')?.value.trim() || '',
        disabled: !row.querySelector('[data-cat-use]')?.checked,
    })).filter(c => c.id !== 'new' || c.name);
    const auto = document.getElementById('cb-cash-auto')?.checked;
    const cashMethodIds = auto ? [] : [...document.querySelectorAll('#cb-cash-methods input[data-method]:checked')].map(i => Number(i.dataset.method));
    if (!auto && !cashMethodIds.length) { toast('現金とみなす支払い方法を1つ以上選んでください', 'warn'); return; }
    const res = await act({ action: 'settings.brand', categories, cashMethodIds }, '科目・支払い方法を保存しました');
    if (res) refresh();
}

// ---- CSV（会計ソフト・税理士への提出用）----
function csvCell(v) {
    let s = v === null || v === undefined ? '' : String(v);
    // 表計算ソフトで数式として実行されないように（CSVインジェクション対策）
    if (typeof v === 'string' && /^[=+\-@\t\r]/.test(s)) s = `'${s}`;
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
function download(name, rows) {
    const csv = '﻿' + rows.map(r => r.map(csvCell).join(',')).join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

function exportBookCsv() {
    const v = currentView();
    if (!v?.rows) return;
    const out = [['日付', '区分', '科目', '摘要', '支払先・入金元', '入金', '出金', '残高', '記録した人', '担当者', '領収書', '状態']];
    let bal = null;
    v.rows.forEach((r, i) => {
        if (i === 0) {
            bal = r.opening;
            out.push([r.date, '繰越', '', '前月からの繰越', '', '', '', bal ?? '', '', '', '', '']);
        } else if (bal !== r.opening) {
            bal = r.opening; // 締めた日は実査額が繰越になる
        }
        if (r.salonCash) {
            bal = bal === null ? null : bal + r.salonCash;
            out.push([r.date, '入金', '現金売上（SalonOne）', (r.cashMethods || []).map(m => m.name).join('・'), '', r.salonCash, '', bal ?? '', 'SalonOne', '', '', '']);
        }
        for (const e of v.entries.filter(x => x.date === r.date)) {
            if (e.voided) {
                out.push([e.date, e.type === 'in' ? '入金' : '出金', catName(v, e.cat), `（取消）${e.memo || ''} 理由: ${e.voided.reason}`, e.payee || '', '', '', bal ?? '', e.createdBy, e.op || '', e.receipt ? 'あり' : '', '取消']);
                continue;
            }
            bal = bal === null ? null : bal + (e.type === 'in' ? e.amount : -e.amount);
            out.push([e.date, e.type === 'in' ? '入金' : '出金', catName(v, e.cat), e.memo || '', e.payee || '', e.type === 'in' ? e.amount : '', e.type === 'out' ? e.amount : '', bal ?? '', e.createdBy, e.op || '', e.receipt ? 'あり' : '', e.rev > 1 ? '修正あり' : '']);
        }
        if (r.closed && r.closed.diff !== 0) {
            bal = r.closed.counted;
            out.push([r.date, r.closed.diff > 0 ? '入金' : '出金', '現金過不足', r.closed.reason || '', '', r.closed.diff > 0 ? r.closed.diff : '', r.closed.diff < 0 ? -r.closed.diff : '', bal, r.closed.by, r.closed.op || '', '', '締め']);
        } else if (r.closed) {
            bal = r.closed.counted;
        }
    });
    download(`出納帳_${shopName(v.shopId)}_${v.month}.csv`, out);
}

async function exportLogCsv() {
    const shopId = activeShopId();
    if (!shopId) return;
    let v = currentView();
    if (!v?.log || v.logStale) v = await loadCashbook(shopId, monthSel(), { log: true }).catch(() => v);
    if (!v?.log) return;
    const out = [['日時', '操作した人', '担当者', '操作', '対象日', '内容', '理由']];
    const plain = s => String(s).replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
    for (const l of v.log) {
        out.push([fullTime(l.at), l.by || '', l.op || '', l.action, l.date || '', plain(describe(l, v)), l.reason || '']);
    }
    download(`出納帳_操作ログ_${shopName(shopId)}_${v.month}.csv`, out);
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}
