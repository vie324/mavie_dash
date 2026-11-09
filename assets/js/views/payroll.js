// インセンティブタブ: 2026年11月分〜の新しい給与・インセンティブ制度（オーナー専用）
//   スタッフごとに 給与パターン・開店締め（◯月から適用）と、その月の勤怠（残業・法定時間外・総労働時間）を入力し、
//   基本給 / 固定残業代 / 税込施術売上 / 到達した率 / インセンティブ / 次の段階 / 物販手当 / 残業の追加支給 / 総支給 を表示する。
//   計算は data/payroll.js → payroll/calc.js（金額・率は payroll/config.js）。

import { state, on, currentShopId, staffsOfShop } from '../core/state.js';
import { yen, esc, monthLabel } from '../core/format.js';
import { toast } from '../core/engage.js';
import { getPattern } from '../payroll/calc.js';
import {
    computeStaffPay, payrollLoaded, payrollStorage, payrollError, staffSettingHistory, staffHours,
    payrollSalesMap, saveStaffSetting, saveStaffHours,
} from '../data/payroll.js';

const HOURS_FIELDS = [
    ['overtime', '残業時間', '所定労働時間を超えた時間'],
    ['statutoryOvertime', '法定時間外', '1日8h・週44hを超えた時間'],
    ['totalHours', '総労働時間', 'その月に働いた時間の合計'],
];

let isActive = () => false;

export function init(opts = {}) {
    if (opts.isActive) isActive = opts.isActive;
    on('data:payroll', () => { if (isActive()) render(); });
    const body = document.getElementById('payroll-body');
    if (!body) return;
    body.addEventListener('change', onChange);
    body.addEventListener('click', onClick);
    body.addEventListener('input', ev => {
        const input = ev.target.closest('input[data-hours-field]');
        if (!input) return;
        input.dataset.dirty = '1';
        const btn = input.closest('[data-staff-card]')?.querySelector('[data-action="save-hours"]');
        if (btn) btn.textContent = '勤怠を保存（未保存）';
    });
}

function anchorMonthKey() {
    return `${state.filters.anchor.y}-${String(state.filters.anchor.m).padStart(2, '0')}`;
}

function fromLabel(monthKey) {
    const [y, m] = monthKey.split('-').map(Number);
    return `${y}年${m}月`;
}

function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

// 表示するスタッフ: 店舗フィルタ内のSalonOneのスタッフ ＋ その月に売上があるスタッフ（退職者を含む）
// 売上の金額は computeStaffPay がサーバーの全店舗の合計を使う（row は取れなかったときの代わり）
function staffList(month) {
    const map = new Map();
    for (const s of staffsOfShop(currentShopId())) map.set(String(s.id), { id: s.id, name: s.name, row: null });
    for (const r of state.data.summary?.by_staff || []) {
        if (r.staff_id === null || r.staff_id === undefined) continue;
        const key = String(r.staff_id);
        const cur = map.get(key) || { id: r.staff_id, name: r.staff_name, row: null };
        cur.row = r;
        if (!cur.name) cur.name = r.staff_name;
        map.set(key, cur);
    }
    if (currentShopId() === 'all') {
        for (const [key, r] of Object.entries(payrollSalesMap(month) || {})) {
            if (!map.has(key)) map.set(key, { id: key, name: r.staff_name, row: null });
        }
    }
    return [...map.values()];
}

export function render() {
    const body = document.getElementById('payroll-body');
    if (!body || !state.data.summary) return;
    const month = anchorMonthKey();
    setText('payroll-period', `対象月: ${monthLabel(state.filters.anchor)}`);

    if (state.filters.periodKind !== 'month') {
        body.innerHTML = '<p class="py-8 text-center text-sm text-surface-500">給与は月単位で計算します。期間フィルタを「今月」にして対象月を選択してください。</p>';
        return;
    }
    if (!payrollLoaded(month)) {
        const err = payrollError();
        body.innerHTML = err
            ? '<p class="py-8 text-center text-sm text-rose-500">給与の設定を読み込めませんでした。時間をおいて「更新」してください。</p>'
            : '<p class="py-8 text-center text-sm text-surface-500 animate-pulse">給与の設定を読み込んでいます…</p>';
        return;
    }

    const preserved = captureDirtyInputs(body);
    // 「未設定のスタッフ」一覧の開閉は描き直しても保つ（続けて設定しやすいように）。初回は全員未設定なら開く
    const prevDetails = body.dataset.month === month ? body.querySelector('details[data-unset]') : null;
    const canEdit = payrollStorage(month) === 'kv';
    const results = staffList(month).map(s => ({ ...s, r: computeStaffPay(month, s.id, s.row || {}) }));
    const configured = results.filter(x => x.r?.pay).sort((a, b) => b.r.pay.total - a.r.pay.total);
    const unset = results.filter(x => !x.r?.pay).sort((a, b) => (b.r?.sales.service || 0) - (a.r?.sales.service || 0));

    const sum = f => configured.reduce((a, x) => a + f(x.r.pay), 0);
    const noHours = configured.filter(x => !x.r.pay.hoursEntered).length;
    const rule = configured[0]?.r.rule || results[0]?.r?.rule;

    const html = [];
    if (!payrollSalesMap(month)) {
        html.push(`<div class="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-700/40 rounded-xl p-3 text-xs text-amber-800 dark:text-amber-200 mb-3">
            SalonOneから全店舗の売上を取得できなかったため、画面の絞り込みの売上で計算しています。2店舗で働くスタッフの売上が欠けることがあるので、「更新」してから確認してください。</div>`);
    }
    if (!canEdit) {
        html.push(`<div class="bg-amber-50 dark:bg-amber-900/20 border border-amber-200 dark:border-amber-700/40 rounded-xl p-3 text-xs text-amber-800 dark:text-amber-200">
            給与パターン・勤怠の保存にはサーバー保存（Supabase）の設定が必要です（docs/SALONONE_INTEGRATION.md 参照）。</div>`);
    }
    html.push(`
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3">
            ${summaryCell('計算対象', `${configured.length}名`, unset.length ? `パターン未設定 ${unset.length}名` : '')}
            ${summaryCell('総支給の合計', yen(sum(p => p.total)), noHours ? `勤怠未入力 ${noHours}名（残業精算前）` : '残業精算込み')}
            ${summaryCell('インセンティブの合計', yen(sum(p => p.incentive)), '税込施術売上から')}
            ${summaryCell('物販手当の合計', yen(sum(p => p.retailAllowance)), '税込物販売上から')}
        </div>`);

    if (configured.length) {
        html.push(`<div class="grid grid-cols-1 xl:grid-cols-2 gap-4 mt-4">${configured.map(x => staffCard(x, month, canEdit)).join('')}</div>`);
    } else {
        html.push('<p class="py-6 text-center text-sm text-surface-500">給与パターンが設定されたスタッフがいません。下の一覧から社員の給与パターンを設定してください。</p>');
    }

    if (unset.length) {
        html.push(`
            <details data-unset class="mt-5 rounded-xl border border-surface-200 dark:border-accent-700 p-3" ${(prevDetails ? prevDetails.open : !configured.length) ? 'open' : ''}>
                <summary class="cursor-pointer text-sm font-semibold text-accent-900">給与パターン未設定のスタッフ（${unset.length}名・計算の対象外）</summary>
                <p class="text-xs text-surface-500 mt-2">社員の方はパターンを選ぶと、この月（${fromLabel(month)}）から計算されます。パート・業務委託の方は未設定のままにしてください。</p>
                <div class="mt-2 divide-y divide-surface-100 dark:divide-accent-800">
                    ${unset.map(x => unsetRow(x, month, rule, canEdit)).join('')}
                </div>
            </details>`);
    }

    body.dataset.month = month;
    body.innerHTML = html.join('');
    restoreDirtyInputs(body, preserved);
    if (window.lucide) lucide.createIcons();
}

function summaryCell(label, value, sub) {
    return `
        <div class="bg-surface-50 dark:bg-gray-700/40 rounded-xl p-3">
            <p class="text-[10px] uppercase tracking-wider text-surface-500 mb-1">${label}</p>
            <p class="text-lg font-display font-bold text-accent-900 tabular-nums">${value}</p>
            ${sub ? `<p class="text-[10px] text-surface-500 mt-0.5">${esc(sub)}</p>` : ''}
        </div>`;
}

function patternOptions(rule, selected) {
    const opts = ['<option value="">未設定（対象外）</option>'];
    for (const [key, p] of Object.entries(rule?.patterns || {})) {
        opts.push(`<option value="${esc(key)}" ${key === selected ? 'selected' : ''}>${esc(p.label)}</option>`);
    }
    return opts.join('');
}

function line(label, value, { strong = false, muted = false, full = false } = {}) {
    return `
        <div class="${full ? 'sm:col-span-2' : ''} flex items-baseline justify-between gap-3 py-1 border-b border-surface-100 dark:border-accent-800">
            <dt class="text-xs text-surface-500 whitespace-nowrap flex-shrink-0">${label}</dt>
            <dd class="tabular-nums text-right ${strong ? 'font-bold text-accent-900' : ''} ${muted ? 'text-surface-400' : ''}">${value}</dd>
        </div>`;
}

function staffCard({ id, name, r }, month, canEdit) {
    const { rule, setting, sales, pay } = r;
    const pattern = getPattern(rule, setting.pattern);
    const hours = staffHours(id, month) || {};
    const ot = pay.overtime;
    const disabled = canEdit ? '' : 'disabled';
    const later = staffSettingHistory(id).find(e => e.from > month);

    const notes = [];
    if (setting.from !== month) notes.push(`この設定は${fromLabel(setting.from)}から適用中です。変更すると${fromLabel(month)}以降に反映され、前の月は変わりません。`);
    if (later) notes.push(`${fromLabel(later.from)}から別の設定があります。`);
    if (sales.retailSource === 'manual') notes.push('物販売上は日報入力タブの手入力値を使っています。');
    if (sales.refund > 0) notes.push(`SalonOneで返金 ${yen(sales.refund)} があります（SalonOneの会計済み売上をそのまま使っています）。`);

    const nextText = pay.next ? esc(pay.nextLabel) : '<span class="text-surface-400">次の段階なし</span>';
    const retailNext = pay.retailNext && pay.retailSales > 0
        ? `<span class="block text-[10px] text-surface-500">${yen(pay.retailNext.threshold)}超で${pay.retailNext.rate}%（あと${yen(pay.retailNext.remaining)}）</span>` : '';
    const otValue = pay.hoursEntered
        ? `${yen(pay.overtimeTopUp)}<span class="block text-[10px] text-surface-500">基本給部分 ${yen(Math.round(ot.basePart))} ＋ インセンティブ部分 ${yen(Math.round(ot.incentivePart))} − 固定残業代 ${yen(ot.fixedOtPay)}</span>`
        : '<span class="text-amber-600 dark:text-amber-300">勤怠未入力（0円で計算）</span>';

    return `
    <div class="rounded-2xl border border-surface-200 dark:border-accent-700 p-4 bg-white/60 dark:bg-gray-800/40" data-staff-card data-staff-id="${esc(id)}">
        <div class="flex items-start justify-between gap-3">
            <div class="min-w-0">
                <p class="font-bold text-accent-900 truncate">${esc(name || `スタッフ${id}`)}</p>
                <p class="text-[11px] text-surface-500">${esc(pattern.label)}・${esc(pattern.hours)}</p>
            </div>
            <div class="text-right flex-shrink-0">
                <p class="text-[10px] text-surface-500">総支給</p>
                <p class="text-xl font-display font-bold text-accent-900 tabular-nums">${yen(pay.total)}</p>
                ${pay.hoursEntered ? '' : '<p class="text-[10px] text-amber-600 dark:text-amber-300">残業精算前</p>'}
            </div>
        </div>

        <div class="grid grid-cols-2 gap-2 mt-3">
            <label class="text-[11px] text-surface-500">給与パターン
                <select data-action="pattern" ${disabled} class="mt-0.5 w-full px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white">${patternOptions(rule, setting.pattern)}</select>
            </label>
            <label class="text-[11px] text-surface-500">開店・締め作業
                <select data-action="duty" ${disabled} class="mt-0.5 w-full px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white">
                    <option value="1" ${pay.storeDuty ? 'selected' : ''}>担う${pattern.defaultStoreDuty ? '（標準）' : '（率+1%）'}</option>
                    <option value="0" ${pay.storeDuty ? '' : 'selected'}>担わない${pattern.defaultStoreDuty ? '（率−1%）' : '（標準）'}</option>
                </select>
            </label>
        </div>

        <dl class="grid grid-cols-1 sm:grid-cols-2 gap-x-4 mt-3 text-sm">
            ${line('基本給', yen(pay.baseSalary))}
            ${line(`固定残業代（${pay.fixedOtHours}h分）`, yen(pay.fixedOtPay))}
            ${line('税込施術売上', yen(pay.serviceSales))}
            ${line('到達した率', pay.rate > 0 ? `${pay.rate}%` : `0%<span class="block text-[10px] text-surface-500">${yen(pay.startLine)}超から</span>`, { muted: pay.rate === 0 })}
            ${line('インセンティブ', yen(pay.incentive), { strong: true, muted: pay.incentive === 0 })}
            ${line('次の段階', nextText)}
            ${line('税込物販売上', yen(pay.retailSales), { muted: pay.retailSales === 0 })}
            ${line(`物販手当${pay.retailRate ? `（${pay.retailRate}%）` : ''}`, `${yen(pay.retailAllowance)}${retailNext}`, { muted: pay.retailAllowance === 0 })}
            ${line('残業の追加支給', otValue, { full: true })}
        </dl>

        <div class="mt-3 grid grid-cols-3 gap-2">
            ${HOURS_FIELDS.map(([f, label, hint]) => `
                <label class="text-[11px] text-surface-500" title="${esc(hint)}">${label}
                    <input type="number" inputmode="decimal" min="0" max="744" step="0.01" ${disabled}
                        data-hours-field="${f}" value="${hours[f] ?? ''}" placeholder="未入力" aria-label="${esc(name || '')}の${label}（時間）"
                        class="mt-0.5 w-full text-right px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                    <span class="block text-[10px] text-surface-400 leading-tight mt-0.5">${hint}</span>
                </label>`).join('')}
        </div>
        <div class="mt-2 flex justify-end">
            <button type="button" data-action="save-hours" ${disabled} class="chip chip-gold text-xs">勤怠を保存</button>
        </div>
        ${notes.length ? `<ul class="mt-2 text-[11px] text-surface-500 list-disc list-inside space-y-0.5">${notes.map(n => `<li>${esc(n)}</li>`).join('')}</ul>` : ''}
    </div>`;
}

function unsetRow({ id, name, r }, month, rule, canEdit) {
    const sales = r?.sales;
    return `
        <div class="flex items-center justify-between gap-3 py-2" data-staff-card data-staff-id="${esc(id)}">
            <div class="min-w-0">
                <p class="text-sm font-medium truncate">${esc(name || `スタッフ${id}`)}</p>
                <p class="text-[11px] text-surface-500 tabular-nums">税込施術売上 ${yen(sales?.service || 0)}</p>
            </div>
            <select data-action="pattern" ${canEdit ? '' : 'disabled'} aria-label="${esc(name || '')}の給与パターン"
                class="px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white">${patternOptions(rule, '')}</select>
        </div>`;
}

// ---- 入力イベント ----
function cardOf(el) {
    const card = el.closest('[data-staff-card]');
    return card ? { card, staffId: card.dataset.staffId } : null;
}

async function onChange(ev) {
    const select = ev.target.closest('select[data-action]');
    if (!select) return;
    const c = cardOf(select);
    const month = document.getElementById('payroll-body')?.dataset.month;
    if (!c || !month) return;
    const r = computeStaffPay(month, c.staffId, null);
    let setting;
    if (select.dataset.action === 'pattern') {
        const key = select.value || null;
        // パターンを変えたら、開店・締めはそのパターンの標準に戻す
        setting = { pattern: key, storeDuty: key ? getPattern(r.rule, key)?.defaultStoreDuty : undefined };
    } else if (select.dataset.action === 'duty') {
        const cur = r?.setting;
        if (!cur) return;
        setting = { pattern: cur.pattern, storeDuty: select.value === '1' };
    } else {
        return;
    }
    select.disabled = true;
    try {
        await saveStaffSetting(month, c.staffId, setting);
        toast(setting.pattern ? `給与の設定を保存しました（${fromLabel(month)}から）` : `${fromLabel(month)}から計算の対象外にしました`, 'success');
    } catch (e) {
        console.error('payroll setting save', e);
        toast(e?.body?.detail || '給与の設定の保存に失敗しました', 'error');
        render();
    }
}

async function onClick(ev) {
    const btn = ev.target.closest('button[data-action="save-hours"]');
    if (!btn) return;
    const c = cardOf(btn);
    const month = document.getElementById('payroll-body')?.dataset.month;
    if (!c || !month) return;
    const hours = {};
    for (const input of c.card.querySelectorAll('input[data-hours-field]')) {
        const v = input.value.trim();
        if (v !== '' && (!isFinite(Number(v)) || Number(v) < 0 || Number(v) > 744)) {
            toast('時間は0〜744の数字で入力してください', 'error');
            input.focus();
            return;
        }
        hours[input.dataset.hoursField] = v;
    }
    btn.disabled = true;
    try {
        for (const input of c.card.querySelectorAll('input[data-hours-field]')) delete input.dataset.dirty;
        await saveStaffHours(month, c.staffId, hours);
        toast('勤怠を保存しました', 'success');
    } catch (e) {
        console.error('payroll hours save', e);
        toast(e?.body?.detail || '勤怠の保存に失敗しました', 'error');
        btn.disabled = false;
    }
}

// 自動更新（5分ごと）などで描き直しても、保存前の入力とカーソル位置を消さない
function captureDirtyInputs(body) {
    const dirty = [];
    for (const input of body.querySelectorAll('input[data-hours-field][data-dirty="1"]')) {
        dirty.push({ staffId: input.closest('[data-staff-card]')?.dataset.staffId, field: input.dataset.hoursField, value: input.value });
    }
    const a = document.activeElement;
    const focus = a && body.contains(a) && a.dataset.hoursField
        ? { staffId: a.closest('[data-staff-card]')?.dataset.staffId, field: a.dataset.hoursField } : null;
    return { dirty, focus, month: body.dataset.month };
}

function restoreDirtyInputs(body, { dirty, focus, month }) {
    if (month !== body.dataset.month) return; // 対象月が変わったら引き継がない
    const find = (staffId, field) => body.querySelector(`[data-staff-card][data-staff-id="${CSS.escape(String(staffId))}"] input[data-hours-field="${field}"]`);
    for (const d of dirty) {
        const input = find(d.staffId, d.field);
        if (!input) continue;
        input.value = d.value;
        input.dataset.dirty = '1';
        const btn = input.closest('[data-staff-card]')?.querySelector('[data-action="save-hours"]');
        if (btn) btn.textContent = '勤怠を保存（未保存）';
    }
    if (focus) find(focus.staffId, focus.field)?.focus();
}
