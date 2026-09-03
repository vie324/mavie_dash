// 日報入力タブ: SalonOne APIにない項目の手入力
//   - 日次: ブログ更新・SNS更新・★5口コミ（スタッフ本人 or 管理者/店舗が入力）
//   - 月次: 物販売上（管理者/店舗のみ・インセンティブ計算用）
//   - 広告費: 媒体別の手入力（管理者のみ・APIに広告費がない媒体用）

import { state, on, isAdmin, isAdminLike, isStaffLocked, isStoreLocked, currentShopId, staffsOfShop, staffName } from '../core/state.js';
import { esc, num, yen, todayStr } from '../core/format.js';
import { loadManual, saveManualPatch, getManual, monthlyTotalsByStaff } from '../data/manual.js';
import { toast } from '../core/engage.js';

const BLOG_TARGET = 10; // 月間ブログ更新目標（従来ツールの値を踏襲）

let currentDate = todayStr();

function monthOf(date) {
    return date.slice(0, 7);
}

export function init() {
    on('tab:shown', id => { if (id === 'input') refresh(); });
    on('data:manual', render);
    on('masters', renderStaffOptions);

    document.getElementById('input-date')?.addEventListener('change', ev => {
        currentDate = ev.target.value || todayStr();
        refresh();
    });
    document.getElementById('input-staff')?.addEventListener('change', fillDailyForm);
    document.getElementById('input-daily-save')?.addEventListener('click', saveDaily);

    document.getElementById('input-monthly-body')?.addEventListener('change', ev => {
        const input = ev.target.closest('input[data-staff-id]');
        if (input) saveMonthly(input);
    });
    document.getElementById('input-adcost-body')?.addEventListener('change', ev => {
        const input = ev.target.closest('input[data-source-id]');
        if (input) saveAdCost(input);
    });
}

async function refresh() {
    const dateInput = document.getElementById('input-date');
    if (dateInput && !dateInput.value) dateInput.value = currentDate;
    if (dateInput) dateInput.max = todayStr();
    renderStaffOptions();
    await loadManual(monthOf(currentDate));
    render();
}

function inputStaffs() {
    if (isStaffLocked()) return state.masters.staffs.filter(s => String(s.id) === String(state.session.staffId));
    const shopId = currentShopId();
    return shopId === 'all' ? state.masters.staffs : staffsOfShop(shopId);
}

function renderStaffOptions() {
    const sel = document.getElementById('input-staff');
    if (!sel) return;
    const staffs = inputStaffs();
    const prev = sel.value;
    sel.innerHTML = staffs.map(s => `<option value="${s.id}">${esc(s.name)}</option>`).join('');
    if (staffs.some(s => String(s.id) === prev)) sel.value = prev;
    sel.disabled = isStaffLocked();
    fillDailyForm();
}

function selectedStaffId() {
    if (isStaffLocked()) return String(state.session.staffId);
    return document.getElementById('input-staff')?.value || '';
}

function fillDailyForm() {
    const data = getManual(monthOf(currentDate));
    const entry = data.daily[`${currentDate}:${selectedStaffId()}`] || {};
    setValue('input-next-new', entry.nextNew);
    setValue('input-next-repeat', entry.nextRepeat);
    setValue('input-blog', entry.blog);
    setValue('input-sns', entry.sns);
    setValue('input-reviews', entry.reviews);
}

async function saveDaily() {
    const staffId = selectedStaffId();
    if (!staffId) return;
    const btn = document.getElementById('input-daily-save');
    btn.disabled = true;
    try {
        const entry = {
            nextNew: numValue('input-next-new'),
            nextRepeat: numValue('input-next-repeat'),
            blog: numValue('input-blog'),
            sns: numValue('input-sns'),
            reviews: numValue('input-reviews'),
        };
        const res = await saveManualPatch(monthOf(currentDate), {
            daily: { [`${currentDate}:${staffId}`]: entry },
        });
        toast(res.storage === 'local' ? '保存しました（この端末のみ）' : '保存しました');
    } catch (e) {
        console.error('daily save', e);
        toast('保存に失敗しました。時間をおいて再度お試しください', 'error');
    } finally {
        btn.disabled = false;
    }
}

async function saveMonthly(input) {
    const v = input.value === '' ? null : Number(input.value);
    if (v !== null && (!isFinite(v) || v < 0)) return;
    try {
        await saveManualPatch(monthOf(currentDate), {
            monthly: { [input.dataset.staffId]: v === null ? null : { productSales: v } },
        });
        toast('物販売上を保存しました');
    } catch (e) {
        console.error('monthly save', e);
        toast('保存に失敗しました', 'error');
    }
}

async function saveAdCost(input) {
    const v = input.value === '' ? null : Number(input.value);
    if (v !== null && (!isFinite(v) || v < 0)) return;
    try {
        await saveManualPatch(monthOf(currentDate), { adCosts: { [input.dataset.sourceId]: v } });
        toast('広告費を保存しました');
    } catch (e) {
        console.error('adcost save', e);
        toast('保存に失敗しました', 'error');
    }
}

function render() {
    const month = monthOf(currentDate);
    const data = getManual(month);

    // ストレージ注意書き
    const notice = document.getElementById('input-storage-notice');
    if (notice) notice.classList.toggle('hidden', state.manualStorage !== 'local');

    fillDailyForm();

    // 月間サマリ（スタッフ別合計）
    const totals = monthlyTotalsByStaff(month);
    const sumBody = document.getElementById('input-summary-body');
    if (sumBody) {
        const staffs = inputStaffs();
        const [y, m] = month.split('-');
        setText('input-summary-month', `${y}年${Number(m)}月の入力合計`);
        // 次回予約率の分母（来店数）はSalonOneの当月スタッフ別実績
        const visitsOf = id => {
            const r = (state.data.nowMonth?.by_staff || []).find(x => String(x.staff_id) === String(id));
            return r ? (r.new_visit_count || 0) + (r.repeat_visit_count || 0) : 0;
        };
        sumBody.innerHTML = staffs.map(s => {
            const t = totals[String(s.id)] || { blog: 0, sns: 0, reviews: 0, nextNew: 0, nextRepeat: 0 };
            const done = t.blog >= BLOG_TARGET;
            const visits = visitsOf(s.id);
            const rate = visits > 0 ? (t.nextNew + t.nextRepeat) / visits * 100 : null;
            return `<tr class="border-b border-surface-100 dark:border-accent-800">
                <td class="py-2 px-3 font-medium">${esc(s.name)}</td>
                <td class="py-2 px-3 text-right tabular-nums text-primary-600 font-semibold">${num(t.nextNew)}</td>
                <td class="py-2 px-3 text-right tabular-nums">${num(t.nextRepeat)}</td>
                <td class="py-2 px-3 text-right tabular-nums">${rate === null ? '—' : rate.toFixed(0) + '%'}<span class="text-[10px] text-surface-400"> /${num(visits)}名</span></td>
                <td class="py-2 px-3 text-right tabular-nums ${done ? 'text-sage-600 font-semibold' : ''}">${num(t.blog)} / ${BLOG_TARGET}${done ? ' ✅' : ''}</td>
                <td class="py-2 px-3 text-right tabular-nums">${num(t.sns)}</td>
                <td class="py-2 px-3 text-right tabular-nums">${num(t.reviews)}</td>
            </tr>`;
        }).join('') || '<tr><td colspan="7" class="py-6 text-center text-surface-500">スタッフがいません</td></tr>';
    }

    // 月次: 物販売上（管理者・店舗）
    const monthlySection = document.getElementById('input-monthly-section');
    if (monthlySection) {
        const show = isAdminLike() || isStoreLocked();
        monthlySection.classList.toggle('hidden', !show);
        if (show) {
            const body = document.getElementById('input-monthly-body');
            const staffs = inputStaffs();
            body.innerHTML = staffs.map(s => {
                const cur = data.monthly[String(s.id)]?.productSales;
                return `<tr class="border-b border-surface-100 dark:border-accent-800">
                    <td class="py-2 px-3 font-medium">${esc(s.name)}</td>
                    <td class="py-2 px-3 text-right">
                        <input type="number" inputmode="numeric" min="0" step="1000" value="${cur ?? ''}" placeholder="未入力"
                            data-staff-id="${s.id}"
                            class="w-32 text-right px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                    </td>
                </tr>`;
            }).join('');
        }
    }

    // 広告費（管理者のみ）
    const adSection = document.getElementById('input-adcost-section');
    if (adSection) {
        adSection.classList.toggle('hidden', !isAdminLike());
        if (isAdminLike()) {
            const body = document.getElementById('input-adcost-body');
            const sources = [...state.masters.visitSources, { id: 'other', name: 'その他' }];
            body.innerHTML = sources.map(src => `
                <tr class="border-b border-surface-100 dark:border-accent-800">
                    <td class="py-2 px-3 font-medium">${esc(src.name || '未設定')}${src.platform_type ? ' <span class="text-[10px] text-surface-400">(API広告費あり・上書き不要)</span>' : ''}</td>
                    <td class="py-2 px-3 text-right">
                        <input type="number" inputmode="numeric" min="0" step="1000" value="${data.adCosts[String(src.id)] ?? ''}" placeholder="未入力"
                            data-source-id="${src.id}"
                            class="w-32 text-right px-2 py-1.5 text-sm border border-surface-300 dark:border-gray-600 rounded-lg bg-white dark:bg-gray-700 dark:text-white tabular-nums">
                    </td>
                </tr>`).join('');
        }
    }
}

// ---- util ----
function setValue(id, v) {
    const el = document.getElementById(id);
    if (el) el.value = v ?? '';
}
function numValue(id) {
    const el = document.getElementById(id);
    const v = el?.value;
    if (v === '' || v === undefined) return null;
    const n = Number(v);
    return isFinite(n) && n >= 0 ? Math.round(n) : null;
}
function setText(id, text) {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
}

export { BLOG_TARGET };
