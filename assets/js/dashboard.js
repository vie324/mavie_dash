// ★ スプレッドシートAPI設定
const SPREADSHEET_API_KEY = 'mavie_spreadsheet_api_url';

// デフォルトのAPI URL（コードに埋め込み）
const DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbwrz7LgQb2uH9VTtmalZxIEcJnHc-Ae53UNMnDi0MM5eLdP7XmZKOlDPTaOL5pmsFwf/exec';

// API URL（デフォルトURLを使用）
let API_URL = DEFAULT_API_URL;

// スタッフ名簿（スプレッドシートから読み込み、未設定時はデフォルト値）
let STAFF_ROSTER = {
    honatsugi: ['haruka', 'vienna'],
    chiba: ['kiki', 'karin', 'nanami', 'kanon', 'ayami', 'miki'],
    yamato: ['amano']
};

// ===== スタッフ別カラーパレット =====
// 上品かつ識別しやすい8色。名前ハッシュで安定割り当て。
const STAFF_COLOR_PALETTE = [
    { main: '#b8956a', light: '#f0e5d5', dark: '#866644' }, // Champagne Gold
    { main: '#566882', light: '#dce1eb', dark: '#3d4859' }, // Slate Blue
    { main: '#739977', light: '#dde9de', dark: '#5d7d60' }, // Sage
    { main: '#b08f8a', light: '#ecdddb', dark: '#9a7873' }, // Dusty Rose
    { main: '#c9a96e', light: '#f0e3c6', dark: '#a08754' }, // Warm Gold
    { main: '#8b7aa1', light: '#e3ddeb', dark: '#6b5d7e' }, // Lavender Grey
    { main: '#7ea5a8', light: '#d5e4e6', dark: '#5d8386' }, // Muted Teal
    { main: '#c08562', light: '#ecd4c1', dark: '#9a6947' }, // Terracotta
];

function _hashStaffName(name) {
    const s = String(name || '').toLowerCase();
    let h = 0;
    for (let i = 0; i < s.length; i++) {
        h = (h << 5) - h + s.charCodeAt(i);
        h |= 0;
    }
    return Math.abs(h);
}

function getStaffColor(name) {
    if (!name) return STAFF_COLOR_PALETTE[0];
    return STAFF_COLOR_PALETTE[_hashStaffName(name) % STAFF_COLOR_PALETTE.length];
}

// スタッフ頭文字アバター（カラー付き円）
function renderStaffAvatar(name, size = 32) {
    const c = getStaffColor(name);
    const initial = String(name || '?').trim().charAt(0).toUpperCase();
    return `<span class="staff-avatar" style="width:${size}px;height:${size}px;background:linear-gradient(135deg, ${c.main}, ${c.dark});" aria-hidden="true">${initial}</span>`;
}

// ===== エンプティステート =====
function renderEmptyState({ icon = 'inbox', title = 'データがありません', desc = '', action = '', colSpan = false } = {}) {
    const wrapClass = colSpan ? 'empty-state col-span-full' : 'empty-state';
    return `
        <div class="${wrapClass}">
            <div class="empty-state-icon"><i data-lucide="${icon}"></i></div>
            <div class="empty-state-title">${title}</div>
            ${desc ? `<p class="empty-state-desc">${desc}</p>` : ''}
            ${action || ''}
        </div>
    `;
}
const DEFAULT_STAFF_ROSTER = {
    honatsugi: ['haruka', 'vienna'],
    chiba: ['kiki', 'karin', 'nanami', 'kanon', 'ayami', 'miki'],
    yamato: ['amano']
};

// Premium Color Palette - Elegant & Harmonious
const BrandColors = {
    // Primary - Slate Blue (落ち着いたスレートブルー)
    primary: '#47566b',
    primaryLight: '#6e819c',
    primaryDark: '#3d4859',
    // Accent - Champagne Gold (エレガントなシャンパンゴールド)
    accent: '#b8956a',
    accentLight: '#d4b896',
    accentDark: '#a07d52',
    // Surface - Warm White (温かみのあるホワイト)
    surface: '#faf9f7',
    surfaceAlt: '#edeae5',
    surfaceDark: '#dedad3',
    // Charts (統一感のある配色)
    gold: '#c9a96e',      // ウォームゴールド
    brown: '#47566b',     // スレートブルー
    beige: '#dcc9b3',     // ウォームベージュ
    light: '#f6f4f1',     // ウォームホワイト
    white: '#ffffff',
    darkBrown: '#3d4859', // ダークスレート
    // Harmonious accent colors
    sage: '#739977',      // セージグリーン
    rose: '#b08f8a',      // ダスティローズ
    warmgold: '#d4b896',  // ライトゴールド
    success: '#739977',   // セージグリーン
    warning: '#c9a96e',   // ウォームゴールド
    purple: '#b08f8a'     // ダスティローズ
};

let rawData = [];
let customerData = []; // 顧客データ
let customerListCurrentPage = 1;
const customerListPageSize = 20;

// 顧客データAPI設定（未設定の場合は売上APIと同じURLを使用）
const CUSTOMER_API_KEY = 'mavie_customer_api_url';
let CUSTOMER_API_URL = localStorage.getItem(CUSTOMER_API_KEY) || "";
let charts = {};
let lockedStore = null;
let lockedStaff = null;
let isStaffAuthenticated = false; // スタッフ認証済みフラグ
let changedRows = new Set();
let monthlyGoal = 1100000; // Default goal

// --- STAFF PASSWORD FUNCTIONS ---
const STAFF_PASSWORD_STORAGE_KEY = 'mavie_staff_passwords';
let staffPasswordsCache = null; // スプレッドシートから読み込んだパスワードのキャッシュ

// ローカルストレージから読み込み（フォールバック用）
function loadStaffPasswordsFromLocalStorage() {
    const stored = localStorage.getItem(STAFF_PASSWORD_STORAGE_KEY);
    return stored ? JSON.parse(stored) : {};
}

// ローカルストレージに保存（フォールバック用）
function saveStaffPasswordsToLocalStorage(passwords) {
    localStorage.setItem(STAFF_PASSWORD_STORAGE_KEY, JSON.stringify(passwords));
}

// スプレッドシートからパスワードを読み込み
async function loadStaffPasswordsFromSpreadsheet() {
    const apiUrl = document.getElementById('spreadsheet-api-url')?.value || API_URL || DEFAULT_API_URL;
    if (!apiUrl) {
        // API未設定の場合はローカルストレージから読み込み
        staffPasswordsCache = loadStaffPasswordsFromLocalStorage();
        return staffPasswordsCache;
    }

    try {
        const response = await fetch(`${apiUrl}?action=load_passwords`);
        const result = await response.json();
        if (result.status === 'success' && result.passwords) {
            staffPasswordsCache = result.passwords;
            // ローカルストレージにもキャッシュ
            saveStaffPasswordsToLocalStorage(result.passwords);
            return result.passwords;
        }
    } catch (error) {
        console.error('パスワード読み込みエラー:', error);
    }

    // エラー時はローカルストレージから読み込み
    staffPasswordsCache = loadStaffPasswordsFromLocalStorage();
    return staffPasswordsCache;
}

// スプレッドシートにパスワードを保存
async function saveStaffPasswordsToSpreadsheet(passwords) {
    const apiUrl = document.getElementById('spreadsheet-api-url')?.value || API_URL || DEFAULT_API_URL;

    // ローカルストレージに保存（常に）
    saveStaffPasswordsToLocalStorage(passwords);
    staffPasswordsCache = passwords;

    if (!apiUrl) {
        return; // API未設定の場合はローカル保存のみ
    }

    try {
        // 階層構造をフラットなキー形式に変換 {chiba: {staff1: pass}} -> {chiba_staff1: pass}
        const flatPasswords = {};
        Object.keys(passwords).forEach(store => {
            Object.keys(passwords[store]).forEach(staff => {
                flatPasswords[store + '_' + staff] = passwords[store][staff];
            });
        });

        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'save_passwords',
                passwords: flatPasswords
            })
        });
        const result = await response.json();
        console.log('パスワード保存結果:', result);
    } catch (error) {
        console.error('パスワード保存エラー:', error);
    }
}

function getStaffPassword(store, staff) {
    // キャッシュがない場合はローカルストレージから読み込み
    const passwords = staffPasswordsCache || loadStaffPasswordsFromLocalStorage();
    if (passwords[store] && passwords[store][staff]) {
        return passwords[store][staff];
    }
    return null; // パスワード未設定
}

function setStaffPassword(store, staff, password) {
    const passwords = staffPasswordsCache || loadStaffPasswordsFromLocalStorage();
    if (!passwords[store]) passwords[store] = {};
    passwords[store][staff] = password;
    // スプレッドシートとローカルストレージ両方に保存
    saveStaffPasswordsToSpreadsheet(passwords);
}

// --- SETTINGS SYNC FUNCTIONS (スプレッドシート連携) ---
const SETTINGS_LOCAL_KEY = 'mavie_dashboard_settings';

// ローカルストレージから設定を読み込み（フォールバック用）
function loadSettingsFromLocalStorage() {
    const stored = localStorage.getItem(SETTINGS_LOCAL_KEY);
    return stored ? JSON.parse(stored) : { staffRoster: null, geminiApiKey: null };
}

// ローカルストレージに設定を保存
function saveSettingsToLocalStorage(settings) {
    localStorage.setItem(SETTINGS_LOCAL_KEY, JSON.stringify(settings));
}

// スプレッドシートから設定を読み込み
async function loadSettingsFromSpreadsheet() {
    const apiUrl = document.getElementById('spreadsheet-api-url')?.value || API_URL || DEFAULT_API_URL;
    if (!apiUrl) {
        // API未設定の場合はローカルストレージから読み込み
        const localSettings = loadSettingsFromLocalStorage();
        if (localSettings.staffRoster) {
            STAFF_ROSTER = localSettings.staffRoster;
        }
        return localSettings;
    }

    try {
        const response = await fetch(`${apiUrl}?action=load_settings`);
        const result = await response.json();
        if (result.status === 'success' && result.settings) {
            // スタッフ名簿を更新
            if (result.settings.staffRoster) {
                STAFF_ROSTER = result.settings.staffRoster;
            }
            // Gemini APIキーを更新
            if (result.settings.geminiApiKey) {
                const geminiInput = document.getElementById('gemini-api-key');
                if (geminiInput) geminiInput.value = result.settings.geminiApiKey;
            }
            // ローカルストレージにもキャッシュ
            saveSettingsToLocalStorage(result.settings);
            updateSyncStatus(true);
            return result.settings;
        }
    } catch (error) {
        console.error('設定読み込みエラー:', error);
    }

    // エラー時はローカルストレージから読み込み
    const localSettings = loadSettingsFromLocalStorage();
    if (localSettings.staffRoster) {
        STAFF_ROSTER = localSettings.staffRoster;
    }
    return localSettings;
}

// トースト通知を表示
function showSettingsToast(message, type = 'success') {
    // 既存のトーストを削除
    const existing = document.getElementById('settings-toast');
    if (existing) existing.remove();

    const toast = document.createElement('div');
    toast.id = 'settings-toast';
    const bgColor = type === 'success' ? 'bg-green-600' : type === 'error' ? 'bg-red-600' : 'bg-amber-600';
    const icon = type === 'success' ? '✓' : type === 'error' ? '✕' : '⏳';
    toast.className = `fixed bottom-6 right-6 ${bgColor} text-white px-6 py-3 rounded-lg shadow-lg z-[9999] flex items-center gap-2 text-sm font-bold transition-opacity duration-300`;
    toast.innerHTML = `<span>${icon}</span> ${message}`;
    document.body.appendChild(toast);

    if (type !== 'saving') {
        setTimeout(() => {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    }
}

// スプレッドシートに設定を保存
async function saveSettingsToSpreadsheet(showFeedback = false) {
    const apiUrl = document.getElementById('spreadsheet-api-url')?.value || API_URL || DEFAULT_API_URL;
    const geminiApiKey = document.getElementById('gemini-api-key')?.value || '';

    const settings = {
        staffRoster: STAFF_ROSTER,
        geminiApiKey: geminiApiKey
    };

    // ローカルストレージに保存（常に）
    saveSettingsToLocalStorage(settings);

    if (!apiUrl) {
        if (showFeedback) showSettingsToast('API未設定のためローカルのみに保存しました', 'warning');
        return false;
    }

    if (showFeedback) showSettingsToast('スプレッドシートに保存中...', 'saving');

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'save_settings',
                settings: settings
            })
        });
        const result = await response.json();
        if (result.status === 'success') {
            console.log('設定をスプレッドシートに保存しました');
            if (showFeedback) showSettingsToast('スプレッドシートに保存しました');
            updateSyncStatus(true);
            return true;
        } else {
            throw new Error(result.message || '保存に失敗しました');
        }
    } catch (error) {
        console.error('設定保存エラー:', error);
        if (showFeedback) showSettingsToast('スプレッドシートへの保存に失敗しました', 'error');
        updateSyncStatus(false);
        return false;
    }
}

// 同期ステータス表示を更新
function updateSyncStatus(synced) {
    const statusEl = document.getElementById('settings-sync-status');
    if (!statusEl) return;
    if (synced) {
        statusEl.innerHTML = '<span class="inline-flex items-center gap-1 text-xs text-green-600 font-bold"><span>✓</span> スプレッドシートと同期済み</span>';
    } else {
        statusEl.innerHTML = '<span class="inline-flex items-center gap-1 text-xs text-red-600 font-bold"><span>!</span> 同期エラー - 手動保存をお試しください</span>';
    }
}

// 手動保存ボタンのハンドラ
async function manualSaveSettingsToSpreadsheet() {
    const btn = document.getElementById('btn-manual-save-settings');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = '<span>⏳</span> 保存中...';
    }
    const success = await saveSettingsToSpreadsheet(true);
    if (btn) {
        btn.disabled = false;
        btn.innerHTML = '<i data-lucide="cloud-upload" class="w-4 h-4"></i> スプレッドシートに設定を保存';
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }
}

function showStaffLoginModal(store, staff) {
    const modal = document.getElementById('staff-login-modal');
    const nameEl = document.getElementById('staff-login-name');
    const passwordInput = document.getElementById('staff-login-password');
    const errorEl = document.getElementById('staff-login-error');
    const mainContent = document.getElementById('main-content');
    const header = document.querySelector('header');

    // スタッフ名を表示
    const storeName = store === 'chiba' ? '千葉店' : store === 'honatsugi' ? '本厚木店' : store === 'yamato' ? '大和店' : store;
    nameEl.textContent = `${storeName} / ${staff}`;

    // エラーメッセージを非表示
    errorEl.classList.add('hidden');
    passwordInput.value = '';

    // モーダルを表示
    modal.classList.remove('hidden');

    // メインコンテンツとヘッダーを非表示
    if (mainContent) mainContent.style.display = 'none';
    if (header) header.style.display = 'none';

    // パスワード入力欄にフォーカス
    setTimeout(() => passwordInput.focus(), 100);
}

function hideStaffLoginModal() {
    const modal = document.getElementById('staff-login-modal');
    const mainContent = document.getElementById('main-content');
    const header = document.querySelector('header');

    modal.classList.add('hidden');

    // メインコンテンツとヘッダーを表示
    if (mainContent) mainContent.style.display = '';
    if (header) header.style.display = '';
}

function handleStaffLogin(event) {
    event.preventDefault();
    const passwordInput = document.getElementById('staff-login-password');
    const errorEl = document.getElementById('staff-login-error');
    const enteredPassword = passwordInput.value;

    const correctPassword = getStaffPassword(lockedStore, lockedStaff);

    // パスワード未設定の場合は空文字列で認証成功
    if (correctPassword === null || correctPassword === '' || enteredPassword === correctPassword) {
        isStaffAuthenticated = true;
        hideStaffLoginModal();
        // ダッシュボードを更新
        updateDashboard();
        // Lucide iconsを更新
        if (typeof lucide !== 'undefined') {
            lucide.createIcons();
        }
    } else {
        // パスワード不一致
        errorEl.classList.remove('hidden');
        passwordInput.value = '';
        passwordInput.focus();
    }
}

// --- GOAL STORAGE FUNCTIONS ---
const GOAL_STORAGE_KEY = 'mavie_staff_goals_v2'; // v2: 月別対応
const GOAL_STORAGE_KEY_LEGACY = 'mavie_staff_goals'; // 旧バージョン
const BASE_SALARY_STORAGE_KEY = 'mavie_staff_base_salary';
const DEFAULT_GOAL = {
    weekdays: 0,
    weekends: 0,
    weekdayTarget: 0,
    weekendTarget: 0,
    retail: 0,
    newCustomers: 0,
    existingCustomers: 0,
    unitPrice: 0,
    newReservationRate: 0,
    reservationRate: 0,
    reviews5Star: 0
};
const DEFAULT_BASE_SALARY = 220000; // Default base salary (vienna: 201600)

// 現在選択されている月を取得（YYYY/M形式 - date-selectorと同じ形式）
function getCurrentGoalMonth() {
    const monthSelector = document.getElementById('goal-month-selector');
    if (monthSelector && monthSelector.value) {
        // YYYY/MM形式をYYYY/M形式に正規化
        const parts = monthSelector.value.split('/');
        if (parts.length === 2) {
            return `${parts[0]}/${parseInt(parts[1])}`;
        }
        return monthSelector.value;
    }
    // デフォルト：現在の月
    const now = new Date();
    return `${now.getFullYear()}/${now.getMonth() + 1}`;
}

// ダッシュボードで選択されている月を取得（YYYY/M形式）
function getSelectedDashboardMonth() {
    const dateSelector = document.getElementById('date-selector');
    if (dateSelector && dateSelector.value) {
        return dateSelector.value;
    }
    const now = new Date();
    return `${now.getFullYear()}/${now.getMonth() + 1}`;
}

// 月の形式を正規化（YYYY/MM -> YYYY/M）
function normalizeMonthFormat(monthStr) {
    if (!monthStr) return monthStr;
    const parts = monthStr.split('/');
    if (parts.length === 2) {
        return `${parts[0]}/${parseInt(parts[1])}`;
    }
    return monthStr;
}

function loadGoalsFromStorage() {
    const stored = localStorage.getItem(GOAL_STORAGE_KEY);
    if (stored) {
        const goals = JSON.parse(stored);
        // 月の形式を正規化（YYYY/MM -> YYYY/M）
        const normalizedGoals = {};
        let needsSave = false;
        for (const [month, storeData] of Object.entries(goals)) {
            const normalizedMonth = normalizeMonthFormat(month);
            if (normalizedMonth !== month) needsSave = true;
            if (!normalizedGoals[normalizedMonth]) {
                normalizedGoals[normalizedMonth] = storeData;
            } else {
                // 同じ月のデータがある場合はマージ
                for (const [store, staffData] of Object.entries(storeData)) {
                    if (!normalizedGoals[normalizedMonth][store]) {
                        normalizedGoals[normalizedMonth][store] = staffData;
                    } else {
                        Object.assign(normalizedGoals[normalizedMonth][store], staffData);
                    }
                }
            }
        }
        // 形式が変更された場合は保存し直す
        if (needsSave) {
            saveGoalsToStorage(normalizedGoals);
        }
        return normalizedGoals;
    }
    // 旧バージョンからのマイグレーション
    const legacy = localStorage.getItem(GOAL_STORAGE_KEY_LEGACY);
    if (legacy) {
        const legacyGoals = JSON.parse(legacy);
        // 現在の月として保存
        const currentMonth = getSelectedDashboardMonth();
        const migratedGoals = { [currentMonth]: legacyGoals };
        saveGoalsToStorage(migratedGoals);
        return migratedGoals;
    }
    return {};
}

function saveGoalsToStorage(goals) {
    localStorage.setItem(GOAL_STORAGE_KEY, JSON.stringify(goals));
}

// デバッグ用: localStorageの目標データを確認
window.debugGoals = function() {
    const goals = loadGoalsFromStorage();
    console.log('=== Goal Storage Debug ===');
    console.log('All stored goals:', goals);
    console.log('Available months:', Object.keys(goals));
    Object.keys(goals).forEach(month => {
        console.log(`Month ${month}:`, goals[month]);
        Object.keys(goals[month] || {}).forEach(store => {
            console.log(`  Store ${store}:`, goals[month][store]);
        });
    });
    console.log('Current goal month (from goal-month-selector):', getCurrentGoalMonth());
    console.log('Dashboard month (from date-selector):', getSelectedDashboardMonth());
    console.log('STAFF_ROSTER:', STAFF_ROSTER);

    // 店舗ごとの月間目標合計を表示
    console.log('=== Store Monthly Target Sums ===');
    const currentMonth = getCurrentGoalMonth();
    Object.keys(STAFF_ROSTER).forEach(store => {
        const storeGoal = getStoreAggregateGoal(store, currentMonth);
        console.log(`${store}: ¥${storeGoal.monthlyTargetSum.toLocaleString()} (個別スタッフ目標の合計)`);
        const incorrectCalc = (storeGoal.weekdays * storeGoal.weekdayTarget) + (storeGoal.weekends * storeGoal.weekendTarget);
        console.log(`  従来の計算: ¥${incorrectCalc.toLocaleString()} (平均出勤日数 × 合計目標)`);
    });

    return goals;
};

function getStaffGoal(store, staff, yearMonth = null) {
    const goals = loadGoalsFromStorage();
    const month = yearMonth || getSelectedDashboardMonth();
    if (goals[month] && goals[month][store] && goals[month][store][staff]) {
        return goals[month][store][staff];
    }
    return { ...DEFAULT_GOAL };
}

function saveStaffGoal(store, staff, goalData, yearMonth = null) {
    const goals = loadGoalsFromStorage();
    const month = yearMonth || getCurrentGoalMonth();
    console.log(`[saveStaffGoal] Saving goal for ${store}/${staff} under month ${month}`);
    if (!goals[month]) goals[month] = {};
    if (!goals[month][store]) goals[month][store] = {};
    goals[month][store][staff] = goalData;
    saveGoalsToStorage(goals);

    // スプレッドシートへの自動保存をスケジュール
    scheduleAutoSaveToSpreadsheet();
}

// スタッフ名から所属店舗を取得（STAFF_ROSTERから検索）- 大文字小文字を区別しない
function getStaffStoreFromRoster(staffName) {
    const staffNameLower = staffName?.toLowerCase();
    for (const [store, staffList] of Object.entries(STAFF_ROSTER)) {
        if (staffList.some(s => s.toLowerCase() === staffNameLower)) return store;
    }
    return null;
}

function getStoreAggregateGoal(store, yearMonth = null) {
    const staffList = STAFF_ROSTER[store] || [];
    const month = yearMonth || getSelectedDashboardMonth();


    let aggregate = {
        weekdays: 0,
        weekends: 0,
        weekdayTarget: 0,
        weekendTarget: 0,
        retail: 0,
        newCustomers: 0,
        existingCustomers: 0,
        unitPrice: 0,
        newReservationRate: 0,
        reservationRate: 0,
        reviews5Star: 0,
        // 個別スタッフの月間目標の正確な合計
        monthlyTargetSum: 0
    };

    let staffCount = 0;
    let totalWeekdays = 0;
    let totalWeekends = 0;

    staffList.forEach(staff => {
        const staffGoal = getStaffGoal(store, staff, month);

        // 各スタッフの月間目標を正確に計算して合計
        const staffMonthlyTarget = ((staffGoal.weekdays || 0) * (staffGoal.weekdayTarget || 0)) +
                                   ((staffGoal.weekends || 0) * (staffGoal.weekendTarget || 0));
        aggregate.monthlyTargetSum += staffMonthlyTarget;

        // 出勤日数は平均を取るために合計
        totalWeekdays += staffGoal.weekdays || 0;
        totalWeekends += staffGoal.weekends || 0;

        // デイリー目標は店舗全体の売上として合計
        aggregate.weekdayTarget += staffGoal.weekdayTarget || 0;
        aggregate.weekendTarget += staffGoal.weekendTarget || 0;

        // その他の目標も合計
        aggregate.retail += staffGoal.retail || 0;
        aggregate.newCustomers += staffGoal.newCustomers || 0;
        aggregate.existingCustomers += staffGoal.existingCustomers || 0;
        aggregate.unitPrice += staffGoal.unitPrice || 0;
        aggregate.newReservationRate += staffGoal.newReservationRate || 0;
        aggregate.reservationRate += staffGoal.reservationRate || 0;
        aggregate.reviews5Star += staffGoal.reviews5Star || 0;
        staffCount++;
    });

    // 出勤日数は平均値（店舗の営業日数の目安）
    if (staffCount > 0) {
        aggregate.weekdays = Math.round(totalWeekdays / staffCount);
        aggregate.weekends = Math.round(totalWeekends / staffCount);
        aggregate.unitPrice = Math.round(aggregate.unitPrice / staffCount);
        aggregate.newReservationRate = Math.round(aggregate.newReservationRate / staffCount);
        aggregate.reservationRate = Math.round(aggregate.reservationRate / staffCount);
    }

    return aggregate;
}

function getAllStoresAggregateGoal(yearMonth = null) {
    const allStores = Object.keys(STAFF_ROSTER);
    const month = yearMonth || getSelectedDashboardMonth();


    let aggregate = {
        weekdays: 0,
        weekends: 0,
        weekdayTarget: 0,
        weekendTarget: 0,
        retail: 0,
        newCustomers: 0,
        existingCustomers: 0,
        unitPrice: 0,
        newReservationRate: 0,
        reservationRate: 0,
        reviews5Star: 0,
        // 個別スタッフの月間目標の正確な合計
        monthlyTargetSum: 0
    };

    let staffCount = 0;
    let totalWeekdays = 0;
    let totalWeekends = 0;

    allStores.forEach(store => {
        const staffList = STAFF_ROSTER[store] || [];
        staffList.forEach(staff => {
            const staffGoal = getStaffGoal(store, staff, month);

            // 各スタッフの月間目標を正確に計算して合計
            const staffMonthlyTarget = ((staffGoal.weekdays || 0) * (staffGoal.weekdayTarget || 0)) +
                                       ((staffGoal.weekends || 0) * (staffGoal.weekendTarget || 0));
            aggregate.monthlyTargetSum += staffMonthlyTarget;

            // 出勤日数は平均を取るために合計
            totalWeekdays += staffGoal.weekdays || 0;
            totalWeekends += staffGoal.weekends || 0;

            // デイリー目標は全店舗の売上として合計
            aggregate.weekdayTarget += staffGoal.weekdayTarget || 0;
            aggregate.weekendTarget += staffGoal.weekendTarget || 0;

            // その他の目標も合計
            aggregate.retail += staffGoal.retail || 0;
            aggregate.newCustomers += staffGoal.newCustomers || 0;
            aggregate.existingCustomers += staffGoal.existingCustomers || 0;
            aggregate.unitPrice += staffGoal.unitPrice || 0;
            aggregate.newReservationRate += staffGoal.newReservationRate || 0;
            aggregate.reservationRate += staffGoal.reservationRate || 0;
            aggregate.reviews5Star += staffGoal.reviews5Star || 0;
            staffCount++;
        });
    });

    // 出勤日数は平均値（全店舗の営業日数の目安）
    if (staffCount > 0) {
        aggregate.weekdays = Math.round(totalWeekdays / staffCount);
        aggregate.weekends = Math.round(totalWeekends / staffCount);
        aggregate.unitPrice = Math.round(aggregate.unitPrice / staffCount);
        aggregate.newReservationRate = Math.round(aggregate.newReservationRate / staffCount);
        aggregate.reservationRate = Math.round(aggregate.reservationRate / staffCount);
    }

    return aggregate;
}

function getCurrentGoalContext() {
    const storeVal = document.getElementById('store-selector').value;
    const staffVal = document.getElementById('staff-selector').value;

    if (staffVal !== 'all') {
        return { type: 'staff', store: storeVal, staff: staffVal };
    } else if (storeVal !== 'all') {
        return { type: 'store', store: storeVal };
    } else {
        return { type: 'all' };
    }
}

// --- BASE SALARY FUNCTIONS ---
function loadBaseSalariesFromStorage() {
    const stored = localStorage.getItem(BASE_SALARY_STORAGE_KEY);
    return stored ? JSON.parse(stored) : {};
}

function saveBaseSalariesToStorage(salaries) {
    localStorage.setItem(BASE_SALARY_STORAGE_KEY, JSON.stringify(salaries));
}

function getStaffBaseSalary(store, staff) {
    const salaries = loadBaseSalariesFromStorage();
    if (salaries[store] && salaries[store][staff] !== undefined) {
        return salaries[store][staff];
    }
    // vienna のみ基本給 ¥201,600、その他は ¥220,000
    if (staff === 'vienna') return 201600;
    return DEFAULT_BASE_SALARY;
}

function saveStaffBaseSalary(store, staff, salary) {
    const salaries = loadBaseSalariesFromStorage();
    if (!salaries[store]) salaries[store] = {};
    salaries[store][staff] = salary;
    saveBaseSalariesToStorage(salaries);

    // スプレッドシートへの自動保存をスケジュール
    scheduleAutoSaveToSpreadsheet();
}

// --- AUTO SAVE TO SPREADSHEET ---
let autoSaveTimeout = null;
let isSavingToSpreadsheet = false;

/**
 * スプレッドシートへの自動保存をスケジュール（5秒後に実行）
 */
function scheduleAutoSaveToSpreadsheet() {
    // 既存のタイマーをクリア
    if (autoSaveTimeout) {
        clearTimeout(autoSaveTimeout);
    }

    // 5秒後に自動保存
    autoSaveTimeout = setTimeout(() => {
        autoSaveToSpreadsheet();
    }, 5000);
}

/**
 * スプレッドシートに自動保存（静かに実行）
 */
async function autoSaveToSpreadsheet() {
    if (!API_URL || isSavingToSpreadsheet) {
        return;
    }

    isSavingToSpreadsheet = true;

    try {
        const goals = loadGoalsFromStorage();
        const salaries = loadBaseSalariesFromStorage();

        const response = await fetch(API_URL, {
            method: 'POST',
            headers: { "Content-Type": "text/plain" },
            body: JSON.stringify({
                action: "save_goals",
                goals: goals,
                salaries: salaries
            })
        });

        const result = await response.json();

        if (result.status === "success") {
            console.log("目標データをスプレッドシートに自動保存しました");
        } else {
            console.warn("自動保存エラー:", result.message);
        }

    } catch (e) {
        console.error("自動保存中にエラーが発生:", e);
    } finally {
        isSavingToSpreadsheet = false;
    }
}

/**
 * ページ読み込み時にスプレッドシートから目標データを自動読み込み
 */
async function autoLoadFromSpreadsheet() {
    if (!API_URL) {
        console.log("API URLが設定されていないため、ローカルストレージのデータを使用します");
        return;
    }

    try {
        const response = await fetch(API_URL + '?action=load_goals');
        const result = await response.json();

        if (result.status === "success") {
            // 目標データをローカルストレージに保存
            if (result.goals) {
                saveGoalsToStorage(result.goals);
            }
            // 基本給データを保存
            if (result.salaries) {
                saveBaseSalariesToStorage(result.salaries);
            }
            console.log("スプレッドシートから目標データを読み込みました");
        } else {
            console.log("スプレッドシートからの読み込みに失敗:", result.message);
        }

    } catch (e) {
        console.error("スプレッドシートからの自動読み込み中にエラーが発生:", e);
    }
}

// --- INCENTIVE CALCULATION ---
function calculateIncentive(data, store, staff) {
    // Get base salary
    const baseSalary = getStaffBaseSalary(store, staff);

    // Calculate service sales (cash + credit + qr) and retail sales
    let serviceSalesRaw = 0;
    let retailSalesRaw = 0;

    // dataが配列でない場合は空配列として扱う
    const dataArray = Array.isArray(data) ? data : [];

    dataArray.forEach(d => {
        if (!d) return;
        const sales = d.sales || {};
        serviceSalesRaw += (sales.cash || 0) + (sales.credit || 0) + (sales.qr || 0);
        retailSalesRaw += sales.product || 0;
    });

    // 税抜計算（消費税5%として計算 - 会社負担5割の考え方）
    const serviceSalesTaxExcl = serviceSalesRaw / 1.05;
    const retailSalesTaxExcl = retailSalesRaw / 1.05;

    // 施術インセンティブ: max(0, 施術売上(税抜) × 40% - 基本給)
    const serviceIncentive = Math.max(0, serviceSalesTaxExcl * 0.4 - baseSalary);

    // 物販インセンティブ: 物販売上(税抜) × 10%
    const retailIncentive = retailSalesTaxExcl * 0.1;

    // 推定報酬合計 = 基本給 + 施術インセンティブ + 物販インセンティブ
    const totalIncentive = baseSalary + serviceIncentive + retailIncentive;

    return {
        baseSalary,
        serviceSalesRaw,
        serviceSalesTaxExcl,
        serviceIncentive,
        retailSalesRaw,
        retailSalesTaxExcl,
        retailIncentive,
        totalIncentive
    };
}

// --- 0. INIT & DATE SELECTOR ---
function initDateSelector() {
    const sel = document.getElementById('date-selector');
    const startYear = 2024;
    const endYear = 2030;
    const today = new Date();
    const currentYear = today.getFullYear();
    const currentMonth = today.getMonth() + 1;

    for (let y = startYear; y <= endYear; y++) {
        for (let m = 1; m <= 12; m++) {
            const opt = document.createElement('option');
            const val = `${y}/${m}`; 
            opt.value = val;
            opt.text = `${y}年 ${m}月`;
            if (y === currentYear && m === currentMonth) opt.selected = true;
            sel.appendChild(opt);
        }
    }
}

function handleDateChange() {
    const dateVal = document.getElementById('date-selector').value;
    if (!API_URL) rawData = generateData(dateVal);
    updateDashboard();
}

// --- 1. MOCK DATA ---
function generateData(yearMonthStr = "2024/1") {
    const data = [];
    const [year, month] = yearMonthStr.split('/').map(Number);
    const daysInMonth = new Date(year, month, 0).getDate();
    let idCounter = 1;

    Object.keys(STAFF_ROSTER).forEach(storeKey => {
        STAFF_ROSTER[storeKey].forEach(staffName => {
            for (let i = 1; i <= daysInMonth; i++) {
                if (Math.random() > 0.8) continue;

                // All values set to 0
                const salesCash = 0;
                const salesCredit = 0;
                const salesQr = 0;
                const salesProduct = 0;

                // All customer values set to 0
                let newHPB = 0;
                let newMiniNai = 0;
                const acquaintance = 0;
                const existing = 0;

                const nextResNewHPB = 0;
                const nextResMiniNai = 0;
                const nextResExisting = 0;

                data.push({
                    id: idCounter++,
                    date: `${year}/${month}/${i}`,
                    store: storeKey,
                    storeName: storeKey === 'chiba' ? '千葉店' : storeKey === 'honatsugi' ? '本厚木店' : storeKey === 'yamato' ? '大和店' : storeKey,
                    staff: staffName,
                    sales: { cash: salesCash, credit: salesCredit, qr: salesQr, product: salesProduct },
                    discounts: { hpbPoints: Math.random()>0.9?500:0, hpbGift: 0, other: 0, refund: Math.random()>0.98?3000:0 },
                    customers: { newHPB, newMiniNai, existing, acquaintance },
                    nextRes: { newHPB: nextResNewHPB, newMiniNai: nextResMiniNai, existing: nextResExisting },
                    reviews5Star: 0,
                    blogUpdates: 0,
                    snsUpdates: 0
                });
            }
        });
    });
    return data;
}

// --- 2. DATA LOADING ---

// スプラッシュスクリーン / ローディングオーバーレイの表示/非表示
function showLoading(text = '読み込み中...') {
    const overlay = document.getElementById('loading-overlay');
    const loadingText = document.getElementById('loading-text');
    const progressBar = document.getElementById('splash-progress-bar');
    if (overlay) {
        overlay.classList.remove('hidden', 'splash-fade-out');
        overlay.style.opacity = '1';
        overlay.style.visibility = 'visible';
    }
    if (loadingText) loadingText.textContent = text;
    // プログレスバーのアニメーションをリセット
    if (progressBar) {
        progressBar.style.animation = 'none';
        progressBar.offsetHeight; // リフロー強制
        progressBar.style.animation = 'progressBar 2s ease-in-out forwards';
    }
}

function hideLoading() {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.classList.add('splash-fade-out');
        setTimeout(() => {
            overlay.classList.add('hidden');
            overlay.style.visibility = 'hidden';
        }, 500);
    }
}

// 更新完了通知を表示（簡易トースト）
function showUpdateNotification(message) {
    // 既存の通知を削除
    const existing = document.getElementById('update-notification');
    if (existing) existing.remove();

    // 新しい通知を作成
    const notification = document.createElement('div');
    notification.id = 'update-notification';
    notification.className = 'fixed bottom-4 right-4 bg-emerald-500 text-white px-4 py-3 rounded-lg shadow-lg z-50 flex items-center gap-2 animate-fade-in';
    notification.innerHTML = `
        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
        </svg>
        <span class="font-medium">${message}</span>
    `;
    document.body.appendChild(notification);

    // 2秒後に自動で消える
    setTimeout(() => {
        notification.style.opacity = '0';
        notification.style.transition = 'opacity 0.3s';
        setTimeout(() => notification.remove(), 300);
    }, 2000);
}

// スプレッドシートAPI URL管理
function saveSpreadsheetApiUrl() {
    const urlInput = document.getElementById('spreadsheet-api-url');
    const url = urlInput ? urlInput.value.trim() : '';
    localStorage.setItem(SPREADSHEET_API_KEY, url);
    API_URL = url;

    const status = document.getElementById('spreadsheet-connection-status');
    if (status) {
        status.classList.remove('hidden');
        status.innerHTML = '<div class="bg-green-50 border border-green-400 text-green-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="check-circle" class="w-4 h-4"></i>APIのURLが保存されました</div>';
        if (typeof lucide !== 'undefined') lucide.createIcons();
        setTimeout(() => status.classList.add('hidden'), 3000);
    }
}

function loadSpreadsheetApiUrl() {
    const urlInput = document.getElementById('spreadsheet-api-url');
    // デフォルトURLをフォールバックとして使用
    const savedUrl = localStorage.getItem(SPREADSHEET_API_KEY) || DEFAULT_API_URL;
    if (urlInput) {
        urlInput.value = savedUrl;
        API_URL = savedUrl;
    }
}

async function testSpreadsheetConnection() {
    const status = document.getElementById('spreadsheet-connection-status');
    const url = document.getElementById('spreadsheet-api-url')?.value.trim() || API_URL;

    if (!url) {
        if (status) {
            status.classList.remove('hidden');
            status.innerHTML = '<div class="bg-yellow-50 border border-yellow-400 text-yellow-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="alert-circle" class="w-4 h-4"></i>URLを入力してください</div>';
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
        return;
    }

    if (status) {
        status.classList.remove('hidden');
        status.innerHTML = '<div class="bg-blue-50 border border-blue-400 text-blue-700 px-4 py-2 rounded text-sm flex items-center gap-2"><svg class="animate-spin w-4 h-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg>接続テスト中...</div>';
    }

    try {
        const response = await fetch(url);
        const data = await response.json();

        if (status) {
            status.innerHTML = `<div class="bg-green-50 border border-green-400 text-green-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="check-circle" class="w-4 h-4"></i>接続成功！ ${Array.isArray(data) ? data.length + '件のデータを取得' : 'データを取得しました'}</div>`;
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }

        // 接続成功時に設定とパスワードもスプレッドシートから読み込み
        await loadSettingsFromSpreadsheet();
        await loadStaffPasswordsFromSpreadsheet();
        updateSettingsList();
        updatePasswordList();
    } catch (e) {
        if (status) {
            status.innerHTML = `<div class="bg-red-50 border border-red-400 text-red-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="x-circle" class="w-4 h-4"></i>接続エラー: ${e.message}</div>`;
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
    }
}

// データ更新（手動）
async function refreshData() {
    const refreshBtn = document.getElementById('refresh-data-btn');
    const refreshIcon = document.getElementById('refresh-icon');

    // ボタンをスピン状態に
    if (refreshIcon) refreshIcon.classList.add('animate-spin');
    if (refreshBtn) refreshBtn.disabled = true;

    showLoading('データを更新中...', '最新データをスプレッドシートから取得しています');

    try {
        await loadDataFromSpreadsheet();
        updateDashboard();
    } finally {
        hideLoading();
        if (refreshIcon) refreshIcon.classList.remove('animate-spin');
        if (refreshBtn) refreshBtn.disabled = false;
    }
}

// スプレッドシートからデータ読み込み
async function loadDataFromSpreadsheet() {
    // APIのURLを最新の状態に更新（デフォルトURLをフォールバック）
    API_URL = localStorage.getItem(SPREADSHEET_API_KEY) || DEFAULT_API_URL;

    if (API_URL) {
        try {
            const response = await fetch(API_URL);
            if (!response.ok) throw new Error(`HTTP ${response.status}`);
            let data = await response.json();

            // レスポンスが配列でない場合（エラーオブジェクトなど）のチェック
            if (!Array.isArray(data)) {
                // { status: 'success', data: [...] } 形式の場合
                if (data && Array.isArray(data.data)) {
                    data = data.data;
                } else if (data && data.status === 'error') {
                    throw new Error(data.message || 'APIエラー');
                } else {
                    console.warn('予期しないレスポンス形式:', data);
                    return false;
                }
            }

            // idが欠けているレコードにのみidを付与し、データ構造を正規化
            rawData = data.map((d, i) => {
                const normalized = d.id ? {...d} : {...d, id: i+1};
                // nextResオブジェクトが存在しない場合は作成
                if (!normalized.nextRes) {
                    normalized.nextRes = {
                        newHPB: normalized.nextResNewHPB || normalized['nextRes.newHPB'] || 0,
                        newMiniNai: normalized.nextResNewMiniNai || normalized['nextRes.newMiniNai'] || 0,
                        existing: normalized.nextResExisting || normalized['nextRes.existing'] || 0
                    };
                }
                // customersオブジェクトが存在しない場合は作成
                if (!normalized.customers) {
                    normalized.customers = {
                        newHPB: normalized.customersNewHPB || normalized['customers.newHPB'] || 0,
                        newMiniNai: normalized.customersNewMiniNai || normalized['customers.newMiniNai'] || 0,
                        existing: normalized.customersExisting || normalized['customers.existing'] || 0,
                        acquaintance: normalized.customersAcquaintance || normalized['customers.acquaintance'] || 0
                    };
                }
                // salesオブジェクトが存在しない場合は作成
                if (!normalized.sales) {
                    normalized.sales = {
                        cash: normalized.salesCash || normalized['sales.cash'] || 0,
                        credit: normalized.salesCredit || normalized['sales.credit'] || 0,
                        qr: normalized.salesQr || normalized['sales.qr'] || 0,
                        product: normalized.salesProduct || normalized['sales.product'] || 0
                    };
                }
                // discountsオブジェクトが存在しない場合は作成
                if (!normalized.discounts) {
                    normalized.discounts = {
                        hpbPoints: normalized.discountsHpbPoints || normalized['discounts.hpbPoints'] || 0,
                        hpbGift: normalized.discountsHpbGift || normalized['discounts.hpbGift'] || 0,
                        other: normalized.discountsOther || normalized['discounts.other'] || 0,
                        refund: normalized.discountsRefund || normalized['discounts.refund'] || 0
                    };
                }
                return normalized;
            });
            console.log(`✓ スプレッドシートから ${rawData.length} 件のデータを取得しました`);
            return true;
        } catch (e) {
            console.error("スプレッドシート読み込みエラー:", e);
            console.error("API URL:", API_URL);
            // エラー通知（ページ上部に表示）
            setTimeout(() => {
                const banner = document.createElement('div');
                banner.className = 'fixed top-0 left-0 right-0 bg-red-500 text-white text-center py-2 text-sm z-50';
                banner.innerHTML = `スプレッドシート読み込みエラー: ${e.message} <button onclick="this.parentElement.remove()" class="ml-4 underline">閉じる</button>`;
                document.body.prepend(banner);
            }, 500);
            return false;
        }
    }
    return false;
}

async function loadData() {
    showLoading();
    try {
        initDateSelector();
        const initialDate = document.getElementById('date-selector').value;

        // スプレッドシートからデータを読み込み
        const loaded = await loadDataFromSpreadsheet();

        // スプレッドシートからの読み込みに失敗した場合はサンプルデータを使用
        if (!loaded) {
            console.log('サンプルデータを使用します');
            rawData = generateData(initialDate);
        }

        // ローカルストレージからスタッフ名簿を事前読み込み（URLパラメータマッチング前に必要）
        try {
            const savedSettings = loadSettingsFromLocalStorage();
            if (savedSettings && savedSettings.staffRoster) {
                STAFF_ROSTER = savedSettings.staffRoster;
            }
        } catch (e) {
            console.warn('スタッフ名簿事前読み込みスキップ:', e);
        }

        checkUrlParams();
        try { initCalendarSelectors(); } catch (e) { console.error('initCalendarSelectors エラー:', e); }
        updateDashboard();
    } catch (e) {
        console.error('loadData エラー:', e);
    } finally {
        hideLoading();
    }
}

// --- 顧客データ（マーケティング）関連 ---
function saveCustomerApiUrl() {
    const urlInput = document.getElementById('customer-api-url');
    const url = urlInput ? urlInput.value.trim() : '';
    localStorage.setItem(CUSTOMER_API_KEY, url);
    CUSTOMER_API_URL = url;

    const status = document.getElementById('customer-api-connection-status');
    if (status) {
        status.classList.remove('hidden');
        status.innerHTML = '<div class="bg-green-50 border border-green-400 text-green-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="check-circle" class="w-4 h-4"></i>URLが保存されました</div>';
        if (typeof lucide !== 'undefined') lucide.createIcons();
        setTimeout(() => status.classList.add('hidden'), 3000);
    }
}

function loadCustomerApiUrl() {
    const urlInput = document.getElementById('customer-api-url');
    const savedUrl = localStorage.getItem(CUSTOMER_API_KEY) || '';
    if (urlInput) urlInput.value = savedUrl;
}

async function testCustomerApiConnection() {
    const status = document.getElementById('customer-api-connection-status');
    const url = document.getElementById('customer-api-url')?.value.trim() || CUSTOMER_API_URL;

    if (!url) {
        if (status) {
            status.classList.remove('hidden');
            status.innerHTML = '<div class="bg-yellow-50 border border-yellow-400 text-yellow-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="alert-circle" class="w-4 h-4"></i>URLを入力してください</div>';
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
        return;
    }

    if (status) {
        status.classList.remove('hidden');
        status.innerHTML = '<div class="bg-blue-50 border border-blue-400 text-blue-700 px-4 py-2 rounded text-sm flex items-center gap-2"><svg class="animate-spin w-4 h-4" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path></svg>接続テスト中...</div>';
    }

    try {
        // action=get_customersパラメータを追加して顧客データを取得
        const testUrl = new URL(url);
        testUrl.searchParams.set('action', 'get_customers');

        // Google Apps Script Web App はリダイレクトするため、redirect: 'follow' が必要
        const response = await fetch(testUrl.toString(), {
            method: 'GET',
            redirect: 'follow'
        });

        // レスポンスのContent-Typeを確認
        const contentType = response.headers.get('content-type');
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }

        const text = await response.text();
        let result;
        try {
            result = JSON.parse(text);
        } catch (parseError) {
            // HTMLが返ってきた場合（認証エラーなど）
            if (text.includes('<!DOCTYPE') || text.includes('<html')) {
                throw new Error('認証エラー: Apps Scriptのデプロイ設定を確認してください（アクセスできるユーザー: 全員）');
            }
            throw new Error('JSONパースエラー: レスポンスがJSON形式ではありません');
        }

        // エラーレスポンスのチェック
        if (result.error || result.status === 'error') {
            throw new Error(`スクリプトエラー: ${result.error || result.message}`);
        }

        // APIレスポンス形式に対応 { status: 'success', data: [...] }
        const data = result.data || result;

        if (status) {
            status.innerHTML = `<div class="bg-green-50 border border-green-400 text-green-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="check-circle" class="w-4 h-4"></i>接続成功！ ${Array.isArray(data) ? data.length + '件の顧客データを取得' : 'データを取得しました'}</div>`;
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
    } catch (e) {
        if (status) {
            status.innerHTML = `<div class="bg-red-50 border border-red-400 text-red-700 px-4 py-2 rounded text-sm flex items-center gap-2"><i data-lucide="x-circle" class="w-4 h-4"></i>接続エラー: ${e.message}</div>`;
            if (typeof lucide !== 'undefined') lucide.createIcons();
        }
    }
}

async function loadCustomerData() {
    CUSTOMER_API_URL = localStorage.getItem(CUSTOMER_API_KEY) || '';
    // 顧客API URLが未設定の場合、売上APIのURLをフォールバック（同じGAS Web App）
    if (!CUSTOMER_API_URL) {
        CUSTOMER_API_URL = API_URL || DEFAULT_API_URL;
    }
    if (!CUSTOMER_API_URL) {
        console.log('顧客データAPIが設定されていません');
        return false;
    }

    try {
        // action=get_customersパラメータを追加して顧客データを取得
        const url = new URL(CUSTOMER_API_URL);
        url.searchParams.set('action', 'get_customers');

        const response = await fetch(url.toString(), {
            method: 'GET',
            redirect: 'follow'
        });
        if (!response.ok) throw new Error(`HTTP ${response.status}`);

        const text = await response.text();
        const result = JSON.parse(text);

        // APIレスポンス形式に対応 { status: 'success', data: [...] }
        if (result.error || result.status === 'error') {
            throw new Error(result.error || result.message || '不明なエラー');
        }

        // dataプロパティがあればそれを使用、なければ結果をそのまま使用
        customerData = result.data || result;

        if (!Array.isArray(customerData)) {
            throw new Error('顧客データの形式が正しくありません');
        }

        console.log(`✓ 顧客データから ${customerData.length} 件のデータを取得しました`);
        return true;
    } catch (e) {
        console.error('顧客データ読み込みエラー:', e);
        return false;
    }
}

async function refreshCustomerData() {
    const refreshBtn = document.getElementById('refresh-customer-btn');
    const refreshIcon = document.getElementById('refresh-customer-icon');

    if (refreshIcon) refreshIcon.classList.add('animate-spin');
    if (refreshBtn) refreshBtn.disabled = true;

    const loaded = await loadCustomerData();
    if (loaded) {
        updateMarketingDashboard();
        document.getElementById('customer-data-status').innerHTML = `<span class="text-emerald-600">${customerData.length}件のデータ</span>`;
        showUpdateNotification('マーケティングデータを更新しました');
    } else {
        document.getElementById('customer-data-status').innerHTML = '<span class="text-red-500">取得失敗</span>';
    }

    if (refreshIcon) refreshIcon.classList.remove('animate-spin');
    if (refreshBtn) refreshBtn.disabled = false;
}

// 店舗フィルターでフィルタリングされた顧客データを取得
function getFilteredCustomerData() {
    const storeFilter = document.getElementById('marketing-store-filter')?.value || 'all';
    if (storeFilter === 'all') {
        return customerData;
    }
    return customerData.filter(c => c.store === storeFilter);
}

// マーケティングダッシュボード更新
function updateMarketingDashboard() {
    if (!customerData || customerData.length === 0) {
        return;
    }

    const filtered = getFilteredCustomerData();
    const storeFilter = document.getElementById('marketing-store-filter')?.value || 'all';
    const storeName = storeFilter === 'all' ? '全店' : storeFilter === 'chiba' ? '千葉店' : storeFilter === 'honatsugi' ? '本厚木店' : storeFilter === 'yamato' ? '大和店' : storeFilter;

    const fmt = n => n.toLocaleString();
    const now = new Date();
    const thisMonth = now.getMonth();
    const thisYear = now.getFullYear();

    // KPI計算（フィルター済みデータを使用）
    const totalCustomers = filtered.length;
    const newThisMonth = filtered.filter(c => {
        if (!c.timestamp) return false;
        const d = parseDate(c.timestamp);
        return d.getMonth() === thisMonth && d.getFullYear() === thisYear;
    }).length;

    // snsOk(Code.gs) または snsPermission の両方に対応
    const snsYes = filtered.filter(c => {
        const snsVal = c.snsOk || c.snsPermission || '';
        return snsVal.includes('はい') || snsVal.includes('OK') || snsVal.includes('許可');
    }).length;
    const snsRate = totalCustomers > 0 ? Math.round((snsYes / totalCustomers) * 100) : 0;

    // 年齢計算 (birthday(Code.gs) または birthDate の両方に対応)
    const ages = filtered.map(c => {
        const birthVal = c.birthday || c.birthDate;
        if (!birthVal) return null;
        const birth = parseDate(birthVal);
        if (isNaN(birth.getTime())) return null;
        return now.getFullYear() - birth.getFullYear();
    }).filter(a => a && a > 0 && a < 100);
    const avgAge = ages.length > 0 ? Math.round(ages.reduce((a, b) => a + b, 0) / ages.length) : '-';

    // KPI更新
    document.getElementById('mkt-total-customers').textContent = fmt(totalCustomers);
    document.getElementById('mkt-new-this-month').textContent = fmt(newThisMonth);
    document.getElementById('mkt-sns-rate').textContent = `${snsRate}%`;
    document.getElementById('mkt-avg-age').textContent = avgAge + (avgAge !== '-' ? '歳' : '');

    // データステータス更新
    document.getElementById('customer-data-status').innerHTML = `<span class="text-emerald-600">${storeName}: ${totalCustomers}件</span>`;

    // チャート更新
    updateVisitReasonChart();
    updateAgeDistributionChart();
    updateOccupationChart();
    updateGeographicAnalysis();
    updateMenuPreferences();
    updateCustomerList();

    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function updateVisitReasonChart() {
    const filtered = getFilteredCustomerData();
    const reasonCounts = {};
    filtered.forEach(c => {
        if (c.visitReason) {
            const reasons = c.visitReason.split(',').map(r => r.trim());
            reasons.forEach(r => {
                if (r) reasonCounts[r] = (reasonCounts[r] || 0) + 1;
            });
        }
    });

    const sorted = Object.entries(reasonCounts).sort((a, b) => b[1] - a[1]);
    const labels = sorted.map(([k]) => k.length > 15 ? k.substring(0, 15) + '...' : k);
    const data = sorted.map(([, v]) => v);
    // Elegant harmonious color palette
    const colors = [
        '#b8956a', // Champagne Gold
        '#47566b', // Slate Blue
        '#739977', // Sage Green
        '#b08f8a', // Dusty Rose
        '#c9a96e', // Warm Gold
        '#6e819c', // Light Slate
        '#8ba88e', // Light Sage
        '#c4a5a0', // Light Rose
        '#d4b896', // Light Gold
        '#566882', // Medium Slate
        '#5d7d60', // Deep Sage
        '#9a7873'  // Deep Rose
    ];

    if (charts.visitReason) charts.visitReason.destroy();
    charts.visitReason = new Chart(document.getElementById('visitReasonChart'), {
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: colors.slice(0, data.length),
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right', labels: { boxWidth: 12, font: { size: 11 } } }
            }
        }
    });

    // 詳細リスト
    document.getElementById('visit-reason-details').innerHTML = sorted.slice(0, 5).map(([reason, count], i) => `
        <div class="flex items-center gap-2">
            <span class="w-3 h-3 rounded-full" style="background-color: ${colors[i]}"></span>
            <span class="flex-1 text-sm text-accent-700 truncate">${reason}</span>
            <span class="text-sm font-semibold text-accent-800">${count}人</span>
        </div>
    `).join('');
}

function updateAgeDistributionChart() {
    const filtered = getFilteredCustomerData();
    const ageGroups = { '10代': 0, '20代': 0, '30代': 0, '40代': 0, '50代': 0, '60代以上': 0 };
    const now = new Date();

    filtered.forEach(c => {
        const birthVal = c.birthday || c.birthDate;
        if (!birthVal) return;
        const birth = parseDate(birthVal);
        if (isNaN(birth.getTime())) return;
        const age = now.getFullYear() - birth.getFullYear();
        if (age < 20) ageGroups['10代']++;
        else if (age < 30) ageGroups['20代']++;
        else if (age < 40) ageGroups['30代']++;
        else if (age < 50) ageGroups['40代']++;
        else if (age < 60) ageGroups['50代']++;
        else ageGroups['60代以上']++;
    });

    // Gradient colors for age groups
    const ageColors = [
        '#d4b896', // 10代 - Light Gold
        '#c9a96e', // 20代 - Warm Gold
        '#b8956a', // 30代 - Champagne Gold (peak)
        '#739977', // 40代 - Sage Green
        '#6e819c', // 50代 - Light Slate
        '#47566b'  // 60代以上 - Slate Blue
    ];

    if (charts.ageDistribution) charts.ageDistribution.destroy();
    charts.ageDistribution = new Chart(document.getElementById('ageDistributionChart'), {
        type: 'bar',
        data: {
            labels: Object.keys(ageGroups),
            datasets: [{
                label: '人数',
                data: Object.values(ageGroups),
                backgroundColor: ageColors,
                borderRadius: 8,
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                y: { beginAtZero: true, grid: { color: '#edeae5' } },
                x: { grid: { display: false } }
            }
        }
    });
}

function updateOccupationChart() {
    const filtered = getFilteredCustomerData();
    const occCounts = {};
    filtered.forEach(c => {
        // job(Code.gs) または occupation の両方に対応
        const occ = c.job || c.occupation || '未回答';
        occCounts[occ] = (occCounts[occ] || 0) + 1;
    });

    const sorted = Object.entries(occCounts).sort((a, b) => b[1] - a[1]).slice(0, 8);
    // Elegant harmonious color palette for occupation
    const colors = [
        '#b8956a', // Champagne Gold
        '#47566b', // Slate Blue
        '#739977', // Sage Green
        '#b08f8a', // Dusty Rose
        '#c9a96e', // Warm Gold
        '#6e819c', // Light Slate
        '#8ba88e', // Light Sage
        '#c4a5a0'  // Light Rose
    ];

    if (charts.occupation) charts.occupation.destroy();
    charts.occupation = new Chart(document.getElementById('occupationChart'), {
        type: 'pie',
        data: {
            labels: sorted.map(([k]) => k),
            datasets: [{
                data: sorted.map(([, v]) => v),
                backgroundColor: colors,
                borderWidth: 2,
                borderColor: '#fff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right', labels: { boxWidth: 12, font: { size: 11 } } }
            }
        }
    });
}

function updateGeographicAnalysis() {
    const filtered = getFilteredCustomerData();
    const areaCounts = {};
    filtered.forEach(c => {
        if (!c.address) return;
        // 都道府県と市区町村を抽出
        const match = c.address.match(/^(.+?[都道府県])(.+?[市区町村])?/);
        if (match) {
            const area = match[2] ? match[1] + match[2] : match[1];
            areaCounts[area] = (areaCounts[area] || 0) + 1;
        }
    });

    const sorted = Object.entries(areaCounts).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const max = sorted[0]?.[1] || 1;

    document.getElementById('geographic-analysis').innerHTML = sorted.map(([area, count]) => `
        <div class="flex items-center gap-3">
            <span class="w-32 text-sm text-accent-700 truncate">${area}</span>
            <div class="flex-1 bg-surface-200 h-3 rounded-full overflow-hidden">
                <div class="h-full bg-gradient-to-r from-primary-400 to-primary-500 rounded-full" style="width: ${(count / max) * 100}%"></div>
            </div>
            <span class="text-sm font-semibold text-accent-800 w-12 text-right">${count}人</span>
        </div>
    `).join('');
}

function updateMenuPreferences() {
    const filtered = getFilteredCustomerData();

    // 眉毛のお悩み (eyebrowConcern(Code.gs) または eyebrowConcerns の両方に対応)
    const eyebrowConcerns = {};
    filtered.forEach(c => {
        const concernVal = c.eyebrowConcern || c.eyebrowConcerns;
        if (concernVal) {
            concernVal.split(',').forEach(concern => {
                const t = concern.trim();
                if (t) eyebrowConcerns[t] = (eyebrowConcerns[t] || 0) + 1;
            });
        }
    });
    const sortedConcerns = Object.entries(eyebrowConcerns).sort((a, b) => b[1] - a[1]).slice(0, 5);
    const maxConcern = sortedConcerns[0]?.[1] || 1;
    document.getElementById('eyebrow-concerns-list').innerHTML = sortedConcerns.map(([concern, count]) => `
        <div class="flex items-center gap-2">
            <div class="flex-1 bg-surface-200 h-2 rounded-full overflow-hidden">
                <div class="h-full bg-[#b08f8a] rounded-full" style="width: ${(count / maxConcern) * 100}%"></div>
            </div>
            <span class="text-xs text-accent-700 w-24 truncate">${concern}</span>
            <span class="text-xs font-semibold text-accent-800">${count}</span>
        </div>
    `).join('') || '<p class="text-xs text-surface-500">データなし</p>';

    // 眉毛の希望印象
    const impressions = {};
    filtered.forEach(c => {
        if (c.eyebrowImpression) {
            impressions[c.eyebrowImpression] = (impressions[c.eyebrowImpression] || 0) + 1;
        }
    });
    document.getElementById('eyebrow-impression-list').innerHTML = Object.entries(impressions)
        .sort((a, b) => b[1] - a[1]).slice(0, 6)
        .map(([imp, count]) => `<span class="px-3 py-1 bg-[#f5ebe7] text-[#b08f8a] rounded-full text-xs">${imp} (${count})</span>`).join('') || '<p class="text-xs text-surface-500">データなし</p>';

    // まつ毛デザイン (lashDesign(Code.gs) または eyelashDesign の両方に対応)
    const eyelashDesigns = {};
    filtered.forEach(c => {
        const designVal = c.lashDesign || c.eyelashDesign;
        if (designVal) {
            eyelashDesigns[designVal] = (eyelashDesigns[designVal] || 0) + 1;
        }
    });
    const sortedDesigns = Object.entries(eyelashDesigns).sort((a, b) => b[1] - a[1]).slice(0, 5);
    const maxDesign = sortedDesigns[0]?.[1] || 1;
    document.getElementById('eyelash-design-list').innerHTML = sortedDesigns.map(([design, count]) => `
        <div class="flex items-center gap-2">
            <div class="flex-1 bg-surface-200 h-2 rounded-full overflow-hidden">
                <div class="h-full bg-[#b8956a] rounded-full" style="width: ${(count / maxDesign) * 100}%"></div>
            </div>
            <span class="text-xs text-accent-700 w-24 truncate">${design}</span>
            <span class="text-xs font-semibold text-accent-800">${count}</span>
        </div>
    `).join('') || '<p class="text-xs text-surface-500">データなし</p>';

    // 目の見え方 (lashEyeLook(Code.gs) または eyelashEyeLook の両方に対応)
    const eyeLooks = {};
    filtered.forEach(c => {
        const eyeLookVal = c.lashEyeLook || c.eyelashEyeLook;
        if (eyeLookVal) {
            eyeLooks[eyeLookVal] = (eyeLooks[eyeLookVal] || 0) + 1;
        }
    });
    document.getElementById('eyelash-eyelook-list').innerHTML = Object.entries(eyeLooks)
        .sort((a, b) => b[1] - a[1]).slice(0, 6)
        .map(([look, count]) => `<span class="px-3 py-1 bg-[#faf5f0] text-[#b8956a] rounded-full text-xs">${look} (${count})</span>`).join('') || '<p class="text-xs text-surface-500">データなし</p>';
}

function updateCustomerList() {
    renderCustomerList();
}

function renderCustomerList() {
    const searchTerm = document.getElementById('customer-search')?.value.toLowerCase() || '';
    const snsFilter = document.getElementById('customer-filter-sns')?.value || 'all';

    // 店舗フィルターを適用したデータを使用
    let filtered = getFilteredCustomerData().filter(c => {
        const matchesSearch = !searchTerm || (c.name && c.name.toLowerCase().includes(searchTerm));
        const snsVal = c.snsOk || c.snsPermission || '';
        const hasSns = snsVal.includes('はい') || snsVal.includes('OK') || snsVal.includes('許可');
        const matchesSns = snsFilter === 'all' || (snsFilter === 'yes' && hasSns) || (snsFilter === 'no' && !hasSns);
        return matchesSearch && matchesSns;
    });

    // Sort by timestamp descending
    filtered.sort((a, b) => parseDate(b.timestamp || 0) - parseDate(a.timestamp || 0));

    const total = filtered.length;
    const startIdx = (customerListCurrentPage - 1) * customerListPageSize;
    const pageData = filtered.slice(startIdx, startIdx + customerListPageSize);

    const now = new Date();
    const tbody = document.getElementById('customer-list-body');
    tbody.innerHTML = pageData.map(c => {
        const regDate = c.timestamp ? parseDate(c.timestamp).toLocaleDateString('ja-JP') : '-';
        let age = '-';
        const birthVal = c.birthday || c.birthDate;
        if (birthVal) {
            const birth = parseDate(birthVal);
            if (!isNaN(birth.getTime())) {
                age = now.getFullYear() - birth.getFullYear() + '歳';
            }
        }
        const area = c.address ? (c.address.match(/^(.+?[都道府県])(.+?[市区町村])?/)?.[0] || c.address.substring(0, 10)) : '-';
        const snsVal = c.snsOk || c.snsPermission || '';
        const hasSns = snsVal.includes('はい') || snsVal.includes('OK') || snsVal.includes('許可');
        const reason = c.visitReason ? (c.visitReason.length > 20 ? c.visitReason.substring(0, 20) + '...' : c.visitReason) : '-';
        const occupation = c.job || c.occupation || '-';

        return `
            <tr class="border-b border-surface-100 hover:bg-surface-50 transition">
                <td class="py-3 px-4 text-surface-600">${regDate}</td>
                <td class="py-3 px-4 font-medium text-accent-800">${c.name || '-'}</td>
                <td class="py-3 px-4">${age}</td>
                <td class="py-3 px-4 text-surface-600">${occupation}</td>
                <td class="py-3 px-4 text-surface-600">${area}</td>
                <td class="py-3 px-4 text-center">
                    ${hasSns ? '<span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-emerald-100 text-emerald-600"><i data-lucide="check" class="w-3 h-3"></i></span>' : '<span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-surface-100 text-surface-400"><i data-lucide="x" class="w-3 h-3"></i></span>'}
                </td>
                <td class="py-3 px-4 text-surface-600 text-xs">${reason}</td>
            </tr>
        `;
    }).join('') || '<tr><td colspan="7" class="text-center py-8 text-surface-500">データがありません</td></tr>';

    document.getElementById('customer-list-count').textContent = `${total}件中 ${startIdx + 1}-${Math.min(startIdx + customerListPageSize, total)}件表示`;
    document.getElementById('customer-list-page').textContent = customerListCurrentPage;

    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function filterCustomerList() {
    customerListCurrentPage = 1;
    renderCustomerList();
}

function customerListPage(delta) {
    const searchTerm = document.getElementById('customer-search')?.value.toLowerCase() || '';
    const snsFilter = document.getElementById('customer-filter-sns')?.value || 'all';

    const filtered = customerData.filter(c => {
        const matchesSearch = !searchTerm || (c.name && c.name.toLowerCase().includes(searchTerm));
        const hasSns = c.snsPermission && (c.snsPermission.includes('はい') || c.snsPermission.includes('OK'));
        const matchesSns = snsFilter === 'all' || (snsFilter === 'yes' && hasSns) || (snsFilter === 'no' && !hasSns);
        return matchesSearch && matchesSns;
    });

    const maxPage = Math.ceil(filtered.length / customerListPageSize);
    const newPage = customerListCurrentPage + delta;
    if (newPage >= 1 && newPage <= maxPage) {
        customerListCurrentPage = newPage;
        renderCustomerList();
    }
}

// カウンセリング回答結果の更新
async function refreshCounselingResults() {
    const refreshIcon = document.getElementById('counseling-refresh-icon');
    if (refreshIcon) refreshIcon.classList.add('animate-spin');

    const loaded = await loadCustomerData();
    renderCounselingResults();

    if (refreshIcon) refreshIcon.classList.remove('animate-spin');

    if (loaded) {
        showUpdateNotification('カウンセリング回答を更新しました');
    }
}

// カウンセリング回答結果の表示
function renderCounselingResults() {
    const container = document.getElementById('counseling-results-container');
    if (!container) return;

    if (!customerData || customerData.length === 0) {
        container.innerHTML = `
            <div class="text-center py-8 text-mavie-400">
                <i data-lucide="clipboard-list" class="w-12 h-12 mx-auto mb-3 opacity-50"></i>
                <p class="text-sm">顧客データAPIを設定すると、カウンセリング回答が表示されます</p>
                <p class="text-xs mt-1">設定タブ → 顧客データ連携設定</p>
            </div>
        `;
        if (typeof lucide !== 'undefined') lucide.createIcons();
        return;
    }

    // スタッフ専用の場合は所属店舗のみ表示
    const storeFilter = lockedStore || (document.getElementById('counseling-store-filter')?.value || 'all');
    const searchTerm = document.getElementById('counseling-result-search')?.value.toLowerCase() || '';
    const dateFilter = document.getElementById('counseling-date-filter')?.value || 'month';

    // スタッフ専用の場合、店舗フィルターを無効化
    const storeFilterElement = document.getElementById('counseling-store-filter');
    if (lockedStore && storeFilterElement) {
        storeFilterElement.value = lockedStore;
        storeFilterElement.disabled = true;
    }

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

    let filtered = customerData.filter(c => {
        // 店舗フィルター（スタッフ専用の場合は所属店舗のみ）
        if (storeFilter !== 'all' && c.store !== storeFilter) return false;

        // 検索フィルター
        if (searchTerm && !(c.name && c.name.toLowerCase().includes(searchTerm))) return false;

        // 日付フィルター
        if (dateFilter !== 'all' && c.timestamp) {
            const recordDate = parseDate(c.timestamp);
            if (dateFilter === 'today' && recordDate < today) return false;
            if (dateFilter === 'week' && recordDate < weekAgo) return false;
            if (dateFilter === 'month' && recordDate < monthStart) return false;
        }

        return true;
    });

    // 日付の新しい順にソート
    filtered.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

    // 件数表示
    document.getElementById('counseling-result-count').textContent = `${filtered.length}件の回答`;

    if (filtered.length === 0) {
        container.innerHTML = `
            <div class="text-center py-8 text-mavie-400">
                <i data-lucide="search-x" class="w-12 h-12 mx-auto mb-3 opacity-50"></i>
                <p class="text-sm">該当する回答がありません</p>
            </div>
        `;
        if (typeof lucide !== 'undefined') lucide.createIcons();
        return;
    }

    container.innerHTML = filtered.map(c => {
        const regDate = c.timestamp ? parseDate(c.timestamp).toLocaleDateString('ja-JP', {year: 'numeric', month: 'short', day: 'numeric'}) : '-';
        const storeName = c.store === 'chiba' ? '千葉店' : c.store === 'honatsugi' ? '本厚木店' : c.store === 'yamato' ? '大和店' : c.storeName || '-';
        const birthVal = c.birthday || c.birthDate;
        let age = '-';
        if (birthVal) {
            const birth = parseDate(birthVal);
            if (!isNaN(birth.getTime())) {
                age = now.getFullYear() - birth.getFullYear();
            }
        }
        const snsVal = c.snsOk || c.snsPermission || '';
        const hasSns = snsVal.includes('はい') || snsVal.includes('OK') || snsVal.includes('許可');

        // 眉毛メニュー情報があるか
        const hasEyebrowInfo = c.eyebrowFrequency || c.eyebrowLastCare || c.eyebrowConcern || c.eyebrowDesign || c.eyebrowDesignImage || c.eyebrowImpression || c.eyebrowTrouble;
        // まつ毛メニュー情報があるか
        const hasLashInfo = c.lashFrequency || c.lashDesign || c.lashDesignImage || c.lashEyeLook || c.lashContact || c.lashTrouble;
        // その他情報があるか
        const hasOtherInfo = c.allergy || c.fromOtherSalon || c.dissatisfaction;

        return `
            <div class="bg-white border border-mavie-200 rounded-lg shadow-sm hover:shadow-md transition overflow-hidden">
                <!-- ヘッダー: 名前・日付・店舗 -->
                <div class="bg-gradient-to-r ${c.store === 'chiba' ? 'from-blue-50 to-blue-100' : 'from-purple-50 to-purple-100'} px-4 py-3 border-b border-mavie-200">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <div class="flex items-center gap-3">
                            <div class="w-12 h-12 rounded-full ${c.store === 'chiba' ? 'bg-blue-200 text-blue-700' : 'bg-purple-200 text-purple-700'} flex items-center justify-center font-bold text-lg">
                                ${c.name ? c.name.charAt(0) : '?'}
                            </div>
                            <div>
                                <h4 class="font-bold text-mavie-800 text-lg">${c.name || '名前未入力'}</h4>
                                <p class="text-sm text-mavie-500">${c.nameKana || ''}</p>
                            </div>
                        </div>
                        <div class="flex items-center gap-2">
                            <span class="px-3 py-1 text-sm font-semibold rounded-full ${c.store === 'chiba' ? 'bg-blue-500 text-white' : 'bg-purple-500 text-white'}">${storeName}</span>
                            <span class="text-sm text-mavie-500 bg-white px-2 py-1 rounded">${regDate}</span>
                        </div>
                    </div>
                </div>

                <div class="p-4 space-y-4">
                    <!-- 基本情報 -->
                    <div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-3 text-sm bg-mavie-50 p-3 rounded-lg">
                        <div>
                            <span class="text-xs text-mavie-400 block">年齢</span>
                            <span class="font-semibold text-mavie-700">${age !== '-' ? age + '歳' : '-'}</span>
                        </div>
                        <div>
                            <span class="text-xs text-mavie-400 block">職業</span>
                            <span class="font-semibold text-mavie-700">${c.job || c.occupation || '-'}</span>
                        </div>
                        <div>
                            <span class="text-xs text-mavie-400 block">電話番号</span>
                            <span class="font-semibold text-mavie-700">${c.phone || '-'}</span>
                        </div>
                        <div>
                            <span class="text-xs text-mavie-400 block">住所</span>
                            <span class="font-semibold text-mavie-700 text-xs">${c.address || '-'}</span>
                        </div>
                        <div>
                            <span class="text-xs text-mavie-400 block">生年月日</span>
                            <span class="font-semibold text-mavie-700 text-xs">${birthVal || '-'}</span>
                        </div>
                        <div>
                            <span class="text-xs text-mavie-400 block">SNS許可</span>
                            <span class="font-semibold ${hasSns ? 'text-emerald-600' : 'text-red-500'}">${c.snsOk || c.snsPermission || '-'}</span>
                        </div>
                    </div>

                    <!-- 来店理由 -->
                    ${c.visitReason ? `
                    <div class="bg-amber-50 p-3 rounded-lg border-l-4 border-amber-400">
                        <span class="text-xs font-bold text-amber-700 block mb-1">来店理由</span>
                        <p class="text-sm text-mavie-700">${c.visitReason}</p>
                    </div>
                    ` : ''}

                    <!-- 他サロン・不満点 -->
                    ${(c.fromOtherSalon || c.dissatisfaction) ? `
                    <div class="bg-orange-50 p-3 rounded-lg border-l-4 border-orange-400">
                        <span class="text-xs font-bold text-orange-700 block mb-2">他サロン情報</span>
                        <div class="grid grid-cols-1 sm:grid-cols-2 gap-2 text-sm">
                            ${c.fromOtherSalon ? `<div><span class="text-mavie-400">他サロン利用:</span> <span class="text-mavie-700">${c.fromOtherSalon}</span></div>` : ''}
                            ${c.dissatisfaction ? `<div><span class="text-mavie-400">不満点:</span> <span class="text-mavie-700">${c.dissatisfaction}</span></div>` : ''}
                        </div>
                    </div>
                    ` : ''}

                    <!-- アレルギー情報 -->
                    ${c.allergy ? `
                    <div class="bg-red-50 p-3 rounded-lg border-l-4 border-red-500">
                        <span class="text-xs font-bold text-red-600 block mb-1">アレルギー情報</span>
                        <p class="text-sm text-red-700 font-semibold">${c.allergy}</p>
                    </div>
                    ` : ''}

                    <!-- 眉毛・まつ毛メニュー情報 -->
                    <div class="grid grid-cols-1 lg:grid-cols-2 gap-4">
                        ${hasEyebrowInfo ? `
                        <div class="bg-[#faf5f2] p-4 rounded-lg border border-[#e8d4cd]">
                            <h5 class="text-sm font-bold text-[#b08f8a] mb-3 flex items-center gap-2 pb-2 border-b border-[#e8d4cd]">
                                <i data-lucide="eye" class="w-4 h-4"></i>眉毛メニュー
                            </h5>
                            <div class="space-y-2 text-sm">
                                ${c.eyebrowFrequency ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">利用頻度:</span><span class="text-mavie-700">${c.eyebrowFrequency}</span></div>` : ''}
                                ${c.eyebrowLastCare ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">最後のお手入れ:</span><span class="text-mavie-700">${c.eyebrowLastCare}</span></div>` : ''}
                                ${c.eyebrowConcern ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">お悩み:</span><span class="text-mavie-700">${c.eyebrowConcern}</span></div>` : ''}
                                ${c.eyebrowDesign ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">希望デザイン:</span><span class="text-mavie-700">${c.eyebrowDesign}</span></div>` : ''}
                                ${c.eyebrowDesignImage ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">デザインイメージ:</span><span class="text-mavie-700 font-medium">${c.eyebrowDesignImage}</span></div>` : ''}
                                ${c.eyebrowImpression ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">希望印象:</span><span class="text-mavie-700">${c.eyebrowImpression}</span></div>` : ''}
                                ${c.eyebrowTrouble ? `<div class="flex"><span class="text-red-400 w-28 shrink-0">施術後トラブル:</span><span class="text-red-600">${c.eyebrowTrouble}</span></div>` : ''}
                            </div>
                        </div>
                        ` : ''}
                        ${hasLashInfo ? `
                        <div class="bg-[#faf8f5] p-4 rounded-lg border border-[#e8ddd0]">
                            <h5 class="text-sm font-bold text-[#b8956a] mb-3 flex items-center gap-2 pb-2 border-b border-[#e8ddd0]">
                                <i data-lucide="sparkles" class="w-4 h-4"></i>まつ毛メニュー
                            </h5>
                            <div class="space-y-2 text-sm">
                                ${c.lashFrequency ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">利用頻度:</span><span class="text-mavie-700">${c.lashFrequency}</span></div>` : ''}
                                ${c.lashDesign ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">希望デザイン:</span><span class="text-mavie-700">${c.lashDesign}</span></div>` : ''}
                                ${c.lashDesignImage ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">デザインイメージ:</span><span class="text-mavie-700 font-medium">${c.lashDesignImage}</span></div>` : ''}
                                ${c.lashEyeLook ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">目の見え方:</span><span class="text-mavie-700">${c.lashEyeLook}</span></div>` : ''}
                                ${c.lashContact ? `<div class="flex"><span class="text-mavie-400 w-28 shrink-0">コンタクト:</span><span class="text-mavie-700">${c.lashContact}</span></div>` : ''}
                                ${c.lashTrouble ? `<div class="flex"><span class="text-red-400 w-28 shrink-0">施術後トラブル:</span><span class="text-red-600">${c.lashTrouble}</span></div>` : ''}
                            </div>
                        </div>
                        ` : ''}
                    </div>

                    <!-- 同意確認 -->
                    ${c.agreement ? `
                    <div class="text-xs text-mavie-400 pt-2 border-t border-mavie-100">
                        <span class="font-semibold">注意事項確認:</span> ${c.agreement}
                    </div>
                    ` : ''}
                </div>
            </div>
        `;
    }).join('');

    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function checkUrlParams() {
    const params = new URLSearchParams(window.location.search);
    // URLパラメータを小文字に正規化（STAFF_ROSTERと一致させるため）
    const pStore = (params.get('store') || '').toLowerCase() || null;
    const pStaff = (params.get('staff') || '').toLowerCase() || null;
    const storeSel = document.getElementById('store-selector');
    const staffSel = document.getElementById('staff-selector');
    const staffModeBadge = document.getElementById('staff-mode-badge');

    if (pStore && STAFF_ROSTER[pStore]) {
        lockedStore = pStore;
        storeSel.value = pStore;
        storeSel.disabled = true;
        staffSel.innerHTML = '<option value="all">店舗合計</option>';

        // 大文字小文字を区別しないでスタッフ名を照合
        const matchedStaff = pStaff && STAFF_ROSTER[pStore].find(s => s.toLowerCase() === pStaff.toLowerCase());
        if (matchedStaff) {
            lockedStaff = matchedStaff; // STAFF_ROSTERの正式な名前を使用
            const opt = document.createElement('option');
            opt.value = matchedStaff;
            opt.text = matchedStaff;
            staffSel.appendChild(opt);
            staffSel.value = matchedStaff;

            // グレーアウトしたセレクターを非表示にする
            storeSel.parentElement.style.display = 'none';
            staffSel.parentElement.style.display = 'none';

            // Show staff mode badge
            staffModeBadge.classList.remove('hidden');
            staffModeBadge.innerText = `${matchedStaff} 専用`;

            // スタッフ専用ページ：不要なタブを非表示にする
            const hideTabs = ['tab-overview', 'tab-customers', 'tab-kpi', 'tab-marketing', 'tab-goal', 'tab-settings', 'tab-incentive'];
            hideTabs.forEach(tabId => {
                const tab = document.getElementById(tabId);
                if (tab) tab.style.display = 'none';
            });

            // スタッフ専用ページ：タブ順序を変更（マイダッシュボード→カウンセリング回答→売上詳細→カレンダー→データ編集）
            const tabContainer = document.getElementById('main-tabs');
            if (tabContainer) {
                const tabOrder = ['tab-staff-dashboard', 'tab-counseling-results', 'tab-sales', 'tab-calendar', 'tab-edit'];
                tabOrder.forEach((tabId, index) => {
                    const tab = document.getElementById(tabId);
                    if (tab) {
                        tab.style.order = index;
                        tab.classList.remove('hidden');
                    }
                });
            }

            // Hide AI advice section for staff-specific URLs
            const aiAdviceSection = document.getElementById('ai-advice-section');
            if (aiAdviceSection) aiAdviceSection.style.display = 'none';

            // スタッフ専用：編集テーブルのプレビューボタンを非表示
            const previewBtn = document.getElementById('btn-preview-edit');
            if (previewBtn) previewBtn.style.display = 'none';

            // スタッフ専用：編集テーブルの店舗・スタッフ列を非表示
            const editStoreHeader = document.getElementById('edit-th-store');
            const editStaffHeader = document.getElementById('edit-th-staff');
            if (editStoreHeader) editStoreHeader.style.display = 'none';
            if (editStaffHeader) editStaffHeader.style.display = 'none';

            // スタッフ専用ページはマイダッシュボードを初期タブとして開く
            setTimeout(() => switchTab('staff-dashboard'), 100);

            // スタッフ専用ページはパスワード認証が必要（matchedStaffを使用）
            const staffPassword = getStaffPassword(pStore, matchedStaff);
            if (staffPassword !== null && staffPassword !== '') {
                // パスワードが設定されている場合はログインモーダルを表示
                showStaffLoginModal(pStore, matchedStaff);
            } else {
                // パスワード未設定の場合は認証済みとする
                isStaffAuthenticated = true;
            }
        } else {
            // === 店舗専用モード（?store=xxx のみ、staffなし） ===
            // スタッフ選択肢を追加（店舗合計 + 各スタッフ）
            STAFF_ROSTER[pStore].forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.text = name;
                staffSel.appendChild(opt);
            });

            // 店舗モードかどうかを判定（staffパラメータが明示的に指定されていない場合のみ）
            if (!pStaff) {
                // 店舗セレクターを非表示（固定のため）
                storeSel.parentElement.style.display = 'none';

                // バッジ表示
                const storeName = pStore === 'chiba' ? '千葉店' : pStore === 'honatsugi' ? '本厚木店' : pStore === 'yamato' ? '大和店' : pStore;
                staffModeBadge.classList.remove('hidden');
                staffModeBadge.innerText = `${storeName} 管理`;

                // 不要なタブを非表示（インセンティブ・マーケティング・目標設定・設定）
                const hideTabs = ['tab-marketing', 'tab-goal', 'tab-settings', 'tab-kpi', 'tab-incentive'];
                hideTabs.forEach(tabId => {
                    const tab = document.getElementById(tabId);
                    if (tab) tab.style.display = 'none';
                });

                // タブ順序を設定（サマリー→カウンセリング→売上詳細→カレンダー→データ編集）
                const tabContainer = document.getElementById('main-tabs');
                if (tabContainer) {
                    const tabOrder = ['tab-overview', 'tab-counseling-results', 'tab-sales', 'tab-calendar', 'tab-edit'];
                    tabOrder.forEach((tabId, index) => {
                        const tab = document.getElementById(tabId);
                        if (tab) {
                            tab.style.order = index;
                            tab.classList.remove('hidden');
                        }
                    });
                    // 顧客・媒体分析タブも表示
                    const custTab = document.getElementById('tab-customers');
                    if (custTab) custTab.style.order = 5;
                }

                // AI adviceセクション非表示
                const aiAdviceSection = document.getElementById('ai-advice-section');
                if (aiAdviceSection) aiAdviceSection.style.display = 'none';

                // 編集テーブルの店舗列を非表示（店舗は固定のため）
                const editStoreHeader = document.getElementById('edit-th-store');
                if (editStoreHeader) editStoreHeader.style.display = 'none';

                // サマリータブを初期表示
                setTimeout(() => switchTab('overview'), 100);
            }
        }
    } else {
        handleStoreChange(false);
    }
}

let currentPeriodFilter = 'month'; // 'month', '3months', '6months', 'year'

function getFilteredData() {
    // rawDataが配列でない場合は空配列を返す
    if (!Array.isArray(rawData)) {
        console.warn('getFilteredData: rawDataが配列ではありません');
        return [];
    }

    const storeFilter = document.getElementById('store-selector')?.value || 'all';
    const staffFilter = document.getElementById('staff-selector')?.value || 'all';
    const dateSelector = document.getElementById('date-selector');
    const selectedDate = dateSelector ? dateSelector.value : null;

    const result = rawData.filter(d => {
        if (!d || !d.date) return false;
        if (storeFilter !== 'all' && d.store !== storeFilter) return false;
        // スタッフ名の比較は大文字小文字を区別しない
        if (staffFilter !== 'all' && d.staff?.toLowerCase() !== staffFilter.toLowerCase()) return false;

        // 期間フィルタの適用
        const recordDate = parseDate(d.date);
        const today = new Date();
        let startDate;

        switch (currentPeriodFilter) {
            case 'month':
                // selectedDateがない場合は現在の月を使用
                const dateStr = selectedDate || `${today.getFullYear()}/${today.getMonth() + 1}`;
                const [year, month] = dateStr.split('/').map(Number);
                return recordDate.getFullYear() === year && (recordDate.getMonth() + 1) === month;
            case '3months':
                startDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);
                return recordDate >= startDate;
            case '6months':
                startDate = new Date(today.getFullYear(), today.getMonth() - 5, 1);
                return recordDate >= startDate;
            case 'year':
                startDate = new Date(today.getFullYear() - 1, today.getMonth(), 1);
                return recordDate >= startDate;
            default:
                return true;
        }
    });

    return result;
}

function setPeriodFilter(period) {
    currentPeriodFilter = period;

    // ボタンのスタイルを更新
    document.querySelectorAll('.period-btn').forEach(btn => {
        btn.classList.remove('active');
        btn.classList.add('text-surface-600', 'hover:bg-surface-200');
    });
    const activeBtn = document.getElementById(`period-${period}`);
    if (activeBtn) {
        activeBtn.classList.remove('text-surface-600', 'hover:bg-surface-200');
        activeBtn.classList.add('active');
    }

    // 日付セレクタの表示/非表示を切り替え
    const dateSelectorWrapper = document.getElementById('date-selector').parentElement;
    if (period === 'month') {
        dateSelectorWrapper.style.display = '';
    } else {
        dateSelectorWrapper.style.display = 'none';
    }

    // ダッシュボードを更新
    updateDashboard();
}

function calculateMetrics(data) {
    let agg = {
        salesTotal: 0, salesCash: 0, salesCredit: 0, salesQR: 0,
        customersTotal: 0, customersNew: 0, customersExisting: 0,
        newByChannel: { hpb: 0, mininai: 0 },
        nextRes: { hpbNew: 0, mininaiNew: 0, existing: 0, total: 0 },
        hpbNewCount: 0, lossTotal: 0, reviews5StarTotal: 0, blogUpdatesTotal: 0, snsUpdatesTotal: 0, daily: {}
    };

    // dataが配列でない場合は空の結果を返す
    if (!Array.isArray(data)) {
        console.warn('calculateMetrics: dataが配列ではありません', data);
        return agg;
    }

    data.forEach(d => {
        if (!d || !d.date) return; // 無効なデータをスキップ
        if (!agg.daily[d.date]) agg.daily[d.date] = { sales: 0, customers: 0, new: 0, existing: 0, hpb: 0, mininai: 0 };
    });

    data.forEach(d => {
        // 無効なデータをスキップ
        if (!d) return;

        // オブジェクトの存在チェックとデフォルト値の設定
        const sales = d.sales || {};
        const discounts = d.discounts || {};
        const customers = d.customers || {};
        const nextRes = d.nextRes || {};

        // 売上 = 現金 + クレジット + QR + HPBポイント + HPBギフト券
        const salesCash = sales.cash || 0;
        const salesCredit = sales.credit || 0;
        const salesQR = sales.qr || 0;
        const hpbPoints = discounts.hpbPoints || 0;
        const hpbGift = discounts.hpbGift || 0;
        const totalSales = salesCash + salesCredit + salesQR + hpbPoints + hpbGift;

        agg.salesTotal += totalSales;
        agg.salesCash += salesCash;
        agg.salesCredit += salesCredit;
        agg.salesQR += salesQR;
        // 損失 = その他割引 + 返金（HPBポイントとギフト券は売上に含まれるため除外）
        agg.lossTotal += ((discounts.other || 0) + (discounts.refund || 0));

        // 既存客には知り合い価格も含める
        const custNewHPB = customers.newHPB || 0;
        const custNewMiniNai = customers.newMiniNai || 0;
        const existingCount = (customers.existing || 0) + (customers.acquaintance || 0);
        const totalCust = custNewHPB + custNewMiniNai + existingCount;
        const newCust = custNewHPB + custNewMiniNai;

        agg.customersTotal += totalCust;
        agg.customersNew += newCust;
        agg.customersExisting += existingCount;
        agg.newByChannel.hpb += custNewHPB;
        agg.newByChannel.mininai += custNewMiniNai;

        const nextResNewHPB = nextRes.newHPB || 0;
        const nextResNewMiniNai = nextRes.newMiniNai || 0;
        const nextResExisting = nextRes.existing || 0;

        agg.nextRes.hpbNew += nextResNewHPB;
        agg.nextRes.mininaiNew += nextResNewMiniNai;
        agg.nextRes.existing += nextResExisting;
        agg.nextRes.total += (nextResNewHPB + nextResNewMiniNai + nextResExisting);
        agg.hpbNewCount += custNewHPB;
        agg.reviews5StarTotal += (d.reviews5Star || 0);
        agg.blogUpdatesTotal += (d.blogUpdates || 0);
        agg.snsUpdatesTotal += (d.snsUpdates || 0);

        if(d.date && agg.daily[d.date]) {
            agg.daily[d.date].sales += totalSales;
            agg.daily[d.date].customers += totalCust;
            agg.daily[d.date].new += newCust;
            agg.daily[d.date].existing += existingCount;
            agg.daily[d.date].hpb += custNewHPB;
            agg.daily[d.date].mininai += custNewMiniNai;
        }
    });
    return agg;
}

// --- 3. UI LOGIC ---
function handleStoreChange(isUserAction = true) {
    if (isUserAction && lockedStore) return;
    const storeSel = document.getElementById('store-selector');
    const staffSel = document.getElementById('staff-selector');
    const selectedStore = storeSel.value;

    staffSel.innerHTML = '<option value="all">店舗合計</option>';
    if (selectedStore === 'all') {
        staffSel.disabled = true;
    } else {
        staffSel.disabled = false;
        if(STAFF_ROSTER[selectedStore]){
            STAFF_ROSTER[selectedStore].forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.text = name;
                staffSel.appendChild(opt);
            });
        }
    }
    if(isUserAction) updateDashboard();
}

// ===== Bottom Sheet (More menu) =====
function openMoreSheet() {
    const sheet = document.getElementById('more-sheet');
    const backdrop = document.getElementById('more-sheet-backdrop');
    if (!sheet || !backdrop) return;
    sheet.classList.add('open');
    sheet.setAttribute('aria-hidden', 'false');
    backdrop.classList.add('open');
    document.body.style.overflow = 'hidden';
    if (window.lucide) lucide.createIcons();
}
function closeMoreSheet() {
    const sheet = document.getElementById('more-sheet');
    const backdrop = document.getElementById('more-sheet-backdrop');
    if (!sheet || !backdrop) return;
    sheet.classList.remove('open');
    sheet.setAttribute('aria-hidden', 'true');
    backdrop.classList.remove('open');
    document.body.style.overflow = '';
}
function switchTabFromSheet(id) {
    closeMoreSheet();
    // Small delay for nicer transition
    setTimeout(() => switchTab(id), 160);
}
// ESC キーで閉じる
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeMoreSheet();
});

// ===== Overview Enhancement Helpers =====
// 前期（直前の同長期間）のデータを取得
function getPreviousPeriodData() {
    if (!Array.isArray(rawData)) return [];
    const storeFilter = document.getElementById('store-selector')?.value || 'all';
    const staffFilter = document.getElementById('staff-selector')?.value || 'all';
    const dateSelector = document.getElementById('date-selector');
    const selectedDate = dateSelector ? dateSelector.value : null;
    const today = new Date();

    let start, end;
    switch (currentPeriodFilter) {
        case 'month': {
            const dateStr = selectedDate || `${today.getFullYear()}/${today.getMonth() + 1}`;
            const [y, m] = dateStr.split('/').map(Number);
            // 前月 (m-1)
            start = new Date(y, m - 2, 1);
            end = new Date(y, m - 1, 0, 23, 59, 59);
            break;
        }
        case '3months':
            start = new Date(today.getFullYear(), today.getMonth() - 5, 1);
            end = new Date(today.getFullYear(), today.getMonth() - 2, 0, 23, 59, 59);
            break;
        case '6months':
            start = new Date(today.getFullYear(), today.getMonth() - 11, 1);
            end = new Date(today.getFullYear(), today.getMonth() - 5, 0, 23, 59, 59);
            break;
        case 'year':
            start = new Date(today.getFullYear() - 2, today.getMonth(), 1);
            end = new Date(today.getFullYear() - 1, today.getMonth(), 0, 23, 59, 59);
            break;
        default:
            return [];
    }

    return rawData.filter(d => {
        if (!d || !d.date) return false;
        if (storeFilter !== 'all' && d.store !== storeFilter) return false;
        if (staffFilter !== 'all' && d.staff?.toLowerCase() !== staffFilter.toLowerCase()) return false;
        const rd = parseDate(d.date);
        return rd >= start && rd <= end;
    });
}

// 前期比ラベル（上下矢印付き）を delta-badge にセット
function setDeltaBadge(id, current, previous, { unit = '%', invert = false } = {}) {
    const el = document.getElementById(id);
    if (!el) return;
    if (!previous || previous === 0) {
        el.className = 'delta-badge';
        el.textContent = '—';
        return;
    }
    const diffRatio = ((current - previous) / Math.abs(previous)) * 100;
    const rounded = Math.abs(diffRatio) >= 10 ? Math.round(diffRatio) : diffRatio.toFixed(1);
    const isUp = diffRatio > 0;
    const isFlat = Math.abs(diffRatio) < 0.1;
    const direction = isFlat ? 'flat' : (isUp ? 'up' : 'down');
    // invert: 下がった方が良い指標 (ロス等) は色を反転
    const colorClass = isFlat ? '' : ((invert ? !isUp : isUp) ? 'up' : 'down');
    el.className = `delta-badge ${colorClass}`;
    const arrow = isFlat ? '→' : (isUp ? '↑' : '↓');
    const sign = isFlat ? '' : (isUp ? '+' : '');
    el.innerHTML = `<span>${arrow}</span><span>${sign}${rounded}${unit === '%' ? '%' : ''}</span>`;
    el.title = `前期比 ${sign}${rounded}%`;
}

// シンプルなSVGスパークライン描画
function renderSparkline(containerId, values, colorStart, colorEnd) {
    const wrap = document.getElementById(containerId);
    if (!wrap) return;
    if (!values || values.length < 2) { wrap.innerHTML = ''; return; }
    const w = 120, h = 28, pad = 2;
    const max = Math.max(...values), min = Math.min(...values);
    const range = max - min || 1;
    const step = (w - pad * 2) / (values.length - 1);
    const points = values.map((v, i) => {
        const x = pad + step * i;
        const y = h - pad - ((v - min) / range) * (h - pad * 2);
        return [x, y];
    });
    const d = points.map((p, i) => (i === 0 ? `M${p[0]},${p[1]}` : `L${p[0]},${p[1]}`)).join(' ');
    const area = `${d} L${points[points.length - 1][0]},${h} L${points[0][0]},${h} Z`;
    const gradId = `spark-grad-${containerId}`;
    wrap.innerHTML = `
        <svg viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">
            <defs>
                <linearGradient id="${gradId}" x1="0" x2="0" y1="0" y2="1">
                    <stop offset="0%" stop-color="${colorStart}" stop-opacity="0.35"/>
                    <stop offset="100%" stop-color="${colorEnd}" stop-opacity="0"/>
                </linearGradient>
            </defs>
            <path d="${area}" fill="url(#${gradId})"/>
            <path d="${d}" fill="none" stroke="${colorStart}" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"/>
            <circle cx="${points[points.length - 1][0]}" cy="${points[points.length - 1][1]}" r="2.4" fill="${colorStart}"/>
        </svg>`;
}

// 日別データを日付順の配列に変換
function dailySeries(daily, key) {
    return Object.keys(daily || {})
        .sort((a, b) => parseDate(a) - parseDate(b))
        .map(k => daily[k][key] || 0);
}

// 数値のカウントアップ（要素の data-value を現在値、 text を更新）
function animateCount(elId, target, { prefix = '', suffix = '', decimals = 0, duration = 800 } = {}) {
    const el = document.getElementById(elId);
    if (!el) return;
    const start = parseFloat(el.dataset.value || '0') || 0;
    const end = Number(target) || 0;
    const wrapSuffix = s => {
        if (!s) return '';
        // % は インラインのまま（既存スタイルに合わせる）
        if (s === '%') return s;
        return `<span class="text-sm md:text-lg font-sans font-normal ml-0.5 md:ml-1 text-surface-500">${s}</span>`;
    };
    if (start === end) {
        const formatted = decimals > 0
            ? end.toFixed(decimals)
            : Math.round(end).toLocaleString();
        el.innerHTML = `${prefix}${formatted}${wrapSuffix(suffix)}`;
        el.dataset.value = String(end);
        return;
    }
    // Brief highlight to indicate value has updated
    el.classList.remove('value-flash');
    void el.offsetWidth; // restart animation
    el.classList.add('value-flash');
    const t0 = performance.now();
    const ease = t => 1 - Math.pow(1 - t, 3);
    function frame(now) {
        const p = Math.min((now - t0) / duration, 1);
        const v = start + (end - start) * ease(p);
        const formatted = decimals > 0
            ? v.toFixed(decimals)
            : Math.round(v).toLocaleString();
        el.innerHTML = `${prefix}${formatted}${wrapSuffix(suffix)}`;
        if (p < 1) requestAnimationFrame(frame);
        else el.dataset.value = String(end);
    }
    requestAnimationFrame(frame);
}

// 本日のスナップショット更新（期間フィルタに関係なく常に today を表示）
function updateTodaySnapshot(currentGoal) {
    try {
        const today = new Date();
        const key = `${today.getFullYear()}/${today.getMonth() + 1}/${today.getDate()}`;
        const storeFilter = document.getElementById('store-selector')?.value || 'all';
        const staffFilter = document.getElementById('staff-selector')?.value || 'all';

        // 本日分を直接集計
        const todayData = (Array.isArray(rawData) ? rawData : []).filter(d => {
            if (!d || !d.date) return false;
            if (storeFilter !== 'all' && d.store !== storeFilter) return false;
            if (staffFilter !== 'all' && d.staff?.toLowerCase() !== staffFilter.toLowerCase()) return false;
            const rd = parseDate(d.date);
            return rd.getFullYear() === today.getFullYear()
                && rd.getMonth() === today.getMonth()
                && rd.getDate() === today.getDate();
        });
        const m = calculateMetrics(todayData);

        // 日付表示
        const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
        const dateEl = document.getElementById('snapshot-date');
        if (dateEl) dateEl.textContent = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日 (${weekdays[today.getDay()]})`;

        // KPI
        const fmt = n => n.toLocaleString();
        const salesEl = document.getElementById('snapshot-sales');
        if (salesEl) salesEl.textContent = `¥${fmt(m.salesTotal)}`;
        const custEl = document.getElementById('snapshot-customers');
        if (custEl) custEl.innerHTML = `${fmt(m.customersTotal)}<span class="text-sm font-sans font-normal ml-1 text-surface-500">名</span>`;
        const nextResEl = document.getElementById('snapshot-nextres');
        if (nextResEl) nextResEl.innerHTML = `${fmt(m.nextRes.total)}<span class="text-sm font-sans font-normal ml-1 text-surface-500">件</span>`;

        // Sub text
        const unit = m.customersTotal > 0 ? Math.round(m.salesTotal / m.customersTotal) : 0;
        const resRate = m.customersTotal > 0 ? ((m.nextRes.total / m.customersTotal) * 100).toFixed(1) : '0.0';
        const salesSub = document.getElementById('snapshot-sales-sub');
        if (salesSub) salesSub.textContent = m.customersTotal > 0 ? `客単価 ¥${fmt(unit)}` : 'まだ売上記録なし';
        const custSub = document.getElementById('snapshot-customers-sub');
        if (custSub) custSub.textContent = `新規 ${m.customersNew} / 既存 ${m.customersExisting}`;
        const nextResSub = document.getElementById('snapshot-nextres-sub');
        if (nextResSub) nextResSub.textContent = `予約率 ${resRate}%`;

        // 月次ペース計算（常に今月ベース）
        const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
        const monthEnd = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        const daysInMonth = monthEnd.getDate();
        const dayOfMonth = today.getDate();
        const expectedPct = Math.round((dayOfMonth / daysInMonth) * 100);

        // 今月分を取得して進捗を出す
        const monthData = (Array.isArray(rawData) ? rawData : []).filter(d => {
            if (!d || !d.date) return false;
            if (storeFilter !== 'all' && d.store !== storeFilter) return false;
            if (staffFilter !== 'all' && d.staff?.toLowerCase() !== staffFilter.toLowerCase()) return false;
            const rd = parseDate(d.date);
            return rd >= monthStart && rd <= today;
        });
        const monthMetrics = calculateMetrics(monthData);
        const actualPct = currentGoal > 0 ? (monthMetrics.salesTotal / currentGoal) * 100 : 0;

        const fillEl = document.getElementById('pace-meter-fill');
        if (fillEl) {
            fillEl.style.width = `${Math.min(actualPct, 100).toFixed(1)}%`;
            fillEl.classList.remove('on-track', 'behind');
            if (currentGoal > 0) {
                if (actualPct >= expectedPct - 3) fillEl.classList.add('on-track');
                else if (actualPct < expectedPct - 10) fillEl.classList.add('behind');
            }
        }
        const markerEl = document.getElementById('pace-marker');
        if (markerEl) markerEl.style.left = `${Math.min(expectedPct, 100)}%`;
        const expEl = document.getElementById('pace-expected');
        if (expEl) expEl.textContent = `${expectedPct}%`;
        const statusEl = document.getElementById('pace-status');
        if (statusEl) {
            if (currentGoal <= 0) statusEl.textContent = '目標未設定';
            else {
                const diff = actualPct - expectedPct;
                if (Math.abs(diff) < 3) statusEl.textContent = `予定通り (${actualPct.toFixed(0)}% / ${expectedPct}%)`;
                else if (diff > 0) statusEl.textContent = `予定より +${diff.toFixed(0)}pt 好調`;
                else statusEl.textContent = `予定より ${diff.toFixed(0)}pt 遅れ`;
            }
        }
        // 着地予測: 現在ペースを月末まで延長
        const forecastEl = document.getElementById('pace-forecast');
        if (forecastEl) {
            const forecast = dayOfMonth > 0 ? Math.round(monthMetrics.salesTotal / dayOfMonth * daysInMonth) : 0;
            forecastEl.textContent = `¥${fmt(forecast)}`;
        }
    } catch (e) {
        console.error('updateTodaySnapshot エラー:', e);
    }
}

// スパークライン & Deltaバッジ更新
function updateOverviewExtras(metrics, currentGoal) {
    try {
        const prevData = getPreviousPeriodData();
        const prev = calculateMetrics(prevData);

        const curUnit = metrics.customersTotal > 0 ? metrics.salesTotal / metrics.customersTotal : 0;
        const prevUnit = prev.customersTotal > 0 ? prev.salesTotal / prev.customersTotal : 0;

        const curNewRes = metrics.customersNew > 0 ? ((metrics.nextRes.hpbNew + metrics.nextRes.mininaiNew) / metrics.customersNew) * 100 : 0;
        const prevNewRes = prev.customersNew > 0 ? ((prev.nextRes.hpbNew + prev.nextRes.mininaiNew) / prev.customersNew) * 100 : 0;

        const curTotalRes = metrics.customersTotal > 0 ? (metrics.nextRes.total / metrics.customersTotal) * 100 : 0;
        const prevTotalRes = prev.customersTotal > 0 ? (prev.nextRes.total / prev.customersTotal) * 100 : 0;

        setDeltaBadge('delta-sales', metrics.salesTotal, prev.salesTotal);
        setDeltaBadge('delta-customers', metrics.customersTotal, prev.customersTotal);
        setDeltaBadge('delta-unit-price', curUnit, prevUnit);
        setDeltaBadge('delta-new-res', curNewRes, prevNewRes);
        setDeltaBadge('delta-total-res', curTotalRes, prevTotalRes);

        // スパークライン（current metrics.daily から生成）
        const salesSeries = dailySeries(metrics.daily, 'sales');
        const custSeries = dailySeries(metrics.daily, 'customers');
        const unitSeries = Object.keys(metrics.daily || {})
            .sort((a, b) => parseDate(a) - parseDate(b))
            .map(k => {
                const d = metrics.daily[k];
                return d.customers > 0 ? d.sales / d.customers : 0;
            });

        renderSparkline('spark-sales', salesSeries, '#b8956a', '#c9a87e');
        renderSparkline('spark-customers', custSeries, '#566882', '#6e819c');
        renderSparkline('spark-unit-price', unitSeries, '#5d7d60', '#8ba88e');
    } catch (e) {
        console.error('updateOverviewExtras エラー:', e);
    }
}

function updateDashboard() {
    const filtered = getFilteredData();
    const metrics = calculateMetrics(filtered);
    const fmt = n => n.toLocaleString();

    // Get current goal based on context
    const context = getCurrentGoalContext();
    let currentGoalData;
    if (context.type === 'staff') {
        currentGoalData = getStaffGoal(context.store, context.staff);
    } else if (context.type === 'store') {
        currentGoalData = getStoreAggregateGoal(context.store);
    } else {
        currentGoalData = getAllStoresAggregateGoal();
    }
    const currentGoal = currentGoalData.weekdays * currentGoalData.weekdayTarget + currentGoalData.weekends * currentGoalData.weekendTarget;
    monthlyGoal = currentGoal; // Update global for backwards compatibility

    animateCount('kpi-sales', metrics.salesTotal, { prefix: '¥' });
    animateCount('kpi-customers', metrics.customersTotal, { suffix: '名' });
    document.getElementById('kpi-new').innerText = fmt(metrics.customersNew);
    document.getElementById('kpi-existing').innerText = fmt(metrics.customersExisting);

    // Update goal ratio
    const goalRatio = currentGoal > 0 ? Math.round((metrics.salesTotal / currentGoal) * 100) : 0;
    document.getElementById('kpi-goal-ratio').innerText = `${goalRatio}%`;

    const unitPrice = metrics.customersTotal > 0 ? Math.round(metrics.salesTotal / metrics.customersTotal) : 0;
    animateCount('kpi-unit-price', unitPrice, { prefix: '¥' });

    // New Customer Reservation Rate: (HPB new reservations + Mini/Nailie new reservations) / All new customers
    const allNewCustomers = metrics.customersNew;
    const newResRate = allNewCustomers > 0 ? (((metrics.nextRes.hpbNew + metrics.nextRes.mininaiNew) / allNewCustomers)*100).toFixed(1) : 0;
    animateCount('kpi-new-reservation-rate', parseFloat(newResRate), { suffix: '%', decimals: 1 });
    // Also update the large display in KPI tab
    const kpiRateHpbLg = document.getElementById('kpi-rate-hpb-lg');
    if (kpiRateHpbLg) kpiRateHpbLg.innerText = `${newResRate}%`;

    // Existing Customer Reservation Rate
    const existingResRate = metrics.customersExisting > 0 ? ((metrics.nextRes.existing / metrics.customersExisting)*100).toFixed(1) : 0;
    // Update the large display in KPI tab
    const kpiRateExistLg = document.getElementById('kpi-rate-exist-lg');
    if (kpiRateExistLg) kpiRateExistLg.innerText = `${existingResRate}%`;

    // Total Reservation Rate: Total next reservations / Total customers
    const totalResRate = metrics.customersTotal > 0 ? ((metrics.nextRes.total / metrics.customersTotal)*100).toFixed(1) : 0;
    animateCount('kpi-total-reservation-rate', parseFloat(totalResRate), { suffix: '%', decimals: 1 });
    // Also update the large display in KPI tab
    const kpiRateTotalLg = document.getElementById('kpi-rate-total-lg');
    if (kpiRateTotalLg) kpiRateTotalLg.innerText = `${totalResRate}%`;

    // Staff Dashboard Tab Logic
    const staffSel = document.getElementById('staff-selector');
    const staffDashboardTab = document.getElementById('tab-staff-dashboard');
    const counselingResultsTab = document.getElementById('tab-counseling-results');

    try {
        const sidebarStaff = document.getElementById('sidebar-staff-dashboard');
        const sidebarCounseling = document.getElementById('sidebar-counseling-results');
        const sheetStaff = document.getElementById('sheet-staff-dashboard');
        const sheetCounseling = document.getElementById('sheet-counseling-results');
        if (staffSel.value !== 'all') {
            staffDashboardTab.classList.remove('hidden');
            if (counselingResultsTab) counselingResultsTab.classList.remove('hidden');
            if (sidebarStaff) sidebarStaff.classList.remove('hidden');
            if (sidebarCounseling) sidebarCounseling.classList.remove('hidden');
            if (sheetStaff) sheetStaff.classList.remove('hidden');
            if (sheetCounseling) sheetCounseling.classList.remove('hidden');

            // Calculate incentive data for staff dashboard
            const storeVal = document.getElementById('store-selector').value;
            const incentiveData = calculateIncentive(filtered, storeVal, staffSel.value);

            // Update Staff Dashboard
            updateStaffDashboard(staffSel.value, metrics, incentiveData);
        } else {
            staffDashboardTab.classList.add('hidden');
            if (counselingResultsTab) counselingResultsTab.classList.add('hidden');
            if (sidebarStaff) sidebarStaff.classList.add('hidden');
            if (sidebarCounseling) sidebarCounseling.classList.add('hidden');
            if (sheetStaff) sheetStaff.classList.add('hidden');
            if (sheetCounseling) sheetCounseling.classList.add('hidden');
        }
    } catch (e) { console.error('staffDashboard エラー:', e); }

    try { updateTable(filtered); } catch (e) { console.error('updateTable エラー:', e); }
    try { updateCharts(metrics); } catch (e) { console.error('updateCharts エラー:', e); }
    try { renderEditTable(filtered); } catch (e) { console.error('renderEditTable エラー:', e); }

    // Update Staff Summary in Overview
    try { updateStaffSummary(); } catch (e) { console.error('updateStaffSummary エラー:', e); }

    // Update Staff Incentive Summary (Admin Only)
    try { updateStaffIncentiveSummary(); } catch (e) { console.error('updateStaffIncentiveSummary エラー:', e); }

    // Update Incentive Tab (if visible)
    try {
        const incentiveSection = document.getElementById('content-incentive');
        if (incentiveSection && !incentiveSection.classList.contains('hidden')) {
            updateIncentiveTab();
        }
    } catch (e) { console.error('updateIncentiveTab エラー:', e); }

    // Update Blog Progress
    try { updateBlogProgress(); } catch (e) { console.error('updateBlogProgress エラー:', e); }

    // Today's Snapshot + Sparklines + Delta Badges (Overview のみ)
    try { updateTodaySnapshot(currentGoal); } catch (e) { console.error('updateTodaySnapshot エラー:', e); }
    try { updateOverviewExtras(metrics, currentGoal); } catch (e) { console.error('updateOverviewExtras エラー:', e); }

    // 初期ローディング（スケルトン）表示を解除
    if (document.body.classList.contains('is-loading')) {
        document.body.classList.remove('is-loading');
    }

    // 新しく追加したアイコンを Lucide に再描画させる
    try { if (window.lucide) lucide.createIcons(); } catch (e) {}
}

function updateStaffDashboard(staffName, personalMetrics, incentiveData) {
    const fmt = n => n.toLocaleString();

    // Update staff name (if element exists)
    const staffNameEl = document.getElementById('staff-dashboard-name');
    if (staffNameEl) staffNameEl.innerText = staffName;

    // Get store total (filter by current store AND period)
    const storeFilter = document.getElementById('store-selector')?.value || 'all';
    const dateSelector = document.getElementById('date-selector');
    const selectedDate = dateSelector ? dateSelector.value : null;
    const safeRawData = Array.isArray(rawData) ? rawData : [];

    // 店舗データも同じ期間でフィルタリング
    const storeData = safeRawData.filter(d => {
        if (!d || !d.date) return false;
        if (storeFilter !== 'all' && d.store !== storeFilter) return false;

        // 期間フィルターの適用
        const recordDate = parseDate(d.date);
        const today = new Date();
        let startDate;

        switch (currentPeriodFilter) {
            case 'month':
                // selectedDateがない場合は現在の月を使用
                const dateStr = selectedDate || `${today.getFullYear()}/${today.getMonth() + 1}`;
                const [year, month] = dateStr.split('/').map(Number);
                return recordDate.getFullYear() === year && (recordDate.getMonth() + 1) === month;
            case '3months':
                startDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);
                return recordDate >= startDate;
            case '6months':
                startDate = new Date(today.getFullYear(), today.getMonth() - 5, 1);
                return recordDate >= startDate;
            case 'year':
                startDate = new Date(today.getFullYear() - 1, today.getMonth(), 1);
                return recordDate >= startDate;
            default:
                return true;
        }
    });
    const storeMetrics = calculateMetrics(storeData);

    // Get staff goals from storage
    const staffGoalData = getStaffGoal(storeFilter, staffName);
    const salesGoal = staffGoalData.weekdays * staffGoalData.weekdayTarget + staffGoalData.weekends * staffGoalData.weekendTarget;
    const newGoal = staffGoalData.newCustomers;
    const existingGoal = staffGoalData.existingCustomers;
    const unitPriceGoal = staffGoalData.unitPrice;
    const newResRateGoal = staffGoalData.newReservationRate;
    const resRateGoal = staffGoalData.reservationRate;

    // Store goals (aggregate from all staff in store)
    const storeGoalData = getStoreAggregateGoal(storeFilter);
    const storeSalesGoal = storeGoalData.weekdays * storeGoalData.weekdayTarget + storeGoalData.weekends * storeGoalData.weekendTarget;
    const storeNewGoal = storeGoalData.newCustomers;
    const storeExistingGoal = storeGoalData.existingCustomers;
    const storeUnitPriceGoal = storeGoalData.unitPrice;
    const storeNewResRateGoal = storeGoalData.newReservationRate;
    const storeResRateGoal = storeGoalData.reservationRate;

    // Personal metrics
    const personalUnitPrice = personalMetrics.customersTotal > 0 ? Math.round(personalMetrics.salesTotal / personalMetrics.customersTotal) : 0;
    const personalNewResRate = personalMetrics.hpbNewCount > 0 ? (((personalMetrics.nextRes.hpbNew + personalMetrics.nextRes.mininaiNew) / personalMetrics.hpbNewCount)*100).toFixed(1) : 0;
    const personalResRate = personalMetrics.customersTotal > 0 ? ((personalMetrics.nextRes.total / personalMetrics.customersTotal)*100).toFixed(1) : 0;
    const personalSalesPercent = salesGoal > 0 ? Math.round((personalMetrics.salesTotal / salesGoal) * 100) : 0;

    // Store metrics
    const storeUnitPrice = storeMetrics.customersTotal > 0 ? Math.round(storeMetrics.salesTotal / storeMetrics.customersTotal) : 0;
    const storeNewResRate = storeMetrics.hpbNewCount > 0 ? (((storeMetrics.nextRes.hpbNew + storeMetrics.nextRes.mininaiNew) / storeMetrics.hpbNewCount)*100).toFixed(1) : 0;
    const storeResRate = storeMetrics.customersTotal > 0 ? ((storeMetrics.nextRes.total / storeMetrics.customersTotal)*100).toFixed(1) : 0;
    const storeSalesPercent = storeSalesGoal > 0 ? Math.round((storeMetrics.salesTotal / storeSalesGoal) * 100) : 0;

    // Update Personal Section
    document.getElementById('staff-personal-sales').innerText = `¥${fmt(personalMetrics.salesTotal)}`;
    document.getElementById('staff-personal-goal').innerText = `¥${fmt(salesGoal)}`;
    document.getElementById('staff-personal-goal-rate').innerText = `${personalSalesPercent}%`;
    document.getElementById('staff-personal-customers').innerText = `${fmt(personalMetrics.customersTotal)}名`;
    document.getElementById('staff-personal-new').innerText = fmt(personalMetrics.customersNew);
    document.getElementById('staff-personal-new-goal').innerText = newGoal;
    document.getElementById('staff-personal-existing').innerText = fmt(personalMetrics.customersExisting);
    document.getElementById('staff-personal-existing-goal').innerText = existingGoal;
    document.getElementById('staff-personal-unit-price').innerText = `¥${fmt(personalUnitPrice)}`;
    document.getElementById('staff-personal-unit-price-goal').innerText = `¥${fmt(unitPriceGoal)}`;
    document.getElementById('staff-personal-new-res-rate').innerText = `${personalNewResRate}%`;
    document.getElementById('staff-personal-new-res-rate-goal').innerText = `${newResRateGoal}%`;
    document.getElementById('staff-personal-res-rate').innerText = `${personalResRate}%`;
    document.getElementById('staff-personal-res-rate-goal').innerText = `${resRateGoal}%`;
    document.getElementById('staff-personal-reviews').innerText = fmt(personalMetrics.reviews5StarTotal);
    document.getElementById('staff-personal-reviews-goal').innerText = staffGoalData.reviews5Star || 0;

    // Update Store Section
    document.getElementById('staff-store-sales').innerText = `¥${fmt(storeMetrics.salesTotal)}`;
    document.getElementById('staff-store-goal').innerText = `¥${fmt(storeSalesGoal)}`;
    document.getElementById('staff-store-goal-rate').innerText = `${storeSalesPercent}%`;
    document.getElementById('staff-store-customers').innerText = `${fmt(storeMetrics.customersTotal)}名`;
    document.getElementById('staff-store-new').innerText = fmt(storeMetrics.customersNew);
    document.getElementById('staff-store-new-goal').innerText = storeNewGoal;
    document.getElementById('staff-store-existing').innerText = fmt(storeMetrics.customersExisting);
    document.getElementById('staff-store-existing-goal').innerText = storeExistingGoal;
    document.getElementById('staff-store-unit-price').innerText = `¥${fmt(storeUnitPrice)}`;
    document.getElementById('staff-store-unit-price-goal').innerText = `¥${fmt(storeUnitPriceGoal)}`;
    document.getElementById('staff-store-new-res-rate').innerText = `${storeNewResRate}%`;
    document.getElementById('staff-store-new-res-rate-goal').innerText = `${storeNewResRateGoal}%`;
    document.getElementById('staff-store-res-rate').innerText = `${storeResRate}%`;
    document.getElementById('staff-store-res-rate-goal').innerText = `${storeResRateGoal}%`;
    document.getElementById('staff-store-reviews').innerText = fmt(storeMetrics.reviews5StarTotal);
    document.getElementById('staff-store-reviews-goal').innerText = storeGoalData.reviews5Star || 0;

    // Update Progress Bars
    const newPercent = newGoal > 0 ? Math.min(Math.round((personalMetrics.customersNew / newGoal) * 100), 100) : 0;
    const existingPercent = existingGoal > 0 ? Math.min(Math.round((personalMetrics.customersExisting / existingGoal) * 100), 100) : 0;

    document.getElementById('staff-progress-sales-percent').innerText = personalSalesPercent;
    document.getElementById('staff-progress-sales-bar').style.width = `${Math.min(personalSalesPercent, 100)}%`;
    document.getElementById('staff-progress-new-percent').innerText = newPercent;
    document.getElementById('staff-progress-new-bar').style.width = `${newPercent}%`;
    document.getElementById('staff-progress-existing-percent').innerText = existingPercent;
    document.getElementById('staff-progress-existing-bar').style.width = `${existingPercent}%`;

    // Update Incentive Details
    if (incentiveData) {
        document.getElementById('staff-incentive-total').innerText = `¥${Math.round(incentiveData.totalIncentive).toLocaleString()}`;
        document.getElementById('staff-incentive-base').innerText = `¥${incentiveData.baseSalary.toLocaleString()}`;
        document.getElementById('staff-incentive-service-sales').innerText = `施術売上(税抜): ¥${Math.round(incentiveData.serviceSalesTaxExcl).toLocaleString()}`;
        document.getElementById('staff-incentive-service-rate').innerText = `×40% − ¥${incentiveData.baseSalary.toLocaleString()}`;
        document.getElementById('staff-incentive-service').innerText = `¥${Math.round(incentiveData.serviceIncentive).toLocaleString()}`;
        document.getElementById('staff-incentive-retail-sales').innerText = `物販売上(税抜): ¥${Math.round(incentiveData.retailSalesTaxExcl).toLocaleString()}`;
        document.getElementById('staff-incentive-retail-rate').innerText = '(10%適用)';
        document.getElementById('staff-incentive-retail').innerText = `¥${Math.round(incentiveData.retailIncentive).toLocaleString()}`;
    }

    // SNS・ブログ更新数
    const snsEl = document.getElementById('staff-personal-sns');
    const blogEl = document.getElementById('staff-personal-blog');
    if (snsEl) snsEl.innerText = fmt(personalMetrics.snsUpdatesTotal || 0);
    if (blogEl) blogEl.innerText = fmt(personalMetrics.blogUpdatesTotal || 0);

    // 店舗内ランキング（上位3名）
    try {
        const rankingStaff = STAFF_ROSTER[storeFilter] || [];
        const staffRankData = rankingStaff.map(name => {
            const nameLower = name.toLowerCase();
            const sData = storeData.filter(d => d.staff?.toLowerCase() === nameLower);
            const m = calculateMetrics(sData);
            const up = m.customersTotal > 0 ? Math.round(m.salesTotal / m.customersTotal) : 0;
            return { name, sales: m.salesTotal, unitPrice: up, customers: m.customersTotal };
        });

        // 売上ランキング
        const salesRank = [...staffRankData].sort((a, b) => b.sales - a.sales).slice(0, 3);
        const salesRankEl = document.getElementById('staff-ranking-sales');
        if (salesRankEl) {
            salesRankEl.innerHTML = salesRank.map((s, i) => {
                const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : '🥉';
                const isMe = s.name.toLowerCase() === staffName.toLowerCase();
                return `<div class="flex items-center justify-between p-2 rounded ${isMe ? 'bg-mavie-100 border border-mavie-300' : 'bg-surface-50'}">
                    <div class="flex items-center gap-2">
                        <span class="text-sm">${medal}</span>
                        <span class="text-xs font-bold ${isMe ? 'text-mavie-800' : 'text-mavie-600'}">${s.name}</span>
                    </div>
                    <span class="text-xs font-bold text-mavie-800">¥${fmt(s.sales)}</span>
                </div>`;
            }).join('');
        }

        // 客単価ランキング（来店ありのみ）
        const upRank = [...staffRankData].filter(s => s.customers > 0).sort((a, b) => b.unitPrice - a.unitPrice).slice(0, 3);
        const upRankEl = document.getElementById('staff-ranking-unit-price');
        if (upRankEl) {
            upRankEl.innerHTML = upRank.length > 0 ? upRank.map((s, i) => {
                const medal = i === 0 ? '🥇' : i === 1 ? '🥈' : '🥉';
                const isMe = s.name.toLowerCase() === staffName.toLowerCase();
                return `<div class="flex items-center justify-between p-2 rounded ${isMe ? 'bg-mavie-100 border border-mavie-300' : 'bg-surface-50'}">
                    <div class="flex items-center gap-2">
                        <span class="text-sm">${medal}</span>
                        <span class="text-xs font-bold ${isMe ? 'text-mavie-800' : 'text-mavie-600'}">${s.name}</span>
                    </div>
                    <span class="text-xs font-bold text-mavie-800">¥${fmt(s.unitPrice)}</span>
                </div>`;
            }).join('') : '<p class="text-xs text-mavie-400">データなし</p>';
        }
    } catch (e) { console.error('ランキング更新エラー:', e); }

    // Load counseling data for staff dashboard
    loadCounselingForStaffDashboard();
}

// --- STAFF SUMMARY FUNCTIONS ---
let staffSummaryView = 'cards';
let staffSummarySortKey = 'sales';
let staffSummarySortAsc = false;

function toggleStaffSummaryView(view) {
    staffSummaryView = view;
    const cardsView = document.getElementById('staff-cards-view');
    const tableView = document.getElementById('staff-table-view');
    const cardsBtn = document.getElementById('staff-view-cards');
    const tableBtn = document.getElementById('staff-view-table');

    if (view === 'cards') {
        cardsView.classList.remove('hidden');
        tableView.classList.add('hidden');
        cardsBtn.classList.add('bg-primary-500', 'text-white');
        cardsBtn.classList.remove('text-surface-600', 'hover:bg-surface-100');
        tableBtn.classList.remove('bg-primary-500', 'text-white');
        tableBtn.classList.add('text-surface-600', 'hover:bg-surface-100');
    } else {
        cardsView.classList.add('hidden');
        tableView.classList.remove('hidden');
        tableBtn.classList.add('bg-primary-500', 'text-white');
        tableBtn.classList.remove('text-surface-600', 'hover:bg-surface-100');
        cardsBtn.classList.remove('bg-primary-500', 'text-white');
        cardsBtn.classList.add('text-surface-600', 'hover:bg-surface-100');
    }
}

function getStaffMetrics() {
    // rawDataが配列でない場合は空配列として扱う
    const safeRawData = Array.isArray(rawData) ? rawData : [];

    const storeFilter = document.getElementById('store-selector')?.value || 'all';
    const dateFilter = document.getElementById('date-selector')?.value || '';
    const dateParts = dateFilter.split('/');
    const year = dateParts[0] ? Number(dateParts[0]) : new Date().getFullYear();
    const month = dateParts[1] ? Number(dateParts[1]) : new Date().getMonth() + 1;

    // Filter data by store and date
    let filtered = safeRawData.filter(d => {
        if (!d || !d.date) return false;
        const dDate = parseDate(d.date);
        const matchesDate = dDate.getFullYear() === year && (dDate.getMonth() + 1) === month;
        const matchesStore = storeFilter === 'all' || d.store === storeFilter;
        return matchesDate && matchesStore;
    });

    // If period filter is not 'month', expand the date range
    // スタッフフィルタを無視して、店舗と期間のみでフィルタリング
    if (currentPeriodFilter !== 'month') {
        filtered = safeRawData.filter(d => {
            if (!d || !d.date) return false;
            const matchesStore = storeFilter === 'all' || d.store === storeFilter;
            if (!matchesStore) return false;

            const recordDate = parseDate(d.date);
            const today = new Date();
            let startDate;

            switch (currentPeriodFilter) {
                case '3months':
                    startDate = new Date(today.getFullYear(), today.getMonth() - 2, 1);
                    return recordDate >= startDate;
                case '6months':
                    startDate = new Date(today.getFullYear(), today.getMonth() - 5, 1);
                    return recordDate >= startDate;
                case 'year':
                    startDate = new Date(today.getFullYear() - 1, today.getMonth(), 1);
                    return recordDate >= startDate;
            }
            return true;
        });
    }

    // Get all staff for the selected store(s)
    let staffList = [];
    if (storeFilter === 'all') {
        Object.values(STAFF_ROSTER).forEach(s => staffList.push(...s));
    } else {
        staffList = STAFF_ROSTER[storeFilter] || [];
    }

    // Calculate metrics for each staff
    const staffMetrics = staffList.map(staff => {
        // スタッフ名の比較は大文字小文字を区別しない
        const staffLower = staff.toLowerCase();
        const staffData = filtered.filter(d => d.staff?.toLowerCase() === staffLower);
        const metrics = calculateMetrics(staffData);

        // Calculate previous period for trend comparison
        let prevFiltered = [];
        if (currentPeriodFilter === 'month') {
            const prevMonth = month === 1 ? 12 : month - 1;
            const prevYear = month === 1 ? year - 1 : year;
            prevFiltered = safeRawData.filter(d => {
                if (!d || !d.date) return false;
                const dDate = parseDate(d.date);
                const matchesDate = dDate.getFullYear() === prevYear && (dDate.getMonth() + 1) === prevMonth;
                const matchesStore = storeFilter === 'all' || d.store === storeFilter;
                return matchesDate && matchesStore && d.staff?.toLowerCase() === staffLower;
            });
        }
        const prevMetrics = calculateMetrics(prevFiltered);

        // Calculate derived metrics
        const unitPrice = metrics.customersTotal > 0 ? Math.round(metrics.salesTotal / metrics.customersTotal) : 0;
        const nextResRate = metrics.customersTotal > 0 ? Math.round((metrics.nextRes.total / metrics.customersTotal) * 100) : 0;
        const salesTrend = prevMetrics.salesTotal > 0 ? Math.round(((metrics.salesTotal - prevMetrics.salesTotal) / prevMetrics.salesTotal) * 100) : 0;

        // Get staff goal（期間フィルターに応じて調整）
        const staffStore = storeFilter === 'all' ? (safeRawData.find(d => d && d.staff?.toLowerCase() === staffLower)?.store || getStaffStoreFromRoster(staff) || 'chiba') : storeFilter;
        const goalData = getStaffGoal(staffStore, staff);
        const monthlyGoal = goalData.weekdays * goalData.weekdayTarget + goalData.weekends * goalData.weekendTarget;

        // 期間フィルターに応じて目標を倍数化
        let goalMultiplier = 1;
        switch (currentPeriodFilter) {
            case '3months': goalMultiplier = 3; break;
            case '6months': goalMultiplier = 6; break;
            case 'year': goalMultiplier = 12; break;
        }
        const salesGoal = monthlyGoal * goalMultiplier;
        const goalRate = salesGoal > 0 ? Math.round((metrics.salesTotal / salesGoal) * 100) : 0;

        return {
            name: staff,
            store: staffStore,
            sales: metrics.salesTotal,
            customers: metrics.customersTotal,
            newCustomers: metrics.customersNew,
            existingCustomers: metrics.customersExisting,
            unitPrice,
            nextResRate,
            salesTrend,
            goalRate,
            salesGoal,
            hpbNew: metrics.newByChannel?.hpb || 0,
            miniNaiNew: metrics.newByChannel?.mininai || 0
        };
    });

    // Sort by sales (default)
    return staffMetrics.sort((a, b) => b.sales - a.sales);
}

function updateStaffSummary() {
    let staffMetrics = getStaffMetrics();
    const fmt = n => n.toLocaleString();

    // スタッフ専用URLの場合は上位3名のみ表示
    if (lockedStaff) {
        staffMetrics = staffMetrics.slice(0, 3);
    }

    // Update Cards View
    const cardsContainer = document.getElementById('staff-cards-view');
    if (staffMetrics.length === 0) {
        cardsContainer.innerHTML = renderEmptyState({
            icon: 'users',
            title: 'スタッフデータがありません',
            desc: '選択中の期間・条件ではスタッフデータがありません。期間を変えて再確認してください。',
            colSpan: true,
        });
    } else {
        cardsContainer.innerHTML = staffMetrics.map((s, i) => {
            const c = getStaffColor(s.name);
            const rankBadge = i < 3
                ? `<span class="absolute -top-1 -right-1 w-5 h-5 rounded-full text-[10px] font-bold text-white flex items-center justify-center shadow" style="background:${i === 0 ? '#d4af37' : i === 1 ? '#a0a0a0' : '#b87333'};">${i + 1}</span>`
                : '';
            return `
            <div class="bg-surface-50 dark:bg-accent-900/50 rounded-xl p-4 hover:shadow-md transition-all duration-200 border border-surface-200 dark:border-accent-700 flex gap-3 relative overflow-hidden">
                <div class="staff-accent-bar" style="background:linear-gradient(180deg, ${c.main}, ${c.dark});"></div>
                <div class="flex-1 min-w-0">
                <div class="flex items-center justify-between mb-3">
                    <div class="flex items-center gap-3">
                        <div class="relative">
                            ${renderStaffAvatar(s.name, 36)}
                            ${rankBadge}
                        </div>
                        <div>
                            <h4 class="font-semibold text-accent-800 dark:text-surface-100">${s.name}</h4>
                            <p class="text-xs text-surface-500">${s.store === 'chiba' ? '千葉店' : s.store === 'honatsugi' ? '本厚木店' : s.store === 'yamato' ? '大和店' : s.store}</p>
                        </div>
                    </div>
                    <div class="text-right">
                        <span class="text-xs px-2 py-1 rounded-full ${s.salesTrend >= 0 ? 'bg-emerald-100 text-emerald-700' : 'bg-red-100 text-red-700'}">
                            ${s.salesTrend >= 0 ? '↑' : '↓'} ${Math.abs(s.salesTrend)}%
                        </span>
                    </div>
                </div>
                <div class="grid grid-cols-2 gap-3">
                    <div class="bg-white dark:bg-accent-800 rounded-lg p-2">
                        <p class="text-xs text-surface-500 mb-1">売上</p>
                        <p class="text-lg font-bold text-accent-800 dark:text-surface-100">¥${fmt(s.sales)}</p>
                        <div class="mt-1 flex items-center gap-1">
                            <div class="flex-1 bg-surface-200 dark:bg-accent-700 h-1.5 rounded-full overflow-hidden">
                                <div class="h-full bg-primary-500 rounded-full" style="width: ${Math.min(s.goalRate, 100)}%"></div>
                            </div>
                            <span class="text-xs text-surface-500">${s.goalRate}%</span>
                        </div>
                    </div>
                    <div class="bg-white dark:bg-accent-800 rounded-lg p-2">
                        <p class="text-xs text-surface-500 mb-1">来店数</p>
                        <p class="text-lg font-bold text-accent-800 dark:text-surface-100">${s.customers}<span class="text-xs font-normal ml-1">名</span></p>
                        <p class="text-xs text-surface-500 mt-1">新規 ${s.newCustomers} / 既存 ${s.existingCustomers}</p>
                    </div>
                    <div class="bg-white dark:bg-accent-800 rounded-lg p-2">
                        <p class="text-xs text-surface-500 mb-1">客単価</p>
                        <p class="text-lg font-bold text-accent-800 dark:text-surface-100">¥${fmt(s.unitPrice)}</p>
                    </div>
                    <div class="bg-white dark:bg-accent-800 rounded-lg p-2">
                        <p class="text-xs text-surface-500 mb-1">次回予約率</p>
                        <p class="text-lg font-bold text-accent-800 dark:text-surface-100">${s.nextResRate}%</p>
                    </div>
                </div>
                </div>
            </div>
        `;}).join('');
    }

    // Update Table View
    const tableBody = document.getElementById('staff-table-body');
    if (staffMetrics.length === 0) {
        tableBody.innerHTML = `<tr><td colspan="8">${renderEmptyState({ icon: 'users', title: 'スタッフデータがありません', desc: '期間を変えて再度確認してください。' })}</td></tr>`;
    } else {
    tableBody.innerHTML = staffMetrics.map((s, i) => `
        <tr class="border-b border-surface-100 hover:bg-surface-50 dark:hover:bg-accent-800/50 transition">
            <td class="py-3 px-4">
                <span class="inline-flex items-center justify-center w-6 h-6 rounded-full ${i === 0 ? 'bg-amber-500 text-white' : i === 1 ? 'bg-gray-400 text-white' : i === 2 ? 'bg-amber-700 text-white' : 'bg-surface-200 text-surface-600'} text-xs font-bold">
                    ${i + 1}
                </span>
            </td>
            <td class="py-3 px-4 font-medium text-accent-800 dark:text-surface-100">
                <span class="inline-flex items-center"><span class="staff-color-dot" style="background:${getStaffColor(s.name).main};"></span>${s.name}</span>
            </td>
            <td class="py-3 px-4 text-right font-semibold text-accent-800 dark:text-surface-100">¥${fmt(s.sales)}</td>
            <td class="py-3 px-4 text-right">${s.customers}</td>
            <td class="py-3 px-4 text-right text-primary-600 font-medium">${s.newCustomers}</td>
            <td class="py-3 px-4 text-right">¥${fmt(s.unitPrice)}</td>
            <td class="py-3 px-4 text-right">${s.nextResRate}%</td>
            <td class="py-3 px-4 text-center">
                <span class="inline-flex items-center gap-1 text-xs px-2 py-1 rounded-full ${s.salesTrend >= 0 ? 'bg-emerald-100 text-emerald-700' : 'bg-red-100 text-red-700'}">
                    ${s.salesTrend >= 0 ? '↑' : '↓'} ${Math.abs(s.salesTrend)}%
                </span>
            </td>
        </tr>
    `).join('');
    }

    // Update Rankings
    updateStaffRankings(staffMetrics);

    // Re-render Lucide icons
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function updateStaffRankings(staffMetrics) {
    const fmt = n => n.toLocaleString();
    // スタッフ専用URLの場合は上位3名、それ以外は上位5名
    const rankingLimit = lockedStaff ? 3 : 5;

    const rankBadge = i => `<span class="w-5 h-5 rounded-full ${i === 0 ? 'bg-amber-500' : i === 1 ? 'bg-gray-400' : i === 2 ? 'bg-amber-700' : 'bg-surface-300'} text-white text-xs flex items-center justify-center font-bold">${i + 1}</span>`;

    // Sales Ranking
    const salesRanking = [...staffMetrics].sort((a, b) => b.sales - a.sales).slice(0, rankingLimit);
    const maxSales = salesRanking[0]?.sales || 1;
    document.getElementById('sales-ranking-list').innerHTML = salesRanking.map((s, i) => {
        const c = getStaffColor(s.name);
        return `
        <div class="flex items-center gap-2">
            ${rankBadge(i)}
            <span class="inline-flex items-center flex-shrink-0 w-20 text-sm text-accent-700 dark:text-surface-200 truncate">
                <span class="staff-color-dot" style="background:${c.main};"></span>${s.name}
            </span>
            <div class="flex-1 bg-surface-200 dark:bg-accent-700 h-2 rounded-full overflow-hidden">
                <div class="h-full rounded-full" style="width: ${(s.sales / maxSales) * 100}%;background:linear-gradient(90deg, ${c.main}, ${c.dark});"></div>
            </div>
            <span class="text-xs text-accent-600 dark:text-surface-300 w-20 text-right">¥${fmt(s.sales)}</span>
        </div>`;
    }).join('');

    // New Customers Ranking
    const newRanking = [...staffMetrics].sort((a, b) => b.newCustomers - a.newCustomers).slice(0, rankingLimit);
    const maxNew = newRanking[0]?.newCustomers || 1;
    document.getElementById('new-customers-ranking-list').innerHTML = newRanking.map((s, i) => {
        const c = getStaffColor(s.name);
        return `
        <div class="flex items-center gap-2">
            ${rankBadge(i)}
            <span class="inline-flex items-center flex-shrink-0 w-20 text-sm text-accent-700 dark:text-surface-200 truncate">
                <span class="staff-color-dot" style="background:${c.main};"></span>${s.name}
            </span>
            <div class="flex-1 bg-surface-200 dark:bg-accent-700 h-2 rounded-full overflow-hidden">
                <div class="h-full rounded-full" style="width: ${(s.newCustomers / maxNew) * 100}%;background:linear-gradient(90deg, ${c.main}, ${c.dark});"></div>
            </div>
            <span class="text-xs text-accent-600 dark:text-surface-300 w-12 text-right">${s.newCustomers}名</span>
        </div>`;
    }).join('');

    // Unit Price Ranking
    const priceRanking = [...staffMetrics].filter(s => s.unitPrice > 0).sort((a, b) => b.unitPrice - a.unitPrice).slice(0, rankingLimit);
    const maxPrice = priceRanking[0]?.unitPrice || 1;
    document.getElementById('unit-price-ranking-list').innerHTML = priceRanking.map((s, i) => {
        const c = getStaffColor(s.name);
        return `
        <div class="flex items-center gap-2">
            ${rankBadge(i)}
            <span class="inline-flex items-center flex-shrink-0 w-20 text-sm text-accent-700 dark:text-surface-200 truncate">
                <span class="staff-color-dot" style="background:${c.main};"></span>${s.name}
            </span>
            <div class="flex-1 bg-surface-200 dark:bg-accent-700 h-2 rounded-full overflow-hidden">
                <div class="h-full rounded-full" style="width: ${(s.unitPrice / maxPrice) * 100}%;background:linear-gradient(90deg, ${c.main}, ${c.dark});"></div>
            </div>
            <span class="text-xs text-accent-600 dark:text-surface-300 w-16 text-right">¥${fmt(s.unitPrice)}</span>
        </div>`;
    }).join('');
}

// 管理者専用：スタッフ別インセンティブ一覧を更新
function updateStaffIncentiveSummary() {
    const fmt = n => n.toLocaleString();
    const incentiveSection = document.getElementById('staff-incentive-section');

    // スタッフ専用ページの場合は非表示
    if (lockedStore || lockedStaff) {
        if (incentiveSection) incentiveSection.style.display = 'none';
        return;
    }

    if (incentiveSection) incentiveSection.style.display = 'block';

    const tableBody = document.getElementById('staff-incentive-table-body');
    if (!tableBody) return;

    // Get current store/staff filter
    const storeFilter = document.getElementById('store-selector')?.value || 'all';
    const dateFilter = document.getElementById('date-selector')?.value || '';

    // Get all staff for all stores (or filtered store)
    const stores = storeFilter === 'all' ? Object.keys(STAFF_ROSTER) : [storeFilter];

    let allIncentives = [];
    let totals = {
        baseSalary: 0,
        serviceSalesTaxExcl: 0,
        serviceIncentive: 0,
        retailSalesTaxExcl: 0,
        retailIncentive: 0,
        totalIncentive: 0
    };

    // Calculate incentives for each staff
    stores.forEach(store => {
        const storeStaff = STAFF_ROSTER[store] || [];
        storeStaff.forEach(staffName => {
            // Filter data for this staff
            const staffData = (Array.isArray(rawData) ? rawData : []).filter(d => {
                if (!d) return false;
                // Date filter
                if (dateFilter && d.date) {
                    const dateParts = d.date.split('/');
                    if (dateParts.length >= 2) {
                        const dataYM = `${dateParts[0]}/${parseInt(dateParts[1])}`;
                        if (dataYM !== dateFilter) return false;
                    }
                }
                return d.store === store && d.staff === staffName;
            });

            // Calculate incentive
            const incentive = calculateIncentive(staffData, store, staffName);

            allIncentives.push({
                store,
                name: staffName,
                ...incentive
            });

            // Add to totals
            totals.baseSalary += incentive.baseSalary;
            totals.serviceSalesTaxExcl += incentive.serviceSalesTaxExcl;
            totals.serviceIncentive += incentive.serviceIncentive;
            totals.retailSalesTaxExcl += incentive.retailSalesTaxExcl;
            totals.retailIncentive += incentive.retailIncentive;
            totals.totalIncentive += incentive.totalIncentive;
        });
    });

    // Sort by total incentive descending
    allIncentives.sort((a, b) => b.totalIncentive - a.totalIncentive);

    // Generate table rows
    tableBody.innerHTML = allIncentives.map(s => {
        const storeName = s.store === 'chiba' ? '千葉店' : s.store === 'honatsugi' ? '本厚木店' : s.store === 'yamato' ? '大和店' : s.store;
        return `
            <tr class="hover:bg-surface-50 transition">
                <td class="py-3 px-4 text-surface-600">${storeName}</td>
                <td class="py-3 px-4 font-medium text-accent-800">${s.name}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(s.baseSalary)}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(Math.round(s.serviceSalesTaxExcl))}</td>
                <td class="py-3 px-4 text-right text-emerald-600">¥${fmt(Math.round(s.serviceIncentive))}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(Math.round(s.retailSalesTaxExcl))}</td>
                <td class="py-3 px-4 text-right text-emerald-600">¥${fmt(Math.round(s.retailIncentive))}</td>
                <td class="py-3 px-4 text-right font-bold text-[#b8956a] text-base">¥${fmt(Math.round(s.totalIncentive))}</td>
            </tr>
        `;
    }).join('');

    // Update totals
    document.getElementById('incentive-total-base').innerText = `¥${fmt(totals.baseSalary)}`;
    document.getElementById('incentive-total-service-sales').innerText = `¥${fmt(Math.round(totals.serviceSalesTaxExcl))}`;
    document.getElementById('incentive-total-service').innerText = `¥${fmt(Math.round(totals.serviceIncentive))}`;
    document.getElementById('incentive-total-retail-sales').innerText = `¥${fmt(Math.round(totals.retailSalesTaxExcl))}`;
    document.getElementById('incentive-total-retail').innerText = `¥${fmt(Math.round(totals.retailIncentive))}`;
    document.getElementById('incentive-grand-total').innerText = `¥${fmt(Math.round(totals.totalIncentive))}`;
}

// インセンティブタブのテーブルを更新
function updateIncentiveTab() {
    const fmt = n => n.toLocaleString();
    const storeFilter = document.getElementById('store-selector')?.value || 'all';
    const dateFilter = document.getElementById('date-selector')?.value || '';
    const stores = storeFilter === 'all' ? Object.keys(STAFF_ROSTER) : [storeFilter];

    let allData = [];
    let serviceTotals = { cashTaxExcl: 0, creditTaxExcl: 0, qrTaxExcl: 0, serviceSalesTaxExcl: 0, serviceIncentive: 0 };
    let retailTotals = { retailSalesTaxExcl: 0, retailIncentive: 0 };
    let combinedTotals = { baseSalary: 0, serviceIncentive: 0, retailIncentive: 0, totalIncentive: 0 };

    stores.forEach(store => {
        const storeStaff = STAFF_ROSTER[store] || [];
        storeStaff.forEach(staffName => {
            const staffData = (Array.isArray(rawData) ? rawData : []).filter(d => {
                if (!d) return false;
                if (dateFilter && d.date) {
                    const dateParts = d.date.split('/');
                    if (dateParts.length >= 2) {
                        const dataYM = `${dateParts[0]}/${parseInt(dateParts[1])}`;
                        if (dataYM !== dateFilter) return false;
                    }
                }
                return d.store === store && d.staff === staffName;
            });

            let cash = 0, credit = 0, qr = 0, product = 0;
            staffData.forEach(d => {
                const sales = d.sales || {};
                cash += sales.cash || 0;
                credit += sales.credit || 0;
                qr += sales.qr || 0;
                product += sales.product || 0;
            });

            // 税抜計算（消費税5%）
            const cashTaxExcl = cash / 1.05;
            const creditTaxExcl = credit / 1.05;
            const qrTaxExcl = qr / 1.05;
            const serviceSalesTaxExcl = cashTaxExcl + creditTaxExcl + qrTaxExcl;
            const retailSalesTaxExcl = product / 1.05;

            const baseSalary = getStaffBaseSalary(store, staffName);
            // 施術手当 = max(0, 施術売上(税抜) × 40% − 基本給)
            const serviceIncentive = Math.max(0, serviceSalesTaxExcl * 0.4 - baseSalary);
            // 物販手当 = 物販売上(税抜) × 10%
            const retailIncentive = retailSalesTaxExcl * 0.1;
            const totalIncentive = baseSalary + serviceIncentive + retailIncentive;
            const salesFortyPercent = serviceSalesTaxExcl * 0.4;

            allData.push({
                store, name: staffName,
                cashTaxExcl, creditTaxExcl, qrTaxExcl,
                serviceSalesTaxExcl, serviceIncentive, salesFortyPercent,
                retailSalesTaxExcl, retailIncentive,
                baseSalary, totalIncentive
            });

            serviceTotals.cashTaxExcl += cashTaxExcl;
            serviceTotals.creditTaxExcl += creditTaxExcl;
            serviceTotals.qrTaxExcl += qrTaxExcl;
            serviceTotals.serviceSalesTaxExcl += serviceSalesTaxExcl;
            serviceTotals.serviceIncentive += serviceIncentive;
            retailTotals.retailSalesTaxExcl += retailSalesTaxExcl;
            retailTotals.retailIncentive += retailIncentive;
            combinedTotals.baseSalary += baseSalary;
            combinedTotals.serviceIncentive += serviceIncentive;
            combinedTotals.retailIncentive += retailIncentive;
            combinedTotals.totalIncentive += totalIncentive;
        });
    });

    allData.sort((a, b) => b.totalIncentive - a.totalIncentive);

    const storeName = s => s === 'chiba' ? '千葉店' : s === 'honatsugi' ? '本厚木店' : s === 'yamato' ? '大和店' : s;

    // Service Incentive Table
    const serviceBody = document.getElementById('incentive-service-table-body');
    if (serviceBody) {
        serviceBody.innerHTML = allData.map(s => `
            <tr class="hover:bg-surface-50 transition">
                <td class="py-3 px-4 text-surface-600">${storeName(s.store)}</td>
                <td class="py-3 px-4 font-medium text-accent-800">${s.name}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(Math.round(s.cashTaxExcl))}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(Math.round(s.creditTaxExcl))}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(Math.round(s.qrTaxExcl))}</td>
                <td class="py-3 px-4 text-right font-medium text-accent-800">¥${fmt(Math.round(s.serviceSalesTaxExcl))}</td>
                <td class="py-3 px-4 text-right font-bold text-blue-600 text-base">¥${fmt(Math.round(s.serviceIncentive))}</td>
            </tr>
        `).join('');
    }
    document.getElementById('incentive-tab-total-cash').innerText = `¥${fmt(Math.round(serviceTotals.cashTaxExcl))}`;
    document.getElementById('incentive-tab-total-credit').innerText = `¥${fmt(Math.round(serviceTotals.creditTaxExcl))}`;
    document.getElementById('incentive-tab-total-qr').innerText = `¥${fmt(Math.round(serviceTotals.qrTaxExcl))}`;
    document.getElementById('incentive-tab-total-service-sales').innerText = `¥${fmt(Math.round(serviceTotals.serviceSalesTaxExcl))}`;
    document.getElementById('incentive-tab-total-service').innerText = `¥${fmt(Math.round(serviceTotals.serviceIncentive))}`;

    // Retail Incentive Table
    const retailBody = document.getElementById('incentive-retail-table-body');
    if (retailBody) {
        retailBody.innerHTML = allData.map(s => `
            <tr class="hover:bg-surface-50 transition">
                <td class="py-3 px-4 text-surface-600">${storeName(s.store)}</td>
                <td class="py-3 px-4 font-medium text-accent-800">${s.name}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(Math.round(s.retailSalesTaxExcl))}</td>
                <td class="py-3 px-4 text-right text-surface-600">10%</td>
                <td class="py-3 px-4 text-right font-bold text-emerald-600 text-base">¥${fmt(Math.round(s.retailIncentive))}</td>
            </tr>
        `).join('');
    }
    document.getElementById('incentive-tab-total-retail-sales').innerText = `¥${fmt(Math.round(retailTotals.retailSalesTaxExcl))}`;
    document.getElementById('incentive-tab-total-retail').innerText = `¥${fmt(Math.round(retailTotals.retailIncentive))}`;

    // Combined Summary Table
    const combinedBody = document.getElementById('incentive-combined-table-body');
    if (combinedBody) {
        combinedBody.innerHTML = allData.map(s => `
            <tr class="hover:bg-surface-50 transition">
                <td class="py-3 px-4 text-surface-600">${storeName(s.store)}</td>
                <td class="py-3 px-4 font-medium text-accent-800">${s.name}</td>
                <td class="py-3 px-4 text-right text-surface-600">¥${fmt(s.baseSalary)}</td>
                <td class="py-3 px-4 text-right text-blue-600">¥${fmt(Math.round(s.serviceIncentive))}</td>
                <td class="py-3 px-4 text-right text-emerald-600">¥${fmt(Math.round(s.retailIncentive))}</td>
                <td class="py-3 px-4 text-right font-bold text-[#b8956a] text-base">¥${fmt(Math.round(s.totalIncentive))}</td>
            </tr>
        `).join('');
    }
    document.getElementById('incentive-tab-combined-base').innerText = `¥${fmt(combinedTotals.baseSalary)}`;
    document.getElementById('incentive-tab-combined-service').innerText = `¥${fmt(Math.round(combinedTotals.serviceIncentive))}`;
    document.getElementById('incentive-tab-combined-retail').innerText = `¥${fmt(Math.round(combinedTotals.retailIncentive))}`;
    document.getElementById('incentive-tab-combined-total').innerText = `¥${fmt(Math.round(combinedTotals.totalIncentive))}`;

    // Re-render Lucide icons
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

// ブログ・SNS更新進捗を更新
function updateBlogProgress() {
    const container = document.getElementById('blog-progress-container');
    if (!container) return;

    const BLOG_TARGET = 10; // 月間目標10件
    const filtered = getFilteredData();

    // 選択中の月のデータを取得
    const dateSelector = document.getElementById('date-selector');
    const selectedDate = dateSelector?.value || '';
    const [targetYear, targetMonth] = selectedDate ? selectedDate.split('/').map(Number) : [new Date().getFullYear(), new Date().getMonth() + 1];

    // スタッフごとのブログ・SNS更新数を集計
    const staffProgress = {};
    filtered.forEach(record => {
        if (!record.date) return;
        const recordDate = parseDate(record.date);
        if (recordDate.getFullYear() === targetYear && (recordDate.getMonth() + 1) === targetMonth) {
            const staffName = record.staff;
            if (!staffProgress[staffName]) {
                staffProgress[staffName] = { blogUpdates: 0, snsUpdates: 0 };
            }
            staffProgress[staffName].blogUpdates += record.blogUpdates || 0;
            staffProgress[staffName].snsUpdates += record.snsUpdates || 0;
        }
    });

    const staffList = Object.entries(staffProgress).sort((a, b) => b[1].blogUpdates - a[1].blogUpdates);

    if (staffList.length === 0) {
        container.innerHTML = `
            <div class="text-center text-surface-500 py-8 col-span-full">
                <i data-lucide="file-x" class="w-8 h-8 mx-auto mb-2 opacity-50"></i>
                <p class="text-sm">ブログ更新データがありません</p>
            </div>
        `;
        if (typeof lucide !== 'undefined') lucide.createIcons();
        return;
    }

    container.innerHTML = staffList.map(([staffName, data]) => {
        const blogPercent = Math.min((data.blogUpdates / BLOG_TARGET) * 100, 100);
        const isComplete = data.blogUpdates >= BLOG_TARGET;
        const barColor = isComplete ? 'bg-emerald-500' : blogPercent >= 70 ? 'bg-amber-500' : 'bg-orange-500';

        return `
            <div class="bg-surface-50 rounded-xl p-4 border border-surface-200">
                <div class="flex items-center justify-between mb-3">
                    <span class="font-semibold text-accent-800">${staffName}</span>
                    ${isComplete ? '<span class="text-emerald-600 text-xs font-bold">🎉 達成!</span>' : ''}
                </div>
                <div class="space-y-2">
                    <div>
                        <div class="flex justify-between text-xs mb-1">
                            <span class="text-surface-600">ブログ</span>
                            <span class="font-bold ${isComplete ? 'text-emerald-600' : 'text-accent-800'}">${data.blogUpdates} / ${BLOG_TARGET}件</span>
                        </div>
                        <div class="w-full bg-surface-200 h-3 rounded-full overflow-hidden">
                            <div class="h-full ${barColor} rounded-full transition-all duration-500" style="width: ${blogPercent}%"></div>
                        </div>
                    </div>
                    <div class="flex justify-between text-xs text-surface-500 pt-1">
                        <span>SNS更新</span>
                        <span class="font-medium text-accent-700">${data.snsUpdates}件</span>
                    </div>
                </div>
            </div>
        `;
    }).join('');

    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function updateTable(data) {
    const tbody = document.getElementById('data-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';
    const dataArray = Array.isArray(data) ? data : [];
    const sorted = [...dataArray].sort((a,b) => parseDate(b.date) - parseDate(a.date)).slice(0, 20);

    if (sorted.length === 0) {
        const colCount = tbody.closest('table')?.querySelectorAll('thead th').length || 6;
        tbody.innerHTML = `<tr><td colspan="${colCount}">${renderEmptyState({
            icon: 'file-search',
            title: '該当するデータがありません',
            desc: '選択中の店舗・スタッフ・期間ではデータが見つかりません。条件を変えて再検索してください。'
        })}</td></tr>`;
        if (window.lucide) lucide.createIcons();
        return;
    }

    sorted.forEach(d => {
        if (!d) return;
        const sales = d.sales || {};
        const customers = d.customers || {};
        const nextRes = d.nextRes || {};
        const total = (sales.cash || 0) + (sales.credit || 0) + (sales.qr || 0);
        const allNewCust = (customers.newHPB || 0) + (customers.newMiniNai || 0);
        const newResRate = allNewCust > 0 ? Math.round((((nextRes.newHPB || 0) + (nextRes.newMiniNai || 0))/allNewCust)*100) : '-';
        const storeLabel = d.store === 'chiba' ? '千葉店' : d.store === 'honatsugi' ? '本厚木店' : d.store === 'yamato' ? '大和店' : d.store;
        const staffColor = d.staff ? getStaffColor(d.staff).main : 'transparent';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap text-xs font-bold text-mavie-800">${storeLabel}</td>
            <td class="px-6 py-4 whitespace-nowrap text-xs text-mavie-600">
                ${d.staff ? `<span class="inline-flex items-center"><span class="staff-color-dot" style="background:${staffColor};"></span>${d.staff}</span>` : ''}
            </td>
            <td class="px-6 py-4 whitespace-nowrap text-xs text-mavie-500">${d.date || ''}</td>
            <td class="px-6 py-4 whitespace-nowrap text-xs text-right font-medium text-mavie-800">¥${total.toLocaleString()}</td>
            <td class="px-6 py-4 whitespace-nowrap text-xs text-right text-mavie-600">${(customers.newHPB || 0) + (customers.newMiniNai || 0) + (customers.existing || 0)}</td>
            <td class="px-6 py-4 whitespace-nowrap text-xs text-right ${newResRate < 40 && newResRate !== '-' ? 'text-red-400' : ''}">${newResRate !== '-' ? newResRate+'%' : '-'}</td>
        `;
        tbody.appendChild(tr);
    });
}

// --- 4. CHARTS ---
function initCharts() {
    const common = { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { usePointStyle: true, font: {family: 'sans-serif'} } } } };
    
    // Overview Chart: Line (Unit Price) + Stacked Bar (New/Existing)
    charts.overview = new Chart(document.getElementById('overviewChart'), {
        type: 'bar', // Base type
        data: {},
        options: {
            ...common,
            scales: {
                x: { stacked: true, grid: { display: false } },
                y: { 
                    type: 'linear', position: 'left', 
                    title: { display: true, text: '単価 (¥)' },
                    grid: { display: false } 
                },
                y1: { 
                    type: 'linear', position: 'right', stacked: true,
                    title: { display: true, text: '来店数 (名)' },
                    grid: { borderDash: [4, 4], color: '#e8e6e1' } 
                }
            }
        }
    });

    charts.ratio = new Chart(document.getElementById('customerRatioChart'), { type: 'doughnut', data: {}, options: { ...common, cutout: '70%' } });
    charts.share = new Chart(document.getElementById('channelShareChart'), { type: 'pie', data: {}, options: common });
    charts.trend = new Chart(document.getElementById('channelTrendChart'), { type: 'bar', data: {}, options: { ...common, scales: { x:{stacked:true}, y:{stacked:true} } } });
    charts.payment = new Chart(document.getElementById('paymentChart'), { type: 'doughnut', data: {}, options: { ...common, cutout: '60%' } });
    charts.loss = new Chart(document.getElementById('lossChart'), { type: 'bar', indexAxis: 'y', data: {}, options: common });
}

function updateCharts(metrics) {
    // Debug log for HPB data
    console.log('📊 媒体データ:', { HPB: metrics.newByChannel?.hpb || 0, minimo: metrics.newByChannel?.mininai || 0 });

    const labels = Object.keys(metrics.daily).sort((a,b) => parseDate(a) - parseDate(b));
    const d = metrics.daily;

    // Calculate Daily Unit Price
    const unitPriceData = labels.map(k => {
        const totalCust = d[k].customers;
        return totalCust > 0 ? Math.round(d[k].sales / totalCust) : 0;
    });

    // Update Overview Chart
    charts.overview.data = { 
        labels: labels, 
        datasets: [
            {
                type: 'line',
                label: '平均客単価',
                data: unitPriceData,
                borderColor: BrandColors.brown,
                backgroundColor: BrandColors.brown,
                borderWidth: 2,
                tension: 0.3,
                yAxisID: 'y',
                pointRadius: 2
            },
            {
                type: 'bar',
                label: '新規来店',
                data: labels.map(k => d[k].new),
                backgroundColor: BrandColors.gold,
                yAxisID: 'y1',
                stack: 'visits'
            },
            {
                type: 'bar',
                label: '既存来店',
                data: labels.map(k => d[k].existing),
                backgroundColor: BrandColors.beige,
                yAxisID: 'y1',
                stack: 'visits'
            }
        ]
    }; 
    charts.overview.update();

    charts.ratio.data = { labels: ['新規', '既存'], datasets: [{ data: [metrics.customersNew, metrics.customersExisting], backgroundColor: [BrandColors.gold, BrandColors.brown], borderWidth:0 }] }; charts.ratio.update();
    // Update HPB share chart and display counts
    const hpbCount = metrics.newByChannel.hpb || 0;
    const mininaiCount = metrics.newByChannel.mininai || 0;
    document.getElementById('hpb-count').innerText = hpbCount;
    document.getElementById('mininai-count').innerText = mininaiCount;
    charts.share.data = { labels: ['HPB', 'minimo/Nailie'], datasets: [{ data: [hpbCount, mininaiCount], backgroundColor: [BrandColors.gold, BrandColors.beige], borderWidth:1, borderColor:'#fff' }] }; charts.share.update();
    charts.trend.data = { labels: labels, datasets: [
        { label: 'HPB', data: labels.map(k=>d[k].hpb), backgroundColor: BrandColors.gold },
        { label: 'minimo/Nailie', data: labels.map(k=>d[k].mininai), backgroundColor: BrandColors.beige }
    ]}; charts.trend.update();
    charts.payment.data = { labels: ['現金', 'クレカ', 'QR'], datasets: [{ data: [metrics.salesCash, metrics.salesCredit, metrics.salesQR], backgroundColor: [BrandColors.beige, BrandColors.gold, BrandColors.brown], borderWidth:0 }] }; charts.payment.update();
    charts.loss.data = { labels: ['損失'], datasets: [{ label: '割引・返金', data: [metrics.lossTotal], backgroundColor: BrandColors.brown, barThickness:40 }] }; charts.loss.update();
}

// --- 5. EDIT & SAVE LOGIC ---
// 日付文字列の安全なパース（"yyyy/M/d" 形式をブラウザ非依存で処理）
// Safari/Firefox は new Date("2026/2/1") で Invalid Date になるケースがあるため
function parseDate(dateStr) {
    if (!dateStr && dateStr !== 0) return new Date(NaN);
    // 数値の場合はエポックミリ秒としてそのまま渡す
    if (typeof dateStr === 'number') return new Date(dateStr);
    const str = String(dateStr);
    const parts = str.split('/');
    if (parts.length === 3) {
        return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    }
    // ISO形式 "2026-02-01" にも対応（時刻部分を含む場合はフォールバック）
    const parts2 = str.split('-');
    if (parts2.length === 3 && !str.includes('T')) {
        return new Date(parseInt(parts2[0]), parseInt(parts2[1]) - 1, parseInt(parts2[2]));
    }
    return new Date(dateStr);
}

// 数値の安全な取得（空白・undefined・NaN → 0）
function safeNum(val) {
    const n = parseInt(val);
    return isNaN(n) ? 0 : n;
}

function renderEditTable(data) {
    const tbody = document.getElementById('edit-table-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    // dataが配列でない場合は空配列として扱う
    const dataArray = Array.isArray(data) ? data : [];
    const sorted = [...dataArray].sort((a,b) => parseDate(b.date) - parseDate(a.date));

    const isStaffMode = !!lockedStaff;
    const isStoreMode = !!lockedStore && !lockedStaff;

    sorted.forEach(row => {
        if (!row) return;

        // オブジェクトのデフォルト値設定
        const sales = row.sales || {};
        const customers = row.customers || {};
        const nextRes = row.nextRes || {};
        const discounts = row.discounts || {};

        // 日付フォーマット: 年を省略して MM/DD 表記にする
        let displayDate = row.date || '';
        if (displayDate) {
            const parts = displayDate.split('/');
            if (parts.length === 3) {
                displayDate = String(parts[1]).padStart(2, '0') + '/' + String(parts[2]).padStart(2, '0');
            }
        }

        // 売上合計 = 現金 + クレカ + QR + HPBポイント + HPBギフト
        const salesTotalRow = safeNum(sales.cash) + safeNum(sales.credit) + safeNum(sales.qr) + safeNum(discounts.hpbPoints) + safeNum(discounts.hpbGift);

        const tr = document.createElement('tr');
        tr.dataset.id = row.id;
        if (changedRows.has(row.id)) tr.classList.add('changed-row');

        // モードに応じた店舗・スタッフ列の表示制御
        let storeStaffCols;
        if (isStaffMode) {
            // スタッフ専用：店舗・スタッフ両方非表示
            storeStaffCols = '';
        } else if (isStoreMode) {
            // 店舗専用：店舗非表示、スタッフ表示
            storeStaffCols = `
            <td class="px-4 py-2"><input type="text" class="edit-input edit-readonly" value="${row.staff || ''}" readonly></td>
        `;
        } else {
            // 通常：両方表示
            storeStaffCols = `
            <td class="px-4 py-2"><input type="text" class="edit-input edit-readonly" value="${row.storeName || ''}" readonly></td>
            <td class="px-4 py-2"><input type="text" class="edit-input edit-readonly" value="${row.staff || ''}" readonly></td>
        `;
        }

        tr.innerHTML = `
            <td class="px-4 py-2"><input type="text" class="edit-input edit-input-date edit-readonly" value="${displayDate}" readonly></td>
            <td class="px-4 py-2" style="background:#fef9ef;"><input type="text" class="edit-input edit-input-amount edit-readonly" value="¥${salesTotalRow.toLocaleString()}" readonly style="font-weight:bold;color:#b8956a;text-align:center;"></td>
            ${storeStaffCols}
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="sales.cash" value="${safeNum(sales.cash)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="sales.credit" value="${safeNum(sales.credit)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="sales.qr" value="${safeNum(sales.qr)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="sales.product" value="${safeNum(sales.product)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="discounts.hpbPoints" value="${safeNum(discounts.hpbPoints)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="discounts.hpbGift" value="${safeNum(discounts.hpbGift)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="discounts.other" value="${safeNum(discounts.other)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input edit-input-amount" data-field="discounts.refund" value="${safeNum(discounts.refund)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="customers.newHPB" value="${safeNum(customers.newHPB)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="customers.newMiniNai" value="${safeNum(customers.newMiniNai)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="customers.existing" value="${safeNum(customers.existing)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="customers.acquaintance" value="${safeNum(customers.acquaintance)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="nextRes.newHPB" value="${safeNum(nextRes.newHPB)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="nextRes.newMiniNai" value="${safeNum(nextRes.newMiniNai)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="nextRes.existing" value="${safeNum(nextRes.existing)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="reviews5Star" value="${safeNum(row.reviews5Star)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="blogUpdates" value="${safeNum(row.blogUpdates)}" onchange="markChanged(${row.id})"></td>
            <td class="px-4 py-2"><input type="number" class="edit-input" data-field="snsUpdates" value="${safeNum(row.snsUpdates)}" onchange="markChanged(${row.id})"></td>
        `;
        tbody.appendChild(tr);
    });
}

function markChanged(id) {
    changedRows.add(id);
    const row = document.querySelector(`tr[data-id='${id}']`);
    if(row) row.classList.add('changed-row');
}

function applyEditsLocally() {
    // rawDataが配列でない場合はエラーを防ぐ
    if (!Array.isArray(rawData)) {
        console.warn('applyEditsLocally: rawDataが配列ではありません');
        return;
    }

    const inputs = document.querySelectorAll('#edit-table-body input:not([readonly])');
    inputs.forEach(input => {
        const tr = input.closest('tr');
        if (!tr) return;
        const id = parseInt(tr.dataset.id);
        if (!changedRows.has(id)) return;

        const fieldPath = input.dataset.field?.split('.') || [];
        // 空白・NaN は 0 として扱う
        const val = safeNum(input.value);
        const item = rawData.find(d => d && d.id === id);
        if (item) {
            if (fieldPath.length === 2) {
                // ネストされたオブジェクトが存在しない場合は作成
                if (!item[fieldPath[0]]) {
                    item[fieldPath[0]] = {};
                }
                item[fieldPath[0]][fieldPath[1]] = val;
            } else if (fieldPath.length === 1) {
                // Handle top-level fields like reviews5Star
                item[fieldPath[0]] = val;
            }
        }
    });
    alert("表示を更新しました (未保存)");
    updateDashboard();
}

async function saveToSpreadsheet() {
    if (!API_URL) {
        alert("API URLが設定されていないため、保存できません。GASをデプロイしてください。");
        applyEditsLocally();
        return;
    }

    const btn = document.getElementById('btn-save-sheet');
    const originalText = btn.innerHTML;
    btn.innerHTML = "保存中...";
    btn.disabled = true;

    applyEditsLocally();
    const safeRawData = Array.isArray(rawData) ? rawData : [];
    const modifiedData = safeRawData.filter(d => d && changedRows.has(d.id));
    
    if (modifiedData.length === 0) {
        alert("変更されたデータがありません");
        btn.innerHTML = originalText;
        btn.disabled = false;
        return;
    }

    try {
        const response = await fetch(API_URL, {
            method: 'POST',
            headers: { "Content-Type": "text/plain" },
            body: JSON.stringify({
                action: "update",
                rows: modifiedData
            })
        });
        
        const result = await response.json();
        
        if (result.status === "success") {
            alert("スプレッドシートへの保存が完了しました！");
            changedRows.clear();
            renderEditTable(getFilteredData());
        } else {
            alert("保存エラー: " + result.message);
        }

    } catch (e) {
        console.error(e);
        alert("通信エラーが発生しました。コンソールを確認してください。");
    } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
    }
}

async function saveGoalsToSpreadsheet() {
    if (!API_URL) {
        alert("API URLが設定されていないため、スプレッドシートに保存できません。ローカルストレージに保存されます。");
        return;
    }

    const btn = document.getElementById('btn-save-goals');
    const originalText = btn.innerHTML;
    btn.innerHTML = "保存中...";
    btn.disabled = true;

    try {
        const goals = loadGoalsFromStorage();
        const salaries = loadBaseSalariesFromStorage();

        const response = await fetch(API_URL, {
            method: 'POST',
            headers: { "Content-Type": "text/plain" },
            body: JSON.stringify({
                action: "save_goals",
                goals: goals,
                salaries: salaries
            })
        });

        const result = await response.json();

        if (result.status === "success") {
            alert("目標データがスプレッドシートに保存されました！");
        } else {
            alert("保存エラー: " + result.message);
        }

    } catch (e) {
        console.error(e);
        alert("通信エラーが発生しました。コンソールを確認してください。");
    } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
    }
}

// スプレッドシートから目標を読み込む
async function loadGoalsFromSpreadsheet() {
    if (!API_URL) {
        alert("API URLが設定されていません。設定タブでAPI URLを入力してください。");
        return;
    }

    const btn = document.getElementById('btn-load-goals');
    const originalText = btn.innerHTML;
    btn.innerHTML = "読み込み中...";
    btn.disabled = true;

    try {
        const response = await fetch(API_URL + '?action=load_goals');
        const result = await response.json();

        if (result.status === "success") {
            // 目標データをローカルストレージに保存
            if (result.goals) {
                saveGoalsToStorage(result.goals);
            }
            // 基本給データを保存
            if (result.salaries) {
                saveBaseSalariesToStorage(result.salaries);
            }
            alert("目標データをスプレッドシートから読み込みました！");
            // UIを更新
            loadGoalInputs();
        } else {
            alert("読み込みエラー: " + (result.message || "不明なエラー"));
        }

    } catch (e) {
        console.error(e);
        alert("通信エラーが発生しました。コンソールを確認してください。");
    } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
    }
}

// --- 6. GOAL CALCULATOR ---

// 目標設定タブの月セレクタを初期化
function initGoalMonthSelector() {
    const monthSelector = document.getElementById('goal-month-selector');
    if (!monthSelector) return;

    // 既存の値をYYYY/M形式に正規化
    let currentValue = monthSelector.value;
    if (currentValue) {
        const parts = currentValue.split('/');
        if (parts.length === 2) {
            currentValue = `${parts[0]}/${parseInt(parts[1])}`;
        }
    }
    monthSelector.innerHTML = '';

    // 過去6ヶ月から未来12ヶ月までの選択肢を生成（YYYY/M形式で統一）
    const now = new Date();
    const options = [];

    for (let i = -6; i <= 12; i++) {
        const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
        const year = d.getFullYear();
        const month = d.getMonth() + 1;
        const value = `${year}/${month}`; // YYYY/M形式（date-selectorと同じ）
        const label = `${year}年${month}月`;
        options.push({ value, label, isCurrent: i === 0 });
    }

    options.forEach(opt => {
        const option = document.createElement('option');
        option.value = opt.value;
        option.text = opt.label + (opt.isCurrent ? ' (今月)' : '');
        monthSelector.appendChild(option);
    });

    // 現在選択されている月またはダッシュボードの月を選択
    const dateSelector = document.getElementById('date-selector');
    const dashboardMonth = dateSelector ? dateSelector.value : null;

    if (currentValue && Array.from(monthSelector.options).some(o => o.value === currentValue)) {
        monthSelector.value = currentValue;
    } else if (dashboardMonth && Array.from(monthSelector.options).some(o => o.value === dashboardMonth)) {
        monthSelector.value = dashboardMonth;
    } else {
        // デフォルト：今月（YYYY/M形式）
        const defaultMonth = `${now.getFullYear()}/${now.getMonth() + 1}`;
        monthSelector.value = defaultMonth;
    }
}

function handleGoalStoreChange() {
    const storeSel = document.getElementById('goal-store-selector');
    const staffSel = document.getElementById('goal-staff-selector');
    const selectedStore = storeSel.value;

    staffSel.innerHTML = '<option value="all">店舗合計</option>';
    if (selectedStore === 'all') {
        staffSel.disabled = true;
    } else {
        staffSel.disabled = false;
        if(STAFF_ROSTER[selectedStore]){
            STAFF_ROSTER[selectedStore].forEach(name => {
                const opt = document.createElement('option');
                opt.value = name;
                opt.text = name;
                staffSel.appendChild(opt);
            });
        }
    }
    loadGoalInputs();
}

function getCurrentGoalContextForGoalTab() {
    const storeVal = document.getElementById('goal-store-selector').value;
    const staffVal = document.getElementById('goal-staff-selector').value;

    if (staffVal !== 'all') {
        return { type: 'staff', store: storeVal, staff: staffVal };
    } else if (storeVal !== 'all') {
        return { type: 'store', store: storeVal };
    } else {
        return { type: 'all' };
    }
}

function loadGoalInputs() {
    const context = getCurrentGoalContextForGoalTab();
    const yearMonth = getCurrentGoalMonth();
    let goalData;

    if (context.type === 'staff') {
        goalData = getStaffGoal(context.store, context.staff, yearMonth);
    } else if (context.type === 'store') {
        goalData = getStoreAggregateGoal(context.store, yearMonth);
    } else {
        goalData = getAllStoresAggregateGoal(yearMonth);
    }

    const fmt = n => n.toLocaleString();
    const inputColumn1 = document.getElementById('goal-input-column-1');
    const inputColumn2 = document.getElementById('goal-input-column-2');
    const storeSummary = document.getElementById('goal-store-summary');
    const hintText = document.getElementById('goal-hint-text');

    if (context.type === 'staff') {
        // スタッフ個人表示: 入力フィールドを表示、サマリーを非表示
        inputColumn1.classList.remove('hidden');
        inputColumn2.classList.remove('hidden');
        storeSummary.classList.add('hidden');

        // Load values into inputs
        document.getElementById('goal-weekdays').value = goalData.weekdays;
        document.getElementById('goal-weekends').value = goalData.weekends;
        document.getElementById('goal-weekday-target').value = goalData.weekdayTarget;
        document.getElementById('goal-weekend-target').value = goalData.weekendTarget;
        document.getElementById('goal-retail').value = goalData.retail;
        document.getElementById('goal-new-customers').value = goalData.newCustomers;
        document.getElementById('goal-existing-customers').value = goalData.existingCustomers;
        document.getElementById('goal-unit-price').value = goalData.unitPrice;
        document.getElementById('goal-new-reservation-rate').value = goalData.newReservationRate;
        document.getElementById('goal-reservation-rate').value = goalData.reservationRate;
        document.getElementById('goal-reviews-5star').value = goalData.reviews5Star || 0;

        const monthlyTarget = (goalData.weekdays * goalData.weekdayTarget) + (goalData.weekends * goalData.weekendTarget);
        document.getElementById('goal-monthly-target').value = monthlyTarget;

        // Update display fields
        document.getElementById('goal-weekday-target-display').innerText = `¥${fmt(goalData.weekdayTarget)}`;
        document.getElementById('goal-weekend-target-display').innerText = `¥${fmt(goalData.weekendTarget)}`;

        // Analyze past performance and display
        const performance = analyzeWeekdayWeekendPerformance();
        document.getElementById('goal-past-weekday-avg').innerText = `¥${fmt(performance.weekdayAvg)}`;
        document.getElementById('goal-past-weekend-avg').innerText = `¥${fmt(performance.weekendAvg)}`;
        document.getElementById('goal-past-ratio').innerText = `${performance.ratio}倍`;

        // Load base salary
        const baseSalary = getStaffBaseSalary(context.store, context.staff);
        document.getElementById('goal-base-salary').value = baseSalary;

        // Enable inputs
        const inputs = document.querySelectorAll('#content-goal input');
        inputs.forEach(input => input.disabled = false);
        hintText.innerHTML = '各項目の目標を設定すると<strong>自動保存</strong>されます。スタッフダッシュボードに反映されます。店舗合計は所属スタッフの目標の<strong>合算値</strong>となります。';
    } else {
        // 店舗または全店舗表示: 入力フィールドを非表示、サマリーを表示
        inputColumn1.classList.add('hidden');
        inputColumn2.classList.add('hidden');
        storeSummary.classList.remove('hidden');

        // サマリーの表示内容を更新
        const monthlyTarget = goalData.monthlyTargetSum || 0;
        document.getElementById('summary-monthly-target').innerText = `¥${fmt(monthlyTarget)}`;
        document.getElementById('summary-retail').innerText = `¥${fmt(goalData.retail)}`;
        document.getElementById('summary-new-customers').innerText = `${goalData.newCustomers}名`;
        document.getElementById('summary-existing-customers').innerText = `${goalData.existingCustomers}名`;
        document.getElementById('summary-unit-price').innerText = `¥${fmt(goalData.unitPrice)}`;
        document.getElementById('summary-reservation-rate').innerText = `${goalData.reservationRate}%`;

        // タイトルと説明を更新
        const titleEl = document.getElementById('goal-summary-title');
        const descEl = document.getElementById('goal-summary-description');
        if (context.type === 'store') {
            const storeName = context.store === 'chiba' ? '千葉店' : context.store === 'honatsugi' ? '本厚木店' : context.store === 'yamato' ? '大和店' : context.store;
            titleEl.innerText = `${storeName} 目標サマリー`;
            descEl.innerText = 'この目標は所属スタッフの目標の合算値です。個別のスタッフを選択すると、個人目標を編集できます。';
        } else {
            titleEl.innerText = '全店舗 目標サマリー';
            descEl.innerText = 'この目標は全スタッフの目標の合算値です。個別のスタッフを選択すると、個人目標を編集できます。';
        }

        // スタッフ別内訳を表示
        updateGoalStaffBreakdown(context, yearMonth);

        if (context.type === 'store') {
            hintText.innerHTML = '<strong>店舗合計表示中</strong>: スタッフを選択すると個人目標を編集できます。';
        } else {
            hintText.innerHTML = '<strong>全店舗表示中</strong>: スタッフを選択すると個人目標を編集できます。';
        }
    }

    calculateGoal();
}

// スタッフ別内訳テーブルを更新
function updateGoalStaffBreakdown(context, yearMonth) {
    const tbody = document.getElementById('goal-staff-breakdown-body');
    if (!tbody) return;

    const fmt = n => n.toLocaleString();
    let rows = [];

    if (context.type === 'store') {
        // 特定店舗のスタッフ
        const staffList = STAFF_ROSTER[context.store] || [];
        staffList.forEach(staff => {
            const staffGoal = getStaffGoal(context.store, staff, yearMonth);
            const monthlyTarget = ((staffGoal.weekdays || 0) * (staffGoal.weekdayTarget || 0)) +
                                  ((staffGoal.weekends || 0) * (staffGoal.weekendTarget || 0));
            rows.push({
                name: staff,
                store: context.store,
                monthlyTarget,
                retail: staffGoal.retail || 0,
                newCustomers: staffGoal.newCustomers || 0,
                existingCustomers: staffGoal.existingCustomers || 0
            });
        });
    } else {
        // 全店舗のスタッフ
        Object.keys(STAFF_ROSTER).forEach(store => {
            const staffList = STAFF_ROSTER[store] || [];
            const storeName = store === 'chiba' ? '千葉店' : store === 'honatsugi' ? '本厚木店' : store === 'yamato' ? '大和店' : store;
            staffList.forEach(staff => {
                const staffGoal = getStaffGoal(store, staff, yearMonth);
                const monthlyTarget = ((staffGoal.weekdays || 0) * (staffGoal.weekdayTarget || 0)) +
                                      ((staffGoal.weekends || 0) * (staffGoal.weekendTarget || 0));
                rows.push({
                    name: `${storeName} - ${staff}`,
                    store,
                    monthlyTarget,
                    retail: staffGoal.retail || 0,
                    newCustomers: staffGoal.newCustomers || 0,
                    existingCustomers: staffGoal.existingCustomers || 0
                });
            });
        });
    }

    tbody.innerHTML = rows.map(r => `
        <tr class="border-b border-mavie-100 hover:bg-mavie-50">
            <td class="py-2 px-3 text-mavie-800 font-medium">${r.name}</td>
            <td class="py-2 px-3 text-right text-mavie-700">¥${fmt(r.monthlyTarget)}</td>
            <td class="py-2 px-3 text-right text-mavie-600">¥${fmt(r.retail)}</td>
            <td class="py-2 px-3 text-right text-mavie-600">${r.newCustomers}名</td>
            <td class="py-2 px-3 text-right text-mavie-600">${r.existingCustomers}名</td>
        </tr>
    `).join('');
}

function saveBaseSalaryAndCalculate() {
    const context = getCurrentGoalContextForGoalTab();
    if (context.type === 'staff') {
        const baseSalary = parseInt(document.getElementById('goal-base-salary').value) || DEFAULT_BASE_SALARY;
        saveStaffBaseSalary(context.store, context.staff, baseSalary);
    }
    calculateGoal();
}

/**
 * 過去の実績データから平日・休日の売上傾向を分析
 */
function analyzeWeekdayWeekendPerformance() {
    const filtered = getFilteredData();

    let weekdayTotal = 0, weekdayCount = 0;
    let weekendTotal = 0, weekendCount = 0;

    filtered.forEach(record => {
        // 日付を解析
        const dateParts = record.date.split('/');
        if (dateParts.length === 3) {
            const year = parseInt(dateParts[0]);
            const month = parseInt(dateParts[1]) - 1; // 月は0-indexed
            const day = parseInt(dateParts[2]);
            const date = new Date(year, month, day);
            const dayOfWeek = date.getDay(); // 0=日, 1=月, ..., 6=土

            const sales = record.sales.cash + record.sales.credit + record.sales.qr;

            // 平日: 月〜金 (1-5), 休日: 土日 (0, 6)
            if (dayOfWeek >= 1 && dayOfWeek <= 5) {
                weekdayTotal += sales;
                weekdayCount++;
            } else {
                weekendTotal += sales;
                weekendCount++;
            }
        }
    });

    const weekdayAvg = weekdayCount > 0 ? Math.round(weekdayTotal / weekdayCount) : 40000;
    const weekendAvg = weekendCount > 0 ? Math.round(weekendTotal / weekendCount) : 50000;
    const ratio = weekdayAvg > 0 ? (weekendAvg / weekdayAvg).toFixed(2) : 1.25;

    return {
        weekdayAvg,
        weekendAvg,
        ratio: parseFloat(ratio),
        weekdayCount,
        weekendCount
    };
}

/**
 * 月間目標金額から平日・休日のデイリー目標を自動計算
 */
function calculateGoalFromMonthlyTarget() {
    const monthlyTarget = parseInt(document.getElementById('goal-monthly-target').value) || 0;
    const weekdays = parseInt(document.getElementById('goal-weekdays').value) || 0;
    const weekends = parseInt(document.getElementById('goal-weekends').value) || 0;

    // 過去の実績を分析
    const performance = analyzeWeekdayWeekendPerformance();

    // 過去の実績データを表示
    const fmt = n => n.toLocaleString();
    document.getElementById('goal-past-weekday-avg').innerText = `¥${fmt(performance.weekdayAvg)}`;
    document.getElementById('goal-past-weekend-avg').innerText = `¥${fmt(performance.weekendAvg)}`;
    document.getElementById('goal-past-ratio').innerText = `${performance.ratio}倍`;

    // 平日と休日の売上比率を使って月間目標を配分
    // 月間目標 = (平日日数 × 平日目標) + (休日日数 × 休日目標)
    // 休日目標 = 平日目標 × ratio
    // 月間目標 = (平日日数 × 平日目標) + (休日日数 × 平日目標 × ratio)
    // 月間目標 = 平日目標 × (平日日数 + 休日日数 × ratio)
    // 平日目標 = 月間目標 / (平日日数 + 休日日数 × ratio)

    const totalDays = weekdays + weekends;
    if (totalDays === 0) {
        alert('平日または休日の出勤日数を入力してください');
        return;
    }

    const weekdayTarget = Math.round(monthlyTarget / (weekdays + weekends * performance.ratio));
    const weekendTarget = Math.round(weekdayTarget * performance.ratio);

    // hiddenフィールドに値を設定
    document.getElementById('goal-weekday-target').value = weekdayTarget;
    document.getElementById('goal-weekend-target').value = weekendTarget;

    // 表示用フィールドに値を設定
    document.getElementById('goal-weekday-target-display').innerText = `¥${fmt(weekdayTarget)}`;
    document.getElementById('goal-weekend-target-display').innerText = `¥${fmt(weekendTarget)}`;

    // 元のcalculateGoal関数を呼び出して目標を更新
    calculateGoal();
}

function calculateGoal() {
    const context = getCurrentGoalContextForGoalTab();
    const yearMonth = getCurrentGoalMonth();

    let weekdays, weekends, weekdayTarget, weekendTarget, retailTarget;
    let newCustomersTarget, existingCustomersTarget, unitPriceTarget;
    let newResRateTarget, resRateTarget, reviews5StarTarget;
    let totalGoal;

    if (context.type === 'staff') {
        // スタッフ個人: 入力フィールドから値を取得
        weekdays = parseInt(document.getElementById('goal-weekdays').value) || 0;
        weekends = parseInt(document.getElementById('goal-weekends').value) || 0;
        weekdayTarget = parseInt(document.getElementById('goal-weekday-target').value) || 0;
        weekendTarget = parseInt(document.getElementById('goal-weekend-target').value) || 0;
        retailTarget = parseInt(document.getElementById('goal-retail').value) || 0;
        newCustomersTarget = parseInt(document.getElementById('goal-new-customers').value) || 0;
        existingCustomersTarget = parseInt(document.getElementById('goal-existing-customers').value) || 0;
        unitPriceTarget = parseInt(document.getElementById('goal-unit-price').value) || 0;
        newResRateTarget = parseInt(document.getElementById('goal-new-reservation-rate').value) || 0;
        resRateTarget = parseInt(document.getElementById('goal-reservation-rate').value) || 0;
        reviews5StarTarget = parseInt(document.getElementById('goal-reviews-5star').value) || 0;

        totalGoal = (weekdays * weekdayTarget) + (weekends * weekendTarget);

        // 目標を保存
        const goalData = {
            weekdays, weekends, weekdayTarget, weekendTarget,
            retail: retailTarget,
            newCustomers: newCustomersTarget,
            existingCustomers: existingCustomersTarget,
            unitPrice: unitPriceTarget,
            newReservationRate: newResRateTarget,
            reservationRate: resRateTarget,
            reviews5Star: reviews5StarTarget
        };
        saveStaffGoal(context.store, context.staff, goalData);

        // Show save notice briefly
        const saveNotice = document.getElementById('goal-save-notice');
        saveNotice.style.display = 'block';
        setTimeout(() => {
            saveNotice.style.display = 'none';
        }, 2000);
    } else {
        // 店舗または全店舗: 集計データから値を取得
        let goalData;
        if (context.type === 'store') {
            goalData = getStoreAggregateGoal(context.store, yearMonth);
        } else {
            goalData = getAllStoresAggregateGoal(yearMonth);
        }

        weekdays = goalData.weekdays;
        weekends = goalData.weekends;
        weekdayTarget = goalData.weekdayTarget;
        weekendTarget = goalData.weekendTarget;
        retailTarget = goalData.retail;
        newCustomersTarget = goalData.newCustomers;
        existingCustomersTarget = goalData.existingCustomers;
        unitPriceTarget = goalData.unitPrice;
        newResRateTarget = goalData.newReservationRate;
        resRateTarget = goalData.reservationRate;
        reviews5StarTarget = goalData.reviews5Star || 0;

        // 店舗/全店舗の場合は正確な合計を使用
        totalGoal = goalData.monthlyTargetSum || 0;
    }

    // Update global goal
    monthlyGoal = totalGoal;

    // Format numbers
    const fmt = n => n.toLocaleString();

    // Update display
    document.getElementById('goal-total').innerText = `¥${fmt(totalGoal)}`;
    const weekdayTotal = weekdays * weekdayTarget;
    const weekendTotal = weekends * weekendTarget;
    document.getElementById('goal-weekday-total').innerText = `¥${fmt(weekdayTotal)}`;
    document.getElementById('goal-weekend-total').innerText = `¥${fmt(weekendTotal)}`;
    document.getElementById('goal-retail-display').innerText = `¥${fmt(retailTarget)}`;

    // Update display fields for daily targets (only for staff view)
    if (context.type === 'staff') {
        document.getElementById('goal-weekday-target-display').innerText = `¥${fmt(weekdayTarget)}`;
        document.getElementById('goal-weekend-target-display').innerText = `¥${fmt(weekendTarget)}`;
    }

    // Get current metrics
    const filtered = getFilteredData();
    const metrics = calculateMetrics(filtered);

    // Update achievement status - Sales
    const currentSales = metrics.salesTotal;
    const percentage = totalGoal > 0 ? Math.round((currentSales / totalGoal) * 100) : 0;
    const remaining = totalGoal - currentSales;
    const progressWidth = Math.min(percentage, 100);

    document.getElementById('goal-current-sales').innerText = `¥${fmt(currentSales)}`;
    document.getElementById('goal-percentage').innerText = `${percentage}%`;
    document.getElementById('goal-remaining').innerText = `¥${fmt(Math.max(remaining, 0))}`;
    document.getElementById('goal-progress-bar').style.width = `${progressWidth}%`;

    // Update achievement status - Customers & KPIs
    document.getElementById('goal-new-current').innerText = `${fmt(metrics.customersNew)}名`;
    document.getElementById('goal-new-target').innerText = `${fmt(newCustomersTarget)}名`;

    document.getElementById('goal-existing-current').innerText = `${fmt(metrics.customersExisting)}名`;
    document.getElementById('goal-existing-target').innerText = `${fmt(existingCustomersTarget)}名`;

    const currentUnitPrice = metrics.customersTotal > 0 ? Math.round(metrics.salesTotal / metrics.customersTotal) : 0;
    document.getElementById('goal-unit-price-current').innerText = `¥${fmt(currentUnitPrice)}`;
    document.getElementById('goal-unit-price-target').innerText = `¥${fmt(unitPriceTarget)}`;

    const currentNewResRate = metrics.hpbNewCount > 0 ? (((metrics.nextRes.hpbNew + metrics.nextRes.mininaiNew) / metrics.hpbNewCount)*100).toFixed(1) : 0;
    document.getElementById('goal-new-res-rate-current').innerText = `${currentNewResRate}%`;
    document.getElementById('goal-new-res-rate-target').innerText = `${newResRateTarget}%`;

    const currentResRate = metrics.customersTotal > 0 ? ((metrics.nextRes.total / metrics.customersTotal)*100).toFixed(1) : 0;
    document.getElementById('goal-res-rate-current').innerText = `${currentResRate}%`;
    document.getElementById('goal-res-rate-target').innerText = `${resRateTarget}%`;

    // Update main KPI card
    updateDashboard();
}

function switchTab(id) {
    document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
    const target = document.getElementById(`content-${id}`);
    target.classList.remove('hidden');
    // Re-trigger the fade-in animation on every switch
    target.classList.remove('animate-fade-in');
    void target.offsetWidth;
    target.classList.add('animate-fade-in');
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active', 'text-accent-800');
        btn.classList.add('text-surface-500');
    });
    const active = document.getElementById(`tab-${id}`);
    active.classList.remove('text-surface-500');
    active.classList.add('active', 'text-accent-800');

    // Sidebar state sync
    document.querySelectorAll('.sidebar-item').forEach(btn => btn.classList.remove('active'));
    const sideActive = document.querySelector(`.sidebar-item[data-tab="${id}"]`);
    if (sideActive) sideActive.classList.add('active');

    // Bottom-nav state sync (モバイル): 一致する項目があれば active、無ければメニューを光らせない
    document.querySelectorAll('.bottom-nav-item[data-tab]').forEach(btn => btn.classList.remove('active'));
    const bottomActive = document.querySelector(`.bottom-nav-item[data-tab="${id}"]`);
    if (bottomActive) bottomActive.classList.add('active');

    // Load goal when switching to goal tab
    if (id === 'goal') {
        // Initialize month selector for goal tab
        initGoalMonthSelector();
        // Initialize goal selectors from main selectors
        const mainStore = document.getElementById('store-selector').value;
        const mainStaff = document.getElementById('staff-selector').value;
        document.getElementById('goal-store-selector').value = mainStore;
        handleGoalStoreChange();
        // Try to set staff selector if available
        const goalStaffSel = document.getElementById('goal-staff-selector');
        const matchingOption = Array.from(goalStaffSel.options).find(opt => opt.value === mainStaff);
        if (matchingOption) {
            goalStaffSel.value = mainStaff;
        }
        loadGoalInputs();
    }

    // Update settings list when switching to settings tab
    if (id === 'settings') {
        updateSettingsList();
        loadGeminiApiKey(); // Load API key
        loadSpreadsheetApiUrl(); // Load spreadsheet API URL
        loadCustomerApiUrl(); // Load customer API URL
        // Re-render Lucide icons for settings tab
        if (typeof lucide !== 'undefined') lucide.createIcons();
    }

    // Render calendar when switching to calendar tab
    if (id === 'calendar') {
        renderCalendar();
    }

    // Load marketing data when switching to marketing tab
    if (id === 'marketing') {
        if (customerData.length === 0) {
            refreshCustomerData();
        } else {
            updateMarketingDashboard();
        }
    }

    // Load counseling results when switching to staff-dashboard or counseling-results tab
    if (id === 'staff-dashboard' || id === 'counseling-results') {
        if (customerData.length === 0) {
            refreshCounselingResults();
        } else {
            renderCounselingResults();
        }
    }

    // Load counseling results when switching to marketing tab
    if (id === 'marketing') {
        if (counselingData.length === 0) {
            // 当日分を先に読み込んで高速表示
            loadCustomerDataFast().then(() => {
                renderAdminCounselingResults();
            });
        } else {
            renderAdminCounselingResults();
        }
    }

    // Load incentive data when switching to incentive tab
    if (id === 'incentive') {
        updateIncentiveTab();
    }
}

function generateStaffURL() {
    const storeVal = document.getElementById('store-selector').value;
    const staffVal = document.getElementById('staff-selector').value;

    if (staffVal === 'all') {
        alert('スタッフを選択してください');
        return;
    }

    const baseURL = window.location.origin + window.location.pathname;
    // パラメータを小文字に正規化
    const staffURL = `${baseURL}?store=${storeVal.toLowerCase()}&staff=${staffVal.toLowerCase()}`;

    // 別タブで開く
    window.open(staffURL, '_blank');
}

// --- COUNSELING DATA FUNCTIONS ---
let counselingData = [];

async function refreshCounselingData() {
    const loaded = await loadCustomerDataForCounseling();
    if (loaded) {
        filterCounselingData();
    }
}

async function loadCustomerDataForCounseling(todayOnly = false) {
    CUSTOMER_API_URL = localStorage.getItem(CUSTOMER_API_KEY) || '';
    // 顧客API URLが未設定の場合、売上APIのURLをフォールバック（同じGAS Web App）
    if (!CUSTOMER_API_URL) {
        CUSTOMER_API_URL = API_URL || DEFAULT_API_URL;
    }
    if (!CUSTOMER_API_URL) {
        return false;
    }
    try {
        let action, url;

        // スタッフ専用URL時は店舗別取得（超高速：30秒キャッシュ）
        if (lockedStore) {
            action = 'get_customers_by_store';
            url = CUSTOMER_API_URL.includes('?')
                ? `${CUSTOMER_API_URL}&action=${action}&store=${lockedStore}`
                : `${CUSTOMER_API_URL}?action=${action}&store=${lockedStore}`;
        } else {
            // 当日分のみの場合は専用エンドポイント使用（高速）
            action = todayOnly ? 'get_customers_today' : 'get_customers';
            url = CUSTOMER_API_URL.includes('?')
                ? `${CUSTOMER_API_URL}&action=${action}`
                : `${CUSTOMER_API_URL}?action=${action}`;
        }

        const response = await fetch(url, {
            method: 'GET',
            headers: { 'Accept': 'application/json' }
        });
        const result = await response.json();
        if (result.status === 'success' && result.data) {
            if (todayOnly && !lockedStore) {
                // 当日分のみの場合は既存データとマージ
                const existingData = counselingData || [];
                const todayData = result.data;
                const todayIds = new Set(todayData.map(d => d.id));
                const otherData = existingData.filter(d => !todayIds.has(d.id));
                counselingData = [...todayData, ...otherData];
            } else {
                counselingData = result.data;
            }
        } else if (Array.isArray(result)) {
            counselingData = result;
        } else {
            counselingData = [];
        }
        return true;
    } catch (e) {
        console.error('顧客データの取得に失敗:', e);
        return false;
    }
}

// 当日分を先に読み込み、その後全データを読み込む（高速表示）
async function loadCustomerDataFast() {
    // まず当日分を高速表示
    const todayLoaded = await loadCustomerDataForCounseling(true);
    if (todayLoaded) {
        filterCounselingData();
    }
    // バックグラウンドで全データを読み込み
    setTimeout(async () => {
        await loadCustomerDataForCounseling(false);
        filterCounselingData();
    }, 100);
    return todayLoaded;
}

function filterCounselingData() {
    const container = document.getElementById('counseling-cards-container');
    if (!container) return;

    const filterVal = document.getElementById('counseling-filter')?.value || 'today';
    const searchVal = document.getElementById('counseling-search')?.value?.toLowerCase() || '';

    // スタッフ専用ページの場合、そのスタッフのデータのみ表示
    const storeFilter = lockedStore || document.getElementById('store-selector')?.value;
    const staffFilter = lockedStaff || document.getElementById('staff-selector')?.value;

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayStr = `${today.getFullYear()}/${today.getMonth() + 1}/${today.getDate()}`;

    const weekAgo = new Date(today);
    weekAgo.setDate(weekAgo.getDate() - 7);

    let filtered = counselingData.filter(c => {
        // 店舗フィルター
        if (storeFilter && storeFilter !== 'all' && c.store !== storeFilter) return false;
        // スタッフフィルター（担当スタッフがある場合）- 大文字小文字を区別しない
        if (staffFilter && staffFilter !== 'all' && c.staff && c.staff.toLowerCase() !== staffFilter.toLowerCase()) return false;
        // 検索フィルター
        if (searchVal && !c.name?.toLowerCase().includes(searchVal)) return false;
        return true;
    });

    // 日付フィルター
    if (filterVal === 'today') {
        filtered = filtered.filter(c => {
            if (!c.date) return false;
            return c.date === todayStr || c.date.replace(/\//g, '/') === todayStr;
        });
    } else if (filterVal === 'week') {
        filtered = filtered.filter(c => {
            if (!c.timestamp) return true;
            return c.timestamp >= weekAgo.getTime();
        });
    } else if (filterVal === 'recent') {
        filtered = filtered.slice(-10).reverse();
    }

    // 新しい順にソート
    filtered.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

    if (filtered.length === 0) {
        container.innerHTML = `
            <div class="text-center py-8 text-mavie-500">
                <i data-lucide="users" class="w-12 h-12 mx-auto mb-3 opacity-50"></i>
                <p class="text-sm">${filterVal === 'today' ? '今日の予約はありません' : '該当するお客様がいません'}</p>
            </div>
        `;
        if (typeof lucide !== 'undefined') lucide.createIcons();
        return;
    }

    container.innerHTML = filtered.map(c => renderCounselingCard(c)).join('');
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function renderCounselingCard(customer) {
    const hasAllergy = customer.allergy && customer.allergy.trim() !== '';
    const hasEyebrowTrouble = customer.eyebrowTrouble && customer.eyebrowTrouble.trim() !== '' && !customer.eyebrowTrouble.includes('特になし');
    const hasLashTrouble = customer.lashTrouble && customer.lashTrouble.trim() !== '' && !customer.lashTrouble.includes('特になし');
    const hasWarning = hasAllergy || hasEyebrowTrouble || hasLashTrouble;
    const alertClass = hasWarning ? 'border-l-4 border-red-400' : '';

    // 眉毛メニュー情報があるかチェック
    const hasEyebrowInfo = customer.eyebrowConcern || customer.eyebrowDesign || customer.eyebrowDesignImage || customer.eyebrowFrequency || customer.eyebrowImpression;
    // まつ毛メニュー情報があるかチェック
    const hasLashInfo = customer.lashDesign || customer.lashDesignImage || customer.lashFrequency || customer.lashEyeLook || customer.lashContact;

    return `
        <div class="bg-gradient-to-r from-white to-mavie-50 rounded-lg border border-mavie-200 p-4 ${alertClass} hover:shadow-md transition">
            <div class="flex items-start justify-between mb-3">
                <div>
                    <h4 class="font-bold text-mavie-800 text-lg">${customer.name || '名前未登録'}</h4>
                    ${customer.nameKana ? `<p class="text-xs text-mavie-400">${customer.nameKana}</p>` : ''}
                    <p class="text-xs text-mavie-500">${customer.date || ''} ${customer.storeName || ''}</p>
                </div>
                <div class="flex gap-1 flex-wrap justify-end">
                    ${hasAllergy ? '<span class="bg-red-100 text-red-700 text-xs px-2 py-1 rounded font-bold">アレルギー</span>' : ''}
                    ${hasEyebrowTrouble ? '<span class="bg-orange-100 text-orange-700 text-xs px-2 py-1 rounded font-bold">眉肌トラブル</span>' : ''}
                    ${hasLashTrouble ? '<span class="bg-orange-100 text-orange-700 text-xs px-2 py-1 rounded font-bold">まつ毛肌トラブル</span>' : ''}
                    ${customer.snsOk && customer.snsOk.includes('OK') ? '<span class="bg-green-100 text-green-700 text-xs px-2 py-1 rounded font-bold">SNS OK</span>' : ''}
                </div>
            </div>

            <div class="grid grid-cols-2 md:grid-cols-4 gap-3 text-sm mb-3">
                ${customer.birthday ? `<div><span class="text-mavie-500">生年月日:</span> <span class="font-medium">${customer.birthday}</span></div>` : ''}
                ${customer.phone ? `<div><span class="text-mavie-500">電話:</span> <span class="font-medium">${customer.phone}</span></div>` : ''}
                ${customer.address ? `<div class="col-span-2"><span class="text-mavie-500">住所:</span> <span class="font-medium">${customer.address}</span></div>` : ''}
                ${customer.job ? `<div><span class="text-mavie-500">職業:</span> <span class="font-medium">${customer.job}</span></div>` : ''}
            </div>

            ${hasEyebrowInfo ? `
            <div class="bg-amber-50 rounded p-3 mb-3">
                <p class="text-xs font-bold text-amber-700 mb-2">🪶 眉毛メニュー情報</p>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm">
                    ${customer.eyebrowConcern ? `<div><span class="text-mavie-500">お悩み:</span> <span class="text-mavie-800">${customer.eyebrowConcern}</span></div>` : ''}
                    ${customer.eyebrowDesign ? `<div><span class="text-mavie-500">ご希望デザイン:</span> <span class="text-mavie-800">${customer.eyebrowDesign}</span></div>` : ''}
                    ${customer.eyebrowDesignImage ? `<div><span class="text-mavie-500">デザインイメージ:</span> <span class="text-mavie-800 font-medium">${customer.eyebrowDesignImage}</span></div>` : ''}
                    ${customer.eyebrowImpression ? `<div><span class="text-mavie-500">印象:</span> <span class="text-mavie-800">${customer.eyebrowImpression}</span></div>` : ''}
                    ${customer.eyebrowFrequency ? `<div><span class="text-mavie-500">利用頻度:</span> <span class="text-mavie-800">${customer.eyebrowFrequency}</span></div>` : ''}
                    ${customer.eyebrowLastCare ? `<div class="col-span-2"><span class="text-mavie-500">最後のお手入れ:</span> <span class="text-mavie-800">${customer.eyebrowLastCare}</span></div>` : ''}
                </div>
            </div>
            ` : ''}

            ${hasLashInfo ? `
            <div class="bg-purple-50 rounded p-3 mb-3">
                <p class="text-xs font-bold text-purple-700 mb-2">👁️ まつ毛メニュー情報</p>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-2 text-sm">
                    ${customer.lashDesign ? `<div><span class="text-mavie-500">ご希望デザイン:</span> <span class="text-mavie-800">${customer.lashDesign}</span></div>` : ''}
                    ${customer.lashDesignImage ? `<div><span class="text-mavie-500">デザインイメージ:</span> <span class="text-mavie-800 font-medium">${customer.lashDesignImage}</span></div>` : ''}
                    ${customer.lashFrequency ? `<div><span class="text-mavie-500">利用頻度:</span> <span class="text-mavie-800">${customer.lashFrequency}</span></div>` : ''}
                    ${customer.lashEyeLook ? `<div><span class="text-mavie-500">目の見え方:</span> <span class="text-mavie-800">${customer.lashEyeLook}</span></div>` : ''}
                    ${customer.lashContact ? `<div><span class="text-mavie-500">コンタクト:</span> <span class="text-mavie-800">${customer.lashContact}</span></div>` : ''}
                </div>
            </div>
            ` : ''}

            ${hasWarning ? `
            <div class="bg-red-50 rounded p-3 mb-3 border border-red-200">
                <p class="text-xs font-bold text-red-700 mb-2">⚠️ 注意事項</p>
                <div class="text-sm space-y-1">
                    ${hasAllergy ? `<div><span class="text-red-600 font-medium">アレルギー:</span> <span class="text-red-800">${customer.allergy}</span></div>` : ''}
                    ${hasEyebrowTrouble ? `<div><span class="text-orange-600 font-medium">眉施術後トラブル:</span> <span class="text-orange-800">${customer.eyebrowTrouble}</span></div>` : ''}
                    ${hasLashTrouble ? `<div><span class="text-orange-600 font-medium">まつ毛施術後トラブル:</span> <span class="text-orange-800">${customer.lashTrouble}</span></div>` : ''}
                </div>
            </div>
            ` : ''}

            ${customer.visitReason ? `
            <div class="bg-blue-50 rounded p-3 text-sm mb-3">
                <span class="text-blue-600 font-medium">来店理由:</span> <span class="text-blue-800">${customer.visitReason}</span>
                ${customer.fromOtherSalon ? `<div class="mt-1"><span class="text-blue-600 font-medium">他サロンからの理由:</span> <span class="text-blue-800">${customer.fromOtherSalon}</span></div>` : ''}
                ${customer.dissatisfaction ? `<div class="mt-1"><span class="text-blue-600 font-medium">不満理由:</span> <span class="text-blue-800">${customer.dissatisfaction}</span></div>` : ''}
            </div>
            ` : ''}
        </div>
    `;
}

// スタッフダッシュボード更新時にカウンセリングデータも読み込む
function loadCounselingForStaffDashboard() {
    if (lockedStaff || document.getElementById('staff-selector')?.value !== 'all') {
        // 当日分を先に読み込んで高速表示
        loadCustomerDataFast();
    }
}

// 管理ページ用カウンセリング表示関数
async function refreshAdminCounselingResults() {
    const loaded = await loadCustomerDataForCounseling();
    if (loaded) {
        renderAdminCounselingResults();
        showUpdateNotification('カウンセリング回答を更新しました');
    }
}

function renderAdminCounselingResults() {
    const container = document.getElementById('admin-counseling-container');
    if (!container) return;

    if (!counselingData || counselingData.length === 0) {
        container.innerHTML = `
            <div class="text-center py-8 text-mavie-400">
                <i data-lucide="clipboard-list" class="w-12 h-12 mx-auto mb-3 opacity-50"></i>
                <p class="text-sm">顧客データがありません</p>
            </div>
        `;
        if (typeof lucide !== 'undefined') lucide.createIcons();
        return;
    }

    const storeFilter = document.getElementById('admin-counseling-store-filter')?.value || 'all';
    const searchTerm = document.getElementById('admin-counseling-search')?.value?.toLowerCase() || '';
    const dateFilter = document.getElementById('admin-counseling-date-filter')?.value || 'month';

    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

    let filtered = counselingData.filter(c => {
        // 店舗フィルター
        if (storeFilter !== 'all' && c.store !== storeFilter) return false;

        // 検索フィルター
        if (searchTerm && !c.name?.toLowerCase().includes(searchTerm)) return false;

        // 日付フィルター
        if (dateFilter !== 'all' && c.timestamp) {
            const recordDate = parseDate(c.timestamp);
            if (dateFilter === 'week' && recordDate < weekAgo) return false;
            if (dateFilter === 'month' && recordDate < monthStart) return false;
        }

        return true;
    });

    // 新しい順にソート
    filtered.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));

    // 件数表示
    document.getElementById('admin-counseling-count').textContent = filtered.length;

    if (filtered.length === 0) {
        container.innerHTML = `
            <div class="text-center py-8 text-mavie-400">
                <i data-lucide="search-x" class="w-12 h-12 mx-auto mb-3 opacity-50"></i>
                <p class="text-sm">該当する回答がありません</p>
            </div>
        `;
        if (typeof lucide !== 'undefined') lucide.createIcons();
        return;
    }

    container.innerHTML = filtered.map(c => renderCounselingCard(c)).join('');
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

// --- STAFF & STORE MANAGEMENT ---
function updateSettingsList() {
    // Update store list
    const storeList = document.getElementById('store-list');
    storeList.innerHTML = '';
    Object.keys(STAFF_ROSTER).forEach(storeId => {
        const storeDiv = document.createElement('div');
        storeDiv.className = 'flex items-center justify-between bg-mavie-50 p-3 rounded';
        storeDiv.innerHTML = `
            <span class="font-medium">${storeId}</span>
            <button onclick="deleteStore('${storeId}')" class="text-red-600 hover:text-red-800 text-sm font-bold">削除</button>
        `;
        storeList.appendChild(storeDiv);
    });

    // Update staff list
    const staffList = document.getElementById('staff-list');
    staffList.innerHTML = '';
    const baseUrl = window.location.origin + window.location.pathname;
    Object.keys(STAFF_ROSTER).forEach(storeId => {
        const storeName = storeId === 'chiba' ? '千葉店' : storeId === 'honatsugi' ? '本厚木店' : storeId === 'yamato' ? '大和店' : storeId;
        STAFF_ROSTER[storeId].forEach(staffName => {
            const staffDiv = document.createElement('div');
            staffDiv.className = 'bg-mavie-50 p-3 rounded';
            staffDiv.innerHTML = `
                <div class="flex items-center justify-between">
                    <div class="flex items-center gap-2">
                        <span class="font-medium text-mavie-800"><strong>${storeName}</strong>: ${staffName}</span>
                        <button onclick="openStaffUrl('${storeId}', '${staffName}')" class="bg-primary-500 text-white px-3 py-1 rounded text-xs font-bold hover:bg-primary-600 transition flex items-center gap-1">
                            <i data-lucide="external-link" class="w-3 h-3"></i>
                            専用ページ
                        </button>
                    </div>
                    <button onclick="deleteStaff('${storeId}', '${staffName}')" class="text-red-600 hover:text-red-800 text-sm font-bold">削除</button>
                </div>
            `;
            staffList.appendChild(staffDiv);
        });
    });

    // Lucide iconsを更新
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }

    // Update staff store selector
    const staffStoreSel = document.getElementById('staff-store-selector');
    staffStoreSel.innerHTML = '<option value="">店舗を選択</option>';
    Object.keys(STAFF_ROSTER).forEach(storeId => {
        const opt = document.createElement('option');
        opt.value = storeId;
        opt.text = storeId;
        staffStoreSel.appendChild(opt);
    });

    // Update password settings UI
    initPasswordSettingsUI();
}

async function addStore() {
    const storeId = document.getElementById('new-store-id').value.trim();
    const storeName = document.getElementById('new-store-name').value.trim();

    if (!storeId) {
        alert('店舗IDを入力してください');
        return;
    }

    if (STAFF_ROSTER[storeId]) {
        alert('この店舗IDは既に存在します');
        return;
    }

    // Confirmation dialog
    if (!confirm(`店舗「${storeId}」を追加してよろしいですか？\n\n店舗ID: ${storeId}\n店舗名: ${storeName || '(未設定)'}\n\nこの操作は取り消せませんので、内容をご確認ください。`)) {
        return;
    }

    STAFF_ROSTER[storeId] = [];
    document.getElementById('new-store-id').value = '';
    document.getElementById('new-store-name').value = '';

    updateSettingsList();
    const saved = await saveSettingsToSpreadsheet(true);
    if (saved) {
        showSettingsToast(`店舗「${storeId}」を追加し、スプレッドシートに保存しました`);
    } else {
        showSettingsToast(`店舗「${storeId}」を追加しました（ローカル保存のみ）`, 'warning');
    }
}

async function deleteStore(storeId) {
    if (!confirm(`店舗「${storeId}」を削除しますか？関連するスタッフと目標データも削除されます。`)) {
        return;
    }

    delete STAFF_ROSTER[storeId];

    // Delete related goal data
    const goals = loadGoalsFromStorage();
    if (goals[storeId]) {
        delete goals[storeId];
        saveGoalsToStorage(goals);
    }

    // Delete related salary data
    const salaries = loadBaseSalariesFromStorage();
    if (salaries[storeId]) {
        delete salaries[storeId];
        saveBaseSalariesToStorage(salaries);
    }

    updateSettingsList();
    const saved = await saveSettingsToSpreadsheet(true);
    if (saved) {
        showSettingsToast(`店舗「${storeId}」を削除し、スプレッドシートに保存しました`);
    } else {
        showSettingsToast(`店舗「${storeId}」を削除しました（ローカル保存のみ）`, 'warning');
    }
}

async function addStaff() {
    const storeId = document.getElementById('staff-store-selector').value;
    const staffName = document.getElementById('new-staff-name').value.trim();

    if (!storeId) {
        alert('店舗を選択してください');
        return;
    }

    if (!staffName) {
        alert('スタッフ名を入力してください');
        return;
    }

    if (!STAFF_ROSTER[storeId]) {
        alert('選択された店舗が存在しません');
        return;
    }

    // 大文字小文字を区別しないで重複チェック
    if (STAFF_ROSTER[storeId].some(s => s.toLowerCase() === staffName.toLowerCase())) {
        alert('このスタッフは既に存在します');
        return;
    }

    // Confirmation dialog
    if (!confirm(`スタッフ「${staffName}」を${storeId}に追加してよろしいですか？\n\nスタッフ名: ${staffName}\n店舗: ${storeId}\n\nこの操作は取り消せませんので、内容をご確認ください。`)) {
        return;
    }

    STAFF_ROSTER[storeId].push(staffName);
    document.getElementById('new-staff-name').value = '';

    updateSettingsList();
    const saved = await saveSettingsToSpreadsheet(true);
    if (saved) {
        showSettingsToast(`スタッフ「${staffName}」を${storeId}に追加し、スプレッドシートに保存しました`);
    } else {
        showSettingsToast(`スタッフ「${staffName}」を追加しました（ローカル保存のみ）`, 'warning');
    }
}

async function deleteStaff(storeId, staffName) {
    if (!confirm(`スタッフ「${staffName}」を削除しますか？関連する目標データも削除されます。`)) {
        return;
    }

    const index = STAFF_ROSTER[storeId].indexOf(staffName);
    if (index > -1) {
        STAFF_ROSTER[storeId].splice(index, 1);
    }

    // Delete related goal data
    const goals = loadGoalsFromStorage();
    if (goals[storeId] && goals[storeId][staffName]) {
        delete goals[storeId][staffName];
        saveGoalsToStorage(goals);
    }

    // Delete related salary data
    const salaries = loadBaseSalariesFromStorage();
    if (salaries[storeId] && salaries[storeId][staffName]) {
        delete salaries[storeId][staffName];
        saveBaseSalariesToStorage(salaries);
    }

    updateSettingsList();
    const saved = await saveSettingsToSpreadsheet(true);
    if (saved) {
        showSettingsToast(`スタッフ「${staffName}」を削除し、スプレッドシートに保存しました`);
    } else {
        showSettingsToast(`スタッフ「${staffName}」を削除しました（ローカル保存のみ）`, 'warning');
    }
}

// スタッフ専用URLを別タブで開く
function openStaffUrl(storeId, staffName) {
    const baseURL = window.location.origin + window.location.pathname;
    // パラメータを小文字に正規化
    const staffURL = `${baseURL}?store=${storeId.toLowerCase()}&staff=${staffName.toLowerCase()}`;
    window.open(staffURL, '_blank');
}

// --- STAFF PASSWORD MANAGEMENT FUNCTIONS ---
function updatePasswordStaffSelector() {
    const storeSel = document.getElementById('password-store-selector');
    const staffSel = document.getElementById('password-staff-selector');
    const selectedStore = storeSel.value;

    staffSel.innerHTML = '<option value="">スタッフを選択</option>';

    if (selectedStore && STAFF_ROSTER[selectedStore]) {
        STAFF_ROSTER[selectedStore].forEach(name => {
            const opt = document.createElement('option');
            opt.value = name;
            opt.text = name;
            staffSel.appendChild(opt);
        });
    }
}

function saveStaffPasswordSettings() {
    const storeId = document.getElementById('password-store-selector').value;
    const staffName = document.getElementById('password-staff-selector').value;
    const password = document.getElementById('staff-password-input').value;
    const statusEl = document.getElementById('staff-password-status');

    if (!storeId) {
        alert('店舗を選択してください');
        return;
    }

    if (!staffName) {
        alert('スタッフを選択してください');
        return;
    }

    setStaffPassword(storeId, staffName, password);

    // ステータス表示
    statusEl.classList.remove('hidden');
    if (password) {
        statusEl.innerHTML = '<p class="text-green-600 text-sm">✓ パスワードを設定しました</p>';
    } else {
        statusEl.innerHTML = '<p class="text-yellow-600 text-sm">✓ パスワードを無効化しました（ログイン不要）</p>';
    }

    // 入力をクリア
    document.getElementById('staff-password-input').value = '';

    // パスワード設定リストを更新
    updatePasswordList();

    // 3秒後にステータスを非表示
    setTimeout(() => {
        statusEl.classList.add('hidden');
    }, 3000);
}

function updatePasswordList() {
    const listContainer = document.getElementById('staff-password-list');
    if (!listContainer) return;

    const passwords = staffPasswordsCache || loadStaffPasswordsFromLocalStorage();
    let html = '';

    Object.keys(STAFF_ROSTER).forEach(storeId => {
        const storeName = storeId === 'chiba' ? '千葉店' : storeId === 'honatsugi' ? '本厚木店' : storeId === 'yamato' ? '大和店' : storeId;
        const staffList = STAFF_ROSTER[storeId];

        staffList.forEach(staff => {
            const hasPassword = passwords[storeId] && passwords[storeId][staff] && passwords[storeId][staff] !== '';
            const statusClass = hasPassword ? 'bg-green-100 text-green-800' : 'bg-gray-100 text-gray-500';
            const statusText = hasPassword ? '設定済み' : '未設定';
            const lockIcon = hasPassword ? 'lock' : 'unlock';

            html += `
                <div class="flex items-center justify-between p-2 rounded ${hasPassword ? 'bg-green-50' : 'bg-gray-50'} border border-${hasPassword ? 'green' : 'gray'}-200">
                    <div class="flex items-center gap-2">
                        <i data-lucide="${lockIcon}" class="w-4 h-4 ${hasPassword ? 'text-green-600' : 'text-gray-400'}"></i>
                        <span class="text-sm font-medium text-mavie-800">${storeName} / ${staff}</span>
                    </div>
                    <span class="text-xs px-2 py-1 rounded ${statusClass}">${statusText}</span>
                </div>
            `;
        });
    });

    listContainer.innerHTML = html || '<p class="text-sm text-mavie-500">スタッフが登録されていません</p>';

    // Lucide iconsを更新
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
}

function initPasswordSettingsUI() {
    const storeSel = document.getElementById('password-store-selector');
    if (!storeSel) return;

    // 店舗セレクタを初期化
    storeSel.innerHTML = '<option value="">店舗を選択</option>';
    Object.keys(STAFF_ROSTER).forEach(storeId => {
        const storeName = storeId === 'chiba' ? '千葉店' : storeId === 'honatsugi' ? '本厚木店' : storeId === 'yamato' ? '大和店' : storeId;
        const opt = document.createElement('option');
        opt.value = storeId;
        opt.text = storeName;
        storeSel.appendChild(opt);
    });

    // パスワード設定リストを更新
    updatePasswordList();
}

// 管理ページパスワードを保存
async function saveAdminPassword() {
    const password = document.getElementById('admin-password-input').value;
    const statusEl = document.getElementById('admin-password-status');
    const apiUrl = document.getElementById('spreadsheet-api-url')?.value;

    if (!apiUrl) {
        alert('先にスプレッドシートAPIのURLを設定してください');
        return;
    }

    try {
        const response = await fetch(apiUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain' },
            body: JSON.stringify({
                action: 'save_settings',
                settings: {
                    adminPassword: password
                }
            })
        });

        const result = await response.json();

        // ステータス表示
        statusEl.classList.remove('hidden');
        if (result.status === 'success') {
            if (password) {
                statusEl.innerHTML = '<p class="text-green-600 text-sm font-bold">✓ 管理パスワードを設定しました</p>';
            } else {
                statusEl.innerHTML = '<p class="text-yellow-600 text-sm font-bold">✓ 管理パスワードを無効化しました（ログイン不要）</p>';
            }
        } else {
            statusEl.innerHTML = '<p class="text-red-600 text-sm font-bold">✗ 保存に失敗しました</p>';
        }

        // 入力をクリア
        document.getElementById('admin-password-input').value = '';

        // 3秒後にステータスを非表示
        setTimeout(() => {
            statusEl.classList.add('hidden');
        }, 3000);
    } catch (error) {
        console.error('管理パスワード保存エラー:', error);
        statusEl.classList.remove('hidden');
        statusEl.innerHTML = '<p class="text-red-600 text-sm font-bold">✗ 保存に失敗しました</p>';
    }
}

// --- PASSWORD AUTHENTICATION FUNCTIONS ---
const SESSION_TOKEN_KEY = 'mavie_session_token';
const SESSION_PAGE_TYPE_KEY = 'mavie_session_page_type';
let authenticationRequired = false;
let authPageType = '';
let authStore = '';
let authStaff = '';

// パスワードダイアログを表示
function showPasswordDialog(pageType, store = '', staff = '') {
    authPageType = pageType;
    authStore = store;
    authStaff = staff;

    const dialog = document.getElementById('password-dialog');
    const title = document.getElementById('password-dialog-title');
    const message = document.getElementById('password-dialog-message');

    if (pageType === 'admin') {
        title.textContent = '管理ページ認証';
        message.textContent = '管理ページへのアクセスにはパスワードが必要です';
    } else {
        title.textContent = 'スタッフページ認証';
        message.textContent = `${store === 'chiba' ? '千葉店' : store === 'honatsugi' ? '本厚木店' : store === 'yamato' ? '大和店' : store} / ${staff} のページへのアクセスにはパスワードが必要です`;
    }

    dialog.classList.remove('hidden');
    document.getElementById('password-dialog-input').focus();

    // Lucide iconsを更新
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
}

// パスワード送信処理
async function handlePasswordSubmit() {
    const password = document.getElementById('password-dialog-input').value;
    const errorEl = document.getElementById('password-error');
    const apiUrl = localStorage.getItem(SPREADSHEET_API_KEY);

    if (!password) {
        errorEl.textContent = 'パスワードを入力してください';
        errorEl.classList.remove('hidden');
        return;
    }

    if (!apiUrl) {
        errorEl.textContent = 'APIが設定されていません';
        errorEl.classList.remove('hidden');
        return;
    }

    try {
        const response = await fetch(`${apiUrl}?action=verify_password&page_type=${authPageType}&store=${authStore}&staff=${authStaff}&password=${encodeURIComponent(password)}`);
        const result = await response.json();

        if (result.status === 'success') {
            // セッショントークンを保存
            localStorage.setItem(SESSION_TOKEN_KEY, result.sessionToken);
            localStorage.setItem(SESSION_PAGE_TYPE_KEY, authPageType);

            // ダイアログを閉じてページを初期化
            document.getElementById('password-dialog').classList.add('hidden');
            document.getElementById('password-dialog-input').value = '';
            errorEl.classList.add('hidden');

            // ページをリロード
            location.reload();
        } else {
            errorEl.textContent = result.message || 'パスワードが正しくありません';
            errorEl.classList.remove('hidden');
            document.getElementById('password-dialog-input').value = '';
            document.getElementById('password-dialog-input').focus();
        }
    } catch (error) {
        console.error('認証エラー:', error);
        errorEl.textContent = '認証に失敗しました。ネットワーク接続を確認してください。';
        errorEl.classList.remove('hidden');
    }
}

// セッション検証（ページロード時）
async function checkAuthentication() {
    const urlParams = new URLSearchParams(window.location.search);
    const store = urlParams.get('store');
    const staff = urlParams.get('staff');
    const apiUrl = localStorage.getItem(SPREADSHEET_API_KEY);

    if (!apiUrl) {
        // API未設定の場合は認証不要
        return true;
    }

    // 管理ページか確認（設定タブへのアクセス）
    const isAdminAccess = !store && !staff;

    // セッショントークンを取得
    const sessionToken = localStorage.getItem(SESSION_TOKEN_KEY);
    const sessionPageType = localStorage.getItem(SESSION_PAGE_TYPE_KEY);

    // セッショントークンがある場合は検証
    if (sessionToken && sessionPageType) {
        const pageType = isAdminAccess ? 'admin' : 'staff';
        if (sessionPageType === pageType) {
            try {
                const response = await fetch(`${apiUrl}?action=verify_session&session_token=${sessionToken}&page_type=${pageType}`);
                const result = await response.json();

                if (result.status === 'success') {
                    // セッション有効
                    return true;
                }
            } catch (error) {
                console.error('セッション検証エラー:', error);
            }
        }

        // セッション無効の場合はクリア
        localStorage.removeItem(SESSION_TOKEN_KEY);
        localStorage.removeItem(SESSION_PAGE_TYPE_KEY);
    }

    // パスワードが必要かチェック（GASから確認）
    try {
        // 空のパスワードで試行して、パスワード必要かどうか確認
        const response = await fetch(`${apiUrl}?action=verify_password&page_type=${isAdminAccess ? 'admin' : 'staff'}&store=${store || ''}&staff=${staff || ''}&password=`);
        const result = await response.json();

        if (result.status === 'success') {
            // パスワード不要（空パスワードで成功）
            localStorage.setItem(SESSION_TOKEN_KEY, result.sessionToken);
            localStorage.setItem(SESSION_PAGE_TYPE_KEY, isAdminAccess ? 'admin' : 'staff');
            return true;
        } else {
            // パスワード必要
            authenticationRequired = true;
            showPasswordDialog(isAdminAccess ? 'admin' : 'staff', store || '', staff || '');
            return false;
        }
    } catch (error) {
        console.error('認証チェックエラー:', error);
        // エラー時は認証不要として処理
        return true;
    }
}

// --- GEMINI API FUNCTIONS ---
const GEMINI_API_KEY_STORAGE = 'mavie_gemini_api_key';

function saveGeminiApiKey() {
    const apiKey = document.getElementById('gemini-api-key').value.trim();
    if (!apiKey) {
        alert('APIキーを入力してください');
        return;
    }
    localStorage.setItem(GEMINI_API_KEY_STORAGE, apiKey);
    // スプレッドシートにも保存
    saveSettingsToSpreadsheet();
    alert('Gemini APIキーを保存しました');
}

function loadGeminiApiKey() {
    const apiKey = localStorage.getItem(GEMINI_API_KEY_STORAGE);
    if (apiKey) {
        const input = document.getElementById('gemini-api-key');
        if (input) input.value = apiKey;
    }
    return apiKey;
}

async function getAIAdvice() {
    const apiKey = loadGeminiApiKey();
    if (!apiKey) {
        alert('先に設定タブでGemini APIキーを登録してください');
        return;
    }

    const btn = document.getElementById('btn-ai-advice');
    const originalText = btn.innerHTML;
    btn.innerHTML = '分析中...';
    btn.disabled = true;

    try {
        // Get current metrics and goals
        const filtered = getFilteredData();
        const metrics = calculateMetrics(filtered);
        const context = getCurrentGoalContext();

        let goalData;
        if (context.type === 'staff') {
            goalData = getStaffGoal(context.store, context.staff);
        } else if (context.type === 'store') {
            goalData = getStoreAggregateGoal(context.store);
        } else {
            goalData = getAllStoresAggregateGoal();
        }

        const salesGoal = goalData.weekdays * goalData.weekdayTarget + goalData.weekends * goalData.weekendTarget;
        const currentSales = metrics.salesTotal;
        const goalAchievementRate = salesGoal > 0 ? ((currentSales / salesGoal) * 100).toFixed(1) : 0;

        // Prepare prompt for Gemini
        const prompt = `あなたはアイラッシュサロンの経営コンサルタントです。以下のデータを分析し、目標達成のための具体的な業務改善アドバイスを日本語で提供してください。

【現在の状況】
- 売上目標: ¥${salesGoal.toLocaleString()}
- 現在の売上: ¥${currentSales.toLocaleString()}
- 達成率: ${goalAchievementRate}%
- 総来店数: ${metrics.customersTotal}名
- 新規客: ${metrics.customersNew}名
- 既存客: ${metrics.customersExisting}名
- 平均客単価: ¥${metrics.customersTotal > 0 ? Math.round(metrics.salesTotal / metrics.customersTotal).toLocaleString() : 0}
- 新規次回予約率: ${metrics.hpbNewCount > 0 ? (((metrics.nextRes.hpbNew + metrics.nextRes.mininaiNew) / metrics.hpbNewCount)*100).toFixed(1) : 0}%

【目標値】
- 新規客目標: ${goalData.newCustomers}名
- 既存客目標: ${goalData.existingCustomers}名
- 目標客単価: ¥${goalData.unitPrice.toLocaleString()}
- 新規次回予約率目標: ${goalData.newReservationRate}%

以下の点について、具体的かつ実行可能なアドバイスを3〜5項目で提供してください：
1. 売上を向上させるための施策
2. 新規・既存客獲得のための具体的なアクション
3. 客単価向上の方法
4. 次回予約率を高めるためのトーク例や施策
5. その他、目標達成に必要な改善点

※アドバイスは箇条書きで、すぐに実行できる具体的な内容にしてください。`;

        // Call Gemini API
        const response = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                contents: [{
                    parts: [{
                        text: prompt
                    }]
                }]
            })
        });

        const data = await response.json();

        if (data.candidates && data.candidates[0] && data.candidates[0].content) {
            const advice = data.candidates[0].content.parts[0].text;
            displayAIAdvice(advice, goalAchievementRate);
        } else {
            throw new Error('APIからの応答が不正です');
        }

    } catch (error) {
        console.error(error);
        alert(`エラーが発生しました: ${error.message}\n\nAPIキーが正しいか確認してください。`);
    } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
    }
}

function displayAIAdvice(advice, achievementRate) {
    const contentDiv = document.getElementById('ai-advice-content');
    const statusColor = achievementRate >= 100 ? 'green' : achievementRate >= 80 ? 'blue' : achievementRate >= 60 ? 'yellow' : 'red';
    const statusText = achievementRate >= 100 ? '目標達成！' : achievementRate >= 80 ? '順調です' : achievementRate >= 60 ? '要注意' : '改善が必要';

    contentDiv.innerHTML = `
        <div class="mb-4 p-4 bg-${statusColor}-50 border-l-4 border-${statusColor}-500 rounded">
            <div class="flex items-center gap-2">
                <span class="text-${statusColor}-800 font-bold">📊 達成率: ${achievementRate}%</span>
                <span class="text-${statusColor}-600">- ${statusText}</span>
            </div>
        </div>
        <div class="prose prose-sm max-w-none">
            <div class="whitespace-pre-wrap text-mavie-800">${advice}</div>
        </div>
        <div class="mt-4 text-xs text-mavie-400 text-right">
            生成日時: ${new Date().toLocaleString('ja-JP')}
        </div>
    `;
}

// --- DEVICE MODE TOGGLE ---
function toggleDeviceMode() {
    const currentMode = localStorage.getItem('deviceMode') || 'mobile';
    const newMode = currentMode === 'mobile' ? 'pc' : 'mobile';

    localStorage.setItem('deviceMode', newMode);
    applyDeviceMode(newMode);
}

function applyDeviceMode(mode) {
    const body = document.body;
    const iconEl = document.getElementById('device-icon-lucide');
    const label = document.getElementById('device-label');

    if (mode === 'mobile') {
        body.classList.remove('pc-mode');
        body.classList.add('mobile-mode');
        if (iconEl) iconEl.setAttribute('data-lucide', 'smartphone');
        label.innerText = 'スマホ';
        enableMobileAccordions();
    } else {
        body.classList.remove('mobile-mode');
        body.classList.add('pc-mode');
        if (iconEl) iconEl.setAttribute('data-lucide', 'monitor');
        label.innerText = 'PC';
        disableMobileAccordions();
    }
    // Re-render Lucide icons
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
}

function enableMobileAccordions() {
    // Wrap sections in accordions for mobile
    const overviewSection = document.getElementById('content-overview');
    const staffDashboardSection = document.getElementById('content-staff-dashboard');

    if (overviewSection && !overviewSection.dataset.accordionEnabled) {
        wrapSectionInAccordion(overviewSection, 'サマリー');
        overviewSection.dataset.accordionEnabled = 'true';
    }

    if (staffDashboardSection && !staffDashboardSection.dataset.accordionEnabled) {
        wrapSectionInAccordion(staffDashboardSection, 'マイダッシュボード');
        staffDashboardSection.dataset.accordionEnabled = 'true';
    }
}

function disableMobileAccordions() {
    // Remove accordion wrappers for PC mode
    document.querySelectorAll('[data-accordion-enabled]').forEach(section => {
        const accordions = section.querySelectorAll('.mobile-accordion-header');
        accordions.forEach(header => {
            const content = header.nextElementSibling;
            if (content && content.classList.contains('mobile-accordion-content')) {
                content.classList.add('open');
                content.style.maxHeight = 'none';
            }
        });
    });
}

function wrapSectionInAccordion(section, title) {
    const children = Array.from(section.children);
    // 逆順で処理してDOM操作の影響を避ける
    for (let index = children.length - 1; index >= 0; index--) {
        const child = children[index];
        if (!child.classList.contains('mobile-accordion-header') && !child.classList.contains('mobile-accordion-content')) {
            const header = document.createElement('div');
            header.className = 'mobile-accordion-header';
            // デフォルトで開いた状態にする（矢印は▲）
            const sectionTitle = child.querySelector('h3')?.textContent || `セクション ${index + 1}`;
            header.innerHTML = `<span>${sectionTitle}</span><span class="accordion-arrow">▲</span>`;

            const content = document.createElement('div');
            // デフォルトで開いた状態にする
            content.className = 'mobile-accordion-content open';

            // 元の要素を移動（クローンではなく）してイベントリスナーを保持
            section.insertBefore(header, child);
            section.insertBefore(content, child);
            content.appendChild(child);

            // クリックハンドラを設定
            header.addEventListener('click', function() {
                content.classList.toggle('open');
                this.querySelector('.accordion-arrow').textContent = content.classList.contains('open') ? '▲' : '▼';
            });
        }
    }
}

function initDeviceMode() {
    // スタッフ専用ページはスマホビューを優先
    if (lockedStaff) {
        applyDeviceMode('mobile');
    } else {
        const savedMode = localStorage.getItem('deviceMode') || 'mobile';
        applyDeviceMode(savedMode);
    }
}

// --- CALENDAR VIEW ---
let currentCalendarDate = new Date();

function initCalendarSelectors() {
    const storeSel = document.getElementById('calendar-store-selector');
    const staffSel = document.getElementById('calendar-staff-selector');
    if (!storeSel || !staffSel) return;

    // スタッフ専用ページの場合、所属店舗のみ表示
    if (lockedStore) {
        const storeName = lockedStore === 'chiba' ? '千葉店' : lockedStore === 'honatsugi' ? '本厚木店' : lockedStore === 'yamato' ? '大和店' : lockedStore;
        storeSel.innerHTML = `<option value="${lockedStore}">${storeName}</option>`;
        storeSel.disabled = true;
        storeSel.parentElement.style.display = 'none';
        handleCalendarStoreChange();
        // スタッフ専用ページ：対象スタッフに固定し、セレクターを非表示
        if (lockedStaff) {
            staffSel.value = lockedStaff;
            staffSel.disabled = true;
            staffSel.parentElement.style.display = 'none';
        }
    } else {
        // 通常ページの場合、全店舗を表示
        storeSel.innerHTML = '<option value="all">全店舗</option>';
        Object.keys(STAFF_ROSTER).forEach(storeId => {
            const storeName = storeId === 'chiba' ? '千葉店' : storeId === 'honatsugi' ? '本厚木店' : storeId === 'yamato' ? '大和店' : storeId;
            const opt = document.createElement('option');
            opt.value = storeId;
            opt.text = storeName;
            storeSel.appendChild(opt);
        });
    }
}

function handleCalendarStoreChange() {
    const storeSel = document.getElementById('calendar-store-selector');
    const staffSel = document.getElementById('calendar-staff-selector');
    if (!storeSel || !staffSel) return;

    const storeValue = storeSel.value;

    // スタッフ専用ページの場合、対象スタッフのみ（全スタッフ選択肢なし）
    if (lockedStaff) {
        staffSel.innerHTML = `<option value="${lockedStaff}">${lockedStaff}</option>`;
        staffSel.value = lockedStaff;
    } else if (storeValue === 'all') {
        staffSel.innerHTML = '<option value="all">全スタッフ</option>';
        // 全店舗の場合、全スタッフを表示
        Object.values(STAFF_ROSTER).flat().forEach(staffName => {
            const opt = document.createElement('option');
            opt.value = staffName;
            opt.text = staffName;
            staffSel.appendChild(opt);
        });
    } else {
        // 特定店舗の場合、その店舗のスタッフのみ表示
        staffSel.innerHTML = '<option value="all">全スタッフ</option>';
        const staffList = STAFF_ROSTER[storeValue] || [];
        staffList.forEach(staffName => {
            const opt = document.createElement('option');
            opt.value = staffName;
            opt.text = staffName;
            staffSel.appendChild(opt);
        });
    }

    renderCalendar();
}

function getCalendarFilteredData() {
    const storeFilter = document.getElementById('calendar-store-selector')?.value || 'all';
    const staffFilter = document.getElementById('calendar-staff-selector')?.value || 'all';
    const safeRawData = Array.isArray(rawData) ? rawData : [];

    return safeRawData.filter(d => {
        if (!d || !d.date) return false;
        if (storeFilter !== 'all' && d.store !== storeFilter) return false;
        // スタッフ名の比較は大文字小文字を区別しない
        if (staffFilter !== 'all' && d.staff?.toLowerCase() !== staffFilter.toLowerCase()) return false;
        return true;
    });
}

function changeCalendarMonth(offset) {
    currentCalendarDate.setMonth(currentCalendarDate.getMonth() + offset);
    renderCalendar();
}

function renderCalendar() {
    const year = currentCalendarDate.getFullYear();
    const month = currentCalendarDate.getMonth();
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    const startingDay = firstDay.getDay();
    const daysInMonth = lastDay.getDate();

    // 月表示を更新
    document.getElementById('calendar-month-display').textContent = `${year}年${month + 1}月`;

    // カレンダーグリッドを生成
    const grid = document.getElementById('calendar-grid');
    grid.innerHTML = '';

    // 空白セル（月初の曜日まで）
    for (let i = 0; i < startingDay; i++) {
        const emptyCell = document.createElement('div');
        emptyCell.className = 'min-h-20 md:min-h-24';
        grid.appendChild(emptyCell);
    }

    // 日別売上データを取得（カレンダー専用フィルター使用）
    const filteredData = getCalendarFilteredData();
    const dailySalesMap = {};

    filteredData.forEach(record => {
        const recordDate = parseDate(record.date);
        if (recordDate.getFullYear() === year && recordDate.getMonth() === month) {
            const day = recordDate.getDate();
            if (!dailySalesMap[day]) {
                dailySalesMap[day] = {
                    sales: 0,
                    customers: 0,
                    newCustomers: 0,
                    existingCustomers: 0,
                    product: 0,
                    hpbPoints: 0,
                    hpbGift: 0,
                    discountOther: 0,
                    discountRefund: 0,
                    nextResNewHPB: 0,
                    nextResNewMiniNai: 0,
                    nextResExisting: 0,
                    reviews5Star: 0
                };
            }
            const salesCash = record.sales?.cash || 0;
            const salesCredit = record.sales?.credit || 0;
            const salesQR = record.sales?.qr || 0;
            const hpbPoints = record.discounts?.hpbPoints || 0;
            const hpbGift = record.discounts?.hpbGift || 0;
            // 売上合計 = 現金 + クレカ + QR + HPBポイント + HPBギフト（他のビューと統一）
            const totalSales = salesCash + salesCredit + salesQR + hpbPoints + hpbGift;
            const productSales = record.sales?.product || 0;
            const newCount = (record.customers?.newHPB || 0) + (record.customers?.newMiniNai || 0);
            const existCount = (record.customers?.existing || 0) + (record.customers?.acquaintance || 0);

            dailySalesMap[day].sales += totalSales;
            dailySalesMap[day].product += productSales;
            dailySalesMap[day].hpbPoints += hpbPoints;
            dailySalesMap[day].hpbGift += hpbGift;
            dailySalesMap[day].discountOther += (record.discounts?.other || 0);
            dailySalesMap[day].discountRefund += (record.discounts?.refund || 0);
            dailySalesMap[day].newCustomers += newCount;
            dailySalesMap[day].existingCustomers += existCount;
            dailySalesMap[day].customers += newCount + existCount;
            dailySalesMap[day].nextResNewHPB += (record.nextRes?.newHPB || 0);
            dailySalesMap[day].nextResNewMiniNai += (record.nextRes?.newMiniNai || 0);
            dailySalesMap[day].nextResExisting += (record.nextRes?.existing || 0);
            dailySalesMap[day].reviews5Star += (record.reviews5Star || 0);
        }
    });

    // 日付セルを生成
    for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement('div');
        const dayData = dailySalesMap[day];
        const salesAmount = dayData?.sales || 0;
        const colorClass = getSalesColorClass(salesAmount);
        const dayOfWeek = new Date(year, month, day).getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

        cell.className = `min-h-20 md:min-h-24 border-2 rounded-lg p-2 cursor-pointer transition-all hover:shadow-md hover:scale-105 ${colorClass}`;
        cell.onclick = () => showDayDetails(year, month, day, dayData);

        const dayNumber = document.createElement('div');
        dayNumber.className = `font-bold text-sm mb-1 ${dayOfWeek === 0 ? 'text-red-500' : dayOfWeek === 6 ? 'text-blue-500' : 'text-mavie-800'}`;
        dayNumber.textContent = day;
        cell.appendChild(dayNumber);

        if (dayData && salesAmount > 0) {
            const salesDiv = document.createElement('div');
            salesDiv.className = 'text-xs text-mavie-700 font-bold';
            salesDiv.textContent = `¥${salesAmount.toLocaleString()}`;
            cell.appendChild(salesDiv);

            const custDiv = document.createElement('div');
            custDiv.className = 'text-xs text-mavie-500 mt-1 hidden md:block';
            custDiv.textContent = `${dayData.customers}名`;
            cell.appendChild(custDiv);
        } else {
            const noData = document.createElement('div');
            noData.className = 'text-xs text-gray-400';
            noData.textContent = '-';
            cell.appendChild(noData);
        }

        grid.appendChild(cell);
    }

    // 月間サマリーを更新
    updateCalendarSummary(dailySalesMap, daysInMonth);
}

function getSalesColorClass(sales) {
    if (!sales || sales === 0) return 'bg-gray-50 border-gray-200';
    if (sales < 50000) return 'bg-blue-50 border-blue-200';
    if (sales < 100000) return 'bg-green-50 border-green-200';
    if (sales < 150000) return 'bg-yellow-50 border-yellow-200';
    if (sales < 200000) return 'bg-orange-50 border-orange-200';
    return 'bg-red-50 border-red-200';
}

function showDayDetails(year, month, day, dayData) {
    const detailsSection = document.getElementById('calendar-day-details');
    const titleEl = document.getElementById('calendar-day-title');
    const statsEl = document.getElementById('calendar-day-stats');

    if (!dayData || dayData.sales === 0) {
        detailsSection.classList.add('hidden');
        return;
    }

    detailsSection.classList.remove('hidden');
    titleEl.textContent = `${month + 1}月${day}日の詳細`;

    const totalNextRes = (dayData.nextResNewHPB || 0) + (dayData.nextResNewMiniNai || 0) + (dayData.nextResExisting || 0);
    const totalLoss = (dayData.discountOther || 0) + (dayData.discountRefund || 0);

    statsEl.innerHTML = `
        <div class="bg-mavie-50 border border-mavie-200 rounded-lg p-3">
            <p class="text-xs text-mavie-500 mb-1">売上合計</p>
            <p class="text-lg font-bold text-mavie-800">¥${dayData.sales.toLocaleString()}</p>
            <p class="text-xs text-mavie-400 mt-1">HPBポイント ¥${(dayData.hpbPoints || 0).toLocaleString()} / ギフト ¥${(dayData.hpbGift || 0).toLocaleString()}</p>
        </div>
        <div class="bg-mavie-50 border border-mavie-200 rounded-lg p-3">
            <p class="text-xs text-mavie-500 mb-1">物販売上</p>
            <p class="text-lg font-bold text-mavie-800">¥${dayData.product.toLocaleString()}</p>
        </div>
        <div class="bg-mavie-50 border border-mavie-200 rounded-lg p-3">
            <p class="text-xs text-mavie-500 mb-1">総来店数</p>
            <p class="text-lg font-bold text-mavie-800">${dayData.customers}名</p>
        </div>
        <div class="bg-blue-50 border border-blue-200 rounded-lg p-3">
            <p class="text-xs text-blue-600 mb-1">新規</p>
            <p class="text-lg font-bold text-blue-800">${dayData.newCustomers}名</p>
        </div>
        <div class="bg-green-50 border border-green-200 rounded-lg p-3">
            <p class="text-xs text-green-600 mb-1">既存</p>
            <p class="text-lg font-bold text-green-800">${dayData.existingCustomers}名</p>
        </div>
        <div class="bg-mavie-100 border border-mavie-300 rounded-lg p-3">
            <p class="text-xs text-mavie-500 mb-1">客単価</p>
            <p class="text-lg font-bold text-mavie-800">¥${dayData.customers > 0 ? Math.round((dayData.sales + dayData.product) / dayData.customers).toLocaleString() : 0}</p>
        </div>
        <div class="bg-purple-50 border border-purple-200 rounded-lg p-3">
            <p class="text-xs text-purple-600 mb-1">次回予約</p>
            <p class="text-lg font-bold text-purple-800">${totalNextRes}件</p>
            <p class="text-xs text-purple-400 mt-1">新規HPB ${dayData.nextResNewHPB || 0} / 新規Mini ${dayData.nextResNewMiniNai || 0} / 既存 ${dayData.nextResExisting || 0}</p>
        </div>
        ${totalLoss > 0 ? `<div class="bg-red-50 border border-red-200 rounded-lg p-3">
            <p class="text-xs text-red-600 mb-1">割引・返金</p>
            <p class="text-lg font-bold text-red-800">¥${totalLoss.toLocaleString()}</p>
            <p class="text-xs text-red-400 mt-1">割引 ¥${(dayData.discountOther || 0).toLocaleString()} / 返金 ¥${(dayData.discountRefund || 0).toLocaleString()}</p>
        </div>` : ''}
        ${(dayData.reviews5Star || 0) > 0 ? `<div class="bg-yellow-50 border border-yellow-200 rounded-lg p-3">
            <p class="text-xs text-yellow-600 mb-1">★5口コミ</p>
            <p class="text-lg font-bold text-yellow-700">${dayData.reviews5Star}件</p>
        </div>` : ''}
    `;

    // スクロールして詳細を表示
    detailsSection.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function updateCalendarSummary(dailySalesMap, daysInMonth) {
    const days = Object.keys(dailySalesMap).filter(d => dailySalesMap[d].sales > 0).length;
    const totalSales = Object.values(dailySalesMap).reduce((sum, d) => sum + d.sales + d.product, 0);
    const avgSales = days > 0 ? Math.round(totalSales / days) : 0;
    const maxSales = Math.max(...Object.values(dailySalesMap).map(d => d.sales + d.product), 0);

    document.getElementById('calendar-summary-days').innerHTML = `${days}<span class="text-sm font-normal ml-1">日</span>`;
    document.getElementById('calendar-summary-sales').textContent = `¥${totalSales.toLocaleString()}`;
    document.getElementById('calendar-summary-avg').textContent = `¥${avgSales.toLocaleString()}`;
    document.getElementById('calendar-summary-max').textContent = `¥${maxSales.toLocaleString()}`;
}

// --- DARK MODE TOGGLE ---
function toggleDarkMode() {
    const html = document.documentElement;
    const isDark = html.classList.contains('dark');

    if (isDark) {
        html.classList.remove('dark');
        localStorage.setItem('darkMode', 'light');
    } else {
        html.classList.add('dark');
        localStorage.setItem('darkMode', 'dark');
    }
    updateDarkModeUI();
    updateChartsForDarkMode();
}

function updateDarkModeUI() {
    const isDark = document.documentElement.classList.contains('dark');
    const iconEl = document.getElementById('dark-mode-icon-lucide');
    const label = document.getElementById('dark-mode-label');
    const btn = document.getElementById('dark-mode-toggle');

    if (isDark) {
        if (iconEl) iconEl.setAttribute('data-lucide', 'sun');
        if (label) label.innerText = 'ライト';
        btn.classList.remove('btn-secondary');
        btn.classList.add('btn-primary');
    } else {
        if (iconEl) iconEl.setAttribute('data-lucide', 'moon');
        if (label) label.innerText = 'ダーク';
        btn.classList.remove('btn-primary');
        btn.classList.add('btn-secondary');
    }
    // Re-render Lucide icons
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }
}

function updateChartsForDarkMode() {
    const isDark = document.documentElement.classList.contains('dark');

    // Chart.js グローバルデフォルトを更新 - 新しいカラーパレット
    if (window.Chart) {
        const textColor = isDark ? '#e8e6e1' : '#334e68';
        const gridColor = isDark ? '#3a3e42' : '#e8e6e1';

        Chart.defaults.color = textColor;
        Chart.defaults.borderColor = gridColor;

        // 既存のチャートを更新
        Object.values(charts).forEach(chart => {
            if (chart && chart.options) {
                // スケールの色を更新
                if (chart.options.scales) {
                    Object.values(chart.options.scales).forEach(scale => {
                        if (scale.ticks) scale.ticks.color = textColor;
                        if (scale.grid) scale.grid.color = gridColor;
                    });
                }
                // 凡例の色を更新
                if (chart.options.plugins && chart.options.plugins.legend && chart.options.plugins.legend.labels) {
                    chart.options.plugins.legend.labels.color = textColor;
                }
                chart.update();
            }
        });
    }
}

function initDarkMode() {
    const savedMode = localStorage.getItem('darkMode');
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;

    if (savedMode === 'dark' || (!savedMode && prefersDark)) {
        document.documentElement.classList.add('dark');
    } else {
        document.documentElement.classList.remove('dark');
    }
    updateDarkModeUI();
}

// ページを離れるときに自動保存
window.addEventListener('beforeunload', (e) => {
    if (autoSaveTimeout) {
        // 保留中の自動保存があれば即座に実行
        clearTimeout(autoSaveTimeout);
        autoSaveToSpreadsheet();
    }
});

document.addEventListener('DOMContentLoaded', async () => {
    // Initialize Lucide icons
    if (typeof lucide !== 'undefined') {
        lucide.createIcons();
    }

    let authenticated = true;
    try {
        // パスワード認証チェック
        authenticated = await checkAuthentication();
    } catch (error) {
        console.error('認証チェックエラー:', error);
        // エラーが発生しても処理を継続（認証スキップ）
        authenticated = true;
    }

    if (!authenticated) {
        // 認証が必要な場合は、ローディングを非表示にしてパスワードダイアログを表示
        hideLoading();
        // ユーザーが正しいパスワードを入力するとページがリロードされる
        return;
    }

    initDarkMode(); // Initialize dark mode
    initDeviceMode(); // Initialize device mode
    loadSpreadsheetApiUrl(); // API URLをinput要素に設定（設定読み込みに必要）
    initCharts();
    await loadData(); // データ読み込みを待機（内部でinitCalendarSelectorsも呼び出し）
    calculateGoal(); // Initialize goal calculation

    // スプレッドシートから設定を読み込み（パスワード、スタッフ名簿、Gemini APIキー等）
    try {
        await loadSettingsFromSpreadsheet();
        await loadStaffPasswordsFromSpreadsheet();
        console.log('設定とパスワードを読み込みました');
        // 設定読み込み後にUIを更新
        updateSettingsList();

        // STAFF_ROSTERが変わった可能性があるため、URLパラメータモードなら再初期化
        const reloadParams = new URLSearchParams(window.location.search);
        if (reloadParams.get('store') || reloadParams.get('staff')) {
            try {
                checkUrlParams();
                initCalendarSelectors();
                updateDashboard();
                console.log('✓ 設定読み込み後にダッシュボードを再初期化しました');
            } catch (e) {
                console.error('設定後の再初期化エラー:', e);
            }
        }
    } catch (e) {
        console.error('設定読み込みエラー:', e);
    }

    // スプレッドシートから目標データを自動読み込み
    try {
        await autoLoadFromSpreadsheet();
    } catch (e) {
        console.error('目標データ読み込みエラー:', e);
    }

    // ダークモードでチャートを初期化した後に更新
    setTimeout(() => updateChartsForDarkMode(), 100);
});
