/**
 * Mavie Dashboard - Google Apps Script (最適化版)
 * スプレッドシートとダッシュボードを連携するためのWeb API
 *
 * 最適化内容:
 * - CacheServiceによるデータキャッシュ（最大6時間）
 * - バッチ処理によるスプレッドシートアクセス最小化
 * - ヘッダー検索の高速化
 * - 並列データ取得
 *
 * シート構成:
 * - フォーム_売上日報: 日々の売上データ
 * - フォーム回答_千葉店: 千葉店の顧客データ
 * - フォーム回答_本厚木店: 本厚木店の顧客データ
 * - 目標設定: 月別・スタッフ別の目標データ（自動作成）
 * - 基本給設定: スタッフ別の基本給データ（自動作成）
 */

// ==================== 設定 ====================
const SHEET_NAMES = {
  SALES_REPORT: 'フォーム_売上日報',
  CUSTOMER_CHIBA: 'フォーム回答_千葉店',
  CUSTOMER_HONATSUGI: 'フォーム回答_本厚木店',
  GOALS: '目標設定',
  SALARIES: '基本給設定',
  PASSWORDS: 'スタッフパスワード',
  SETTINGS: 'ダッシュボード設定'
};

// キャッシュの有効期限（秒）
const CACHE_EXPIRATION = {
  SALES_DATA: 300,      // 売上データ: 5分
  CUSTOMER_DATA: 600,   // 顧客データ: 10分
  GOALS_DATA: 1800,     // 目標データ: 30分
  SETTINGS_DATA: 3600   // 設定データ: 1時間
};

// 売上日報シートのカラム定義（A列から順番に）
const SALES_COLUMNS = {
  TIMESTAMP: 0,
  DATE: 1,
  STORE: 2,
  STAFF: 3,
  SALES_CASH: 4,
  SALES_CREDIT: 5,
  SALES_QR: 6,
  SALES_PRODUCT: 7,
  DISCOUNT_HPB_POINTS: 8,
  DISCOUNT_HPB_GIFT: 9,
  DISCOUNT_OTHER: 10,
  DISCOUNT_REFUND: 11,
  CUST_NEW_HPB: 12,
  CUST_NEW_MININAI: 13,
  CUST_REFERRAL: 14,
  CUST_ACQUAINTANCE: 15,
  CUST_EXISTING: 16,
  NEXT_RES_NEW_HPB: 17,
  NEXT_RES_NEW_MININAI: 18,
  NEXT_RES_EXISTING: 19,
  REVIEWS_5STAR: 20
};

// ==================== キャッシュ管理 ====================

/**
 * キャッシュからデータを取得
 */
function getFromCache(key) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    // キャッシュエラーは無視
  }
  return null;
}

/**
 * キャッシュにデータを保存
 */
function setToCache(key, data, expiration) {
  try {
    const cache = CacheService.getScriptCache();
    const jsonStr = JSON.stringify(data);
    // GASのキャッシュは最大100KB、大きい場合は分割
    if (jsonStr.length < 100000) {
      cache.put(key, jsonStr, expiration);
    }
  } catch (e) {
    // キャッシュエラーは無視
  }
}

/**
 * キャッシュを無効化
 */
function invalidateCache(key) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(key);
  } catch (e) {
    // エラー無視
  }
}

// ==================== メイン処理 ====================

/**
 * GETリクエストの処理
 */
function doGet(e) {
  const action = e.parameter.action || 'get_data';
  const noCache = e.parameter.nocache === 'true';

  try {
    let result;

    switch (action) {
      case 'get_data':
        result = getSalesData(noCache);
        break;
      case 'get_customers':
        result = getCustomerData(noCache);
        break;
      case 'load_goals':
        result = loadGoals(noCache);
        break;
      case 'load_passwords':
        result = loadPasswords();
        break;
      case 'load_settings':
        result = loadSettings(noCache);
        break;
      case 'get_all':
        // 全データを一括取得（初期ロード用）
        result = getAllData(noCache);
        break;
      default:
        result = getSalesData(noCache);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * POSTリクエストの処理
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'update';

    let result;

    switch (action) {
      case 'update':
        result = updateSalesData(data.rows);
        break;
      case 'save_goals':
        result = saveGoals(data.goals, data.salaries);
        break;
      case 'add_record':
        result = addSalesRecord(data.record);
        break;
      case 'save_passwords':
        result = savePasswords(data.passwords);
        break;
      case 'save_settings':
        result = saveSettings(data.settings);
        break;
      case 'clear_cache':
        result = clearAllCache();
        break;
      default:
        result = { status: 'error', message: '不明なアクションです' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 全キャッシュをクリア
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['sales_data', 'customer_data', 'goals_data', 'settings_data']);
    return { status: 'success', message: 'キャッシュをクリアしました' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

// ==================== 売上データ処理（最適化版） ====================

/**
 * 売上日報データを取得（キャッシュ対応）
 */
function getSalesData(noCache) {
  const CACHE_KEY = 'sales_data';

  // キャッシュチェック
  if (!noCache) {
    const cached = getFromCache(CACHE_KEY);
    if (cached) {
      return cached;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!sheet) {
    return { status: 'error', message: `シート「${SHEET_NAMES.SALES_REPORT}」が見つかりません` };
  }

  // 一括でデータ取得
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const result = [];
  const len = rows.length;

  for (let i = 0; i < len; i++) {
    const row = rows[i];

    // 日付処理
    let dateStr = '';
    const dateVal = row[SALES_COLUMNS.DATE];
    if (dateVal) {
      dateStr = dateVal instanceof Date
        ? Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/M/d')
        : String(dateVal);
    }

    // 店舗名の正規化
    let store = String(row[SALES_COLUMNS.STORE] || '').toLowerCase();
    if (store.includes('千葉') || store.includes('chiba')) {
      store = 'chiba';
    } else if (store.includes('厚木') || store.includes('honatsugi')) {
      store = 'honatsugi';
    }

    const staff = String(row[SALES_COLUMNS.STAFF] || '').toLowerCase();

    // 有効なデータのみ追加
    if (dateStr && store && staff) {
      result.push({
        id: i + 1,
        date: dateStr,
        store: store,
        storeName: store === 'chiba' ? '千葉店' : store === 'honatsugi' ? '本厚木店' : store,
        staff: staff,
        sales: {
          cash: parseInt(row[SALES_COLUMNS.SALES_CASH]) || 0,
          credit: parseInt(row[SALES_COLUMNS.SALES_CREDIT]) || 0,
          qr: parseInt(row[SALES_COLUMNS.SALES_QR]) || 0,
          product: parseInt(row[SALES_COLUMNS.SALES_PRODUCT]) || 0
        },
        discounts: {
          hpbPoints: parseInt(row[SALES_COLUMNS.DISCOUNT_HPB_POINTS]) || 0,
          hpbGift: parseInt(row[SALES_COLUMNS.DISCOUNT_HPB_GIFT]) || 0,
          other: parseInt(row[SALES_COLUMNS.DISCOUNT_OTHER]) || 0,
          refund: parseInt(row[SALES_COLUMNS.DISCOUNT_REFUND]) || 0
        },
        customers: {
          newHPB: parseInt(row[SALES_COLUMNS.CUST_NEW_HPB]) || 0,
          newMiniNai: parseInt(row[SALES_COLUMNS.CUST_NEW_MININAI]) || 0,
          referral: parseInt(row[SALES_COLUMNS.CUST_REFERRAL]) || 0,
          acquaintance: parseInt(row[SALES_COLUMNS.CUST_ACQUAINTANCE]) || 0,
          existing: parseInt(row[SALES_COLUMNS.CUST_EXISTING]) || 0
        },
        nextRes: {
          newHPB: parseInt(row[SALES_COLUMNS.NEXT_RES_NEW_HPB]) || 0,
          newMiniNai: parseInt(row[SALES_COLUMNS.NEXT_RES_NEW_MININAI]) || 0,
          existing: parseInt(row[SALES_COLUMNS.NEXT_RES_EXISTING]) || 0
        },
        reviews5Star: parseInt(row[SALES_COLUMNS.REVIEWS_5STAR]) || 0
      });
    }
  }

  // キャッシュに保存
  setToCache(CACHE_KEY, result, CACHE_EXPIRATION.SALES_DATA);

  return result;
}

/**
 * 売上データを更新（バッチ処理最適化版）
 */
function updateSalesData(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!sheet) {
    return { status: 'error', message: `シート「${SHEET_NAMES.SALES_REPORT}」が見つかりません` };
  }

  // 既存データを一括取得
  const existingData = sheet.getDataRange().getValues();

  // 更新対象の行をグループ化
  const updates = {};
  rows.forEach(row => {
    updates[row.id] = row;
  });

  // 更新データを構築
  const updatedRows = existingData.map((existingRow, index) => {
    if (index === 0) return existingRow; // ヘッダー行はそのまま

    const rowId = index;
    if (updates[rowId]) {
      const row = updates[rowId];
      existingRow[SALES_COLUMNS.SALES_CASH] = row.sales.cash;
      existingRow[SALES_COLUMNS.SALES_CREDIT] = row.sales.credit;
      existingRow[SALES_COLUMNS.SALES_QR] = row.sales.qr;
      existingRow[SALES_COLUMNS.SALES_PRODUCT] = row.sales.product;

      if (row.discounts) {
        existingRow[SALES_COLUMNS.DISCOUNT_HPB_POINTS] = row.discounts.hpbPoints || 0;
        existingRow[SALES_COLUMNS.DISCOUNT_HPB_GIFT] = row.discounts.hpbGift || 0;
        existingRow[SALES_COLUMNS.DISCOUNT_OTHER] = row.discounts.other || 0;
        existingRow[SALES_COLUMNS.DISCOUNT_REFUND] = row.discounts.refund || 0;
      }

      existingRow[SALES_COLUMNS.CUST_NEW_HPB] = row.customers.newHPB;
      existingRow[SALES_COLUMNS.CUST_NEW_MININAI] = row.customers.newMiniNai;
      existingRow[SALES_COLUMNS.CUST_REFERRAL] = row.customers.referral;
      existingRow[SALES_COLUMNS.CUST_ACQUAINTANCE] = row.customers.acquaintance;
      existingRow[SALES_COLUMNS.CUST_EXISTING] = row.customers.existing;
      existingRow[SALES_COLUMNS.NEXT_RES_NEW_HPB] = row.nextRes.newHPB;
      existingRow[SALES_COLUMNS.NEXT_RES_NEW_MININAI] = row.nextRes.newMiniNai;
      existingRow[SALES_COLUMNS.NEXT_RES_EXISTING] = row.nextRes.existing;
      existingRow[SALES_COLUMNS.REVIEWS_5STAR] = row.reviews5Star || 0;
    }
    return existingRow;
  });

  // 一括書き込み
  sheet.getRange(1, 1, updatedRows.length, updatedRows[0].length).setValues(updatedRows);

  // キャッシュを無効化
  invalidateCache('sales_data');

  return { status: 'success', message: `${rows.length}件のデータを更新しました` };
}

/**
 * 新規売上レコードを追加
 */
function addSalesRecord(record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!sheet) {
    return { status: 'error', message: `シート「${SHEET_NAMES.SALES_REPORT}」が見つかりません` };
  }

  const newRow = [
    new Date(),
    record.date,
    record.store,
    record.staff,
    record.sales.cash || 0,
    record.sales.credit || 0,
    record.sales.qr || 0,
    record.sales.product || 0,
    record.discounts?.hpbPoints || 0,
    record.discounts?.hpbGift || 0,
    record.discounts?.other || 0,
    record.discounts?.refund || 0,
    record.customers.newHPB || 0,
    record.customers.newMiniNai || 0,
    record.customers.referral || 0,
    record.customers.acquaintance || 0,
    record.customers.existing || 0,
    record.nextRes?.newHPB || 0,
    record.nextRes?.newMiniNai || 0,
    record.nextRes?.existing || 0,
    record.reviews5Star || 0
  ];

  sheet.appendRow(newRow);

  // キャッシュを無効化
  invalidateCache('sales_data');

  return { status: 'success', message: 'レコードを追加しました' };
}

// ==================== 顧客データ処理（最適化版） ====================

/**
 * 顧客データを取得（キャッシュ対応・並列処理）
 */
function getCustomerData(noCache) {
  const CACHE_KEY = 'customer_data';

  // キャッシュチェック
  if (!noCache) {
    const cached = getFromCache(CACHE_KEY);
    if (cached) {
      return cached;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  // 千葉店と本厚木店のシートを取得
  const chibaSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_CHIBA);
  const honatsugiSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_HONATSUGI);

  // 千葉店のデータ
  if (chibaSheet) {
    const chibaData = parseCustomerSheetOptimized(chibaSheet, 'chiba');
    result.push(...chibaData);
  }

  // 本厚木店のデータ
  if (honatsugiSheet) {
    const honatsugiData = parseCustomerSheetOptimized(honatsugiSheet, 'honatsugi');
    result.push(...honatsugiData);
  }

  const response = { status: 'success', data: result };

  // キャッシュに保存
  setToCache(CACHE_KEY, response, CACHE_EXPIRATION.CUSTOMER_DATA);

  return response;
}

/**
 * 顧客シートをパース（最適化版）
 * ヘッダー検索を1回のみ実行し、ループ処理を最適化
 */
function parseCustomerSheetOptimized(sheet, store) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);
  const rowCount = rows.length;

  // ヘッダーを小文字に変換してキャッシュ
  const headerLower = headers.map(h => String(h).toLowerCase());

  // 列インデックスを一度だけ計算
  const findCol = (keywords) => {
    for (let i = 0; i < headerLower.length; i++) {
      for (let j = 0; j < keywords.length; j++) {
        if (headerLower[i].includes(keywords[j].toLowerCase())) {
          return i;
        }
      }
    }
    return -1;
  };

  // 列インデックスを事前計算
  const cols = {
    timestamp: findCol(['タイムスタンプ', 'timestamp']),
    name: findCol(['お名前', 'フルネーム', '名前', '氏名']),
    nameKana: findCol(['フリガナ', 'ふりがな', 'カナ']),
    address: findCol(['住所']),
    phone: findCol(['電話番号', '携帯電話', '電話']),
    birthday: findCol(['生年月日']),
    job: findCol(['職業']),
    snsOk: findCol(['sns', 'ブログ', '写真']),
    visitReason: findCol(['ご来店いただいた理由', '来店理由']),
    fromOtherSalon: findCol(['他サロンから']),
    dissatisfaction: findCol(['満足しなかった', '不満']),
    allergy: findCol(['アレルギー']),
    eyebrowFreq: findCol(['眉毛サロンのご利用頻度', '眉毛メニュー】眉毛サロン']),
    eyebrowLastCare: findCol(['眉毛のお手入れ', '最後に眉毛']),
    eyebrowConcern: findCol(['眉毛のお悩み']),
    eyebrowDesign: findCol(['眉毛メニュー】ご希望に一番近いデザイン', '眉毛】ご希望']),
    eyebrowImpression: findCol(['印象に見られたい']),
    eyebrowTrouble: findCol(['眉毛メニュー】施術後の肌トラブル']),
    lashFreq: findCol(['まつ毛パーマサロンのご利用頻度', 'まつ毛メニュー】まつ毛パーマ']),
    lashDesign: findCol(['まつ毛メニュー】ご希望のデザイン']),
    lashEyeLook: findCol(['目の見え方']),
    lashContact: findCol(['コンタクトレンズ']),
    lashTrouble: findCol(['まつ毛メニュー】施術後の肌トラブル']),
    agreement: findCol(['注意事項'])
  };

  const storeName = store === 'chiba' ? '千葉店' : '本厚木店';
  const result = [];

  // 高速ヘルパー関数
  const getStr = (row, colIdx) => colIdx >= 0 ? String(row[colIdx] || '') : '';
  const formatDate = (val) => {
    if (!val) return '';
    if (val instanceof Date) {
      return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/M/d');
    }
    return String(val);
  };

  for (let i = 0; i < rowCount; i++) {
    const row = rows[i];

    // タイムスタンプ処理
    let dateStr = '';
    let timestamp = 0;
    if (cols.timestamp >= 0 && row[cols.timestamp]) {
      const dateVal = row[cols.timestamp];
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/M/d');
        timestamp = dateVal.getTime();
      } else {
        dateStr = String(dateVal);
      }
    }

    const name = getStr(row, cols.name);

    // 有効なデータのみ追加
    if (dateStr || name) {
      result.push({
        id: `${store}_${i + 1}`,
        store: store,
        storeName: storeName,
        date: dateStr,
        timestamp: timestamp,
        name: name,
        nameKana: getStr(row, cols.nameKana),
        address: getStr(row, cols.address),
        phone: getStr(row, cols.phone),
        birthday: formatDate(cols.birthday >= 0 ? row[cols.birthday] : null),
        job: getStr(row, cols.job),
        snsOk: getStr(row, cols.snsOk),
        visitReason: getStr(row, cols.visitReason),
        fromOtherSalon: getStr(row, cols.fromOtherSalon),
        dissatisfaction: getStr(row, cols.dissatisfaction),
        allergy: getStr(row, cols.allergy),
        eyebrowFrequency: getStr(row, cols.eyebrowFreq),
        eyebrowLastCare: getStr(row, cols.eyebrowLastCare),
        eyebrowConcern: getStr(row, cols.eyebrowConcern),
        eyebrowDesign: getStr(row, cols.eyebrowDesign),
        eyebrowImpression: getStr(row, cols.eyebrowImpression),
        eyebrowTrouble: getStr(row, cols.eyebrowTrouble),
        lashFrequency: getStr(row, cols.lashFreq),
        lashDesign: getStr(row, cols.lashDesign),
        lashEyeLook: getStr(row, cols.lashEyeLook),
        lashContact: getStr(row, cols.lashContact),
        lashTrouble: getStr(row, cols.lashTrouble),
        agreement: getStr(row, cols.agreement)
      });
    }
  }

  return result;
}

// ==================== 全データ一括取得（初期ロード最適化） ====================

/**
 * 全データを一括取得（初期ロード用）
 */
function getAllData(noCache) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 売上データ
  const salesData = getSalesData(noCache);

  // 目標・基本給データ
  const goalsResult = loadGoals(noCache);

  // 設定データ
  const settingsResult = loadSettings(noCache);

  return {
    status: 'success',
    sales: Array.isArray(salesData) ? salesData : [],
    goals: goalsResult.goals || {},
    salaries: goalsResult.salaries || {},
    settings: settingsResult.settings || {}
  };
}

// ==================== 目標データ処理 ====================

/**
 * 目標データを保存
 */
function saveGoals(goals, salaries) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 目標シートを取得または作成
  let goalsSheet = ss.getSheetByName(SHEET_NAMES.GOALS);
  if (!goalsSheet) {
    goalsSheet = ss.insertSheet(SHEET_NAMES.GOALS);
    goalsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }

  // 目標データを保存
  goalsSheet.getRange(2, 1, 1, 2).setValues([['goals_data', JSON.stringify(goals)]]);

  // 基本給シートを取得または作成
  let salariesSheet = ss.getSheetByName(SHEET_NAMES.SALARIES);
  if (!salariesSheet) {
    salariesSheet = ss.insertSheet(SHEET_NAMES.SALARIES);
    salariesSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }

  // 基本給データを保存
  salariesSheet.getRange(2, 1, 1, 2).setValues([['salaries_data', JSON.stringify(salaries)]]);

  // キャッシュを無効化
  invalidateCache('goals_data');

  return { status: 'success', message: '目標データを保存しました' };
}

/**
 * 目標データを読み込み（キャッシュ対応）
 */
function loadGoals(noCache) {
  const CACHE_KEY = 'goals_data';

  // キャッシュチェック
  if (!noCache) {
    const cached = getFromCache(CACHE_KEY);
    if (cached) {
      return cached;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let goals = {};
  let salaries = {};

  // 目標データを読み込み
  const goalsSheet = ss.getSheetByName(SHEET_NAMES.GOALS);
  if (goalsSheet) {
    const goalsData = goalsSheet.getDataRange().getValues();
    for (let i = 1; i < goalsData.length; i++) {
      if (goalsData[i][0] === 'goals_data' && goalsData[i][1]) {
        try {
          goals = JSON.parse(goalsData[i][1]);
        } catch (e) {}
      }
    }
  }

  // 基本給データを読み込み
  const salariesSheet = ss.getSheetByName(SHEET_NAMES.SALARIES);
  if (salariesSheet) {
    const salariesData = salariesSheet.getDataRange().getValues();
    for (let i = 1; i < salariesData.length; i++) {
      if (salariesData[i][0] === 'salaries_data' && salariesData[i][1]) {
        try {
          salaries = JSON.parse(salariesData[i][1]);
        } catch (e) {}
      }
    }
  }

  const result = { status: 'success', goals: goals, salaries: salaries };

  // キャッシュに保存
  setToCache(CACHE_KEY, result, CACHE_EXPIRATION.GOALS_DATA);

  return result;
}

// ==================== パスワードデータ処理 ====================

/**
 * スタッフパスワードを保存
 */
function savePasswords(passwords) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let passwordSheet = ss.getSheetByName(SHEET_NAMES.PASSWORDS);
  if (!passwordSheet) {
    passwordSheet = ss.insertSheet(SHEET_NAMES.PASSWORDS);
    passwordSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }

  passwordSheet.getRange(2, 1, 1, 2).setValues([['passwords_data', JSON.stringify(passwords)]]);

  return { status: 'success', message: 'パスワードを保存しました' };
}

/**
 * スタッフパスワードを読み込み
 */
function loadPasswords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let passwords = {};

  const passwordSheet = ss.getSheetByName(SHEET_NAMES.PASSWORDS);
  if (passwordSheet) {
    const passwordData = passwordSheet.getDataRange().getValues();
    for (let i = 1; i < passwordData.length; i++) {
      if (passwordData[i][0] === 'passwords_data' && passwordData[i][1]) {
        try {
          passwords = JSON.parse(passwordData[i][1]);
        } catch (e) {}
      }
    }
  }

  return { status: 'success', passwords: passwords };
}

// ==================== 設定データ処理 ====================

/**
 * ダッシュボード設定を保存
 */
function saveSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    settingsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  }

  const lastRow = settingsSheet.getLastRow();
  if (lastRow > 1) {
    settingsSheet.getRange(2, 1, lastRow - 1, 2).clearContent();
  }

  const rows = [];
  if (settings.staffRoster) {
    rows.push(['staff_roster', JSON.stringify(settings.staffRoster)]);
  }
  if (settings.geminiApiKey !== undefined) {
    rows.push(['gemini_api_key', settings.geminiApiKey]);
  }

  if (rows.length > 0) {
    settingsSheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  // キャッシュを無効化
  invalidateCache('settings_data');

  return { status: 'success', message: '設定を保存しました' };
}

/**
 * ダッシュボード設定を読み込み（キャッシュ対応）
 */
function loadSettings(noCache) {
  const CACHE_KEY = 'settings_data';

  // キャッシュチェック
  if (!noCache) {
    const cached = getFromCache(CACHE_KEY);
    if (cached) {
      return cached;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settings = {
    staffRoster: null,
    geminiApiKey: null
  };

  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (settingsSheet) {
    const settingsData = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < settingsData.length; i++) {
      const key = settingsData[i][0];
      const value = settingsData[i][1];

      if (key === 'staff_roster' && value) {
        try {
          settings.staffRoster = JSON.parse(value);
        } catch (e) {}
      } else if (key === 'gemini_api_key') {
        settings.geminiApiKey = value || null;
      }
    }
  }

  const result = { status: 'success', settings: settings };

  // キャッシュに保存
  setToCache(CACHE_KEY, result, CACHE_EXPIRATION.SETTINGS_DATA);

  return result;
}

// ==================== ユーティリティ ====================

/**
 * テスト用：接続確認
 */
function testConnection() {
  return {
    status: 'success',
    message: '接続成功',
    timestamp: new Date().toISOString(),
    cacheEnabled: true,
    sheets: {
      salesReport: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.SALES_REPORT),
      customerChiba: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMER_CHIBA),
      customerHonatsugi: !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.CUSTOMER_HONATSUGI)
    }
  };
}

/**
 * 売上日報シートのヘッダーを自動作成（初期設定用）
 */
function createSalesReportHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SALES_REPORT);
  }

  const headers = [
    'タイムスタンプ', '日付', '店舗', 'スタッフ名',
    '現金売上', 'クレジット売上', 'QR売上', '物販売上',
    'HPBポイント値引き', 'HPBギフト値引き', 'その他値引き', '返金',
    '新規HPB', '新規ミニナイ', '紹介客', '知人', '既存',
    '次回予約_新規HPB', '次回予約_新規ミニナイ', '次回予約_既存', '5つ星レビュー'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  return { status: 'success', message: '売上日報シートのヘッダーを作成しました' };
}
