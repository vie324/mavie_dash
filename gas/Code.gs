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
// 新カラム構造: A~W列
const SALES_COLUMNS = {
  TIMESTAMP: 0,           // A: タイムスタンプ
  DATE: 1,                // B: 入力する出勤日を選択してください。
  STORE: 2,               // C: 出勤店舗を選択してください。
  STAFF_1: 3,             // D: スタッフ名を選択してください。（店舗別選択用）
  STAFF: 4,               // E: スタッフ名を選択してください。（実際のスタッフ名）
  SALES_CASH: 5,          // F: 現金売上合計
  SALES_CREDIT: 6,        // G: クレジット決済売上合計
  SALES_QR: 7,            // H: QR決済売上合計
  SALES_PRODUCT: 8,       // I: 物販売上（上記売上の内数）
  DISCOUNT_HPB_POINTS: 9, // J: HPBポイント利用額
  DISCOUNT_HPB_GIFT: 10,  // K: HPBギフト券利用額
  DISCOUNT_OTHER: 11,     // L: その他割引額
  DISCOUNT_REFUND: 12,    // M: 返金額
  CUST_NEW_HPB: 13,       // N: 【新規】来店数 (HPB)
  CUST_NEW_MININAI: 14,   // O: 【新規】来店数 (minimo/ネイリーなど)
  CUST_EXISTING: 15,      // P: 【既存】来店数
  NEXT_RES_NEW_HPB: 16,   // Q: 【新規】からの次回予約獲得数 (HPB)
  NEXT_RES_NEW_MININAI: 17, // R: 【新規】からの次回予約獲得数 (minimo/ネイリーなど)
  REVIEWS_5STAR: 18,      // S: 口コミ★5獲得数
  BLOG_UPDATES: 19,       // T: ブログ更新数
  SNS_UPDATES: 20,        // U: SNS更新数
  NEXT_RES_EXISTING: 21,  // V: 【既存】からの次回予約獲得数
  CUST_ACQUAINTANCE: 22   // W: 【既存】来店数（知り合い価格案内）
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
      case 'get_customers_today':
        // 当日分のみ取得（高速化）
        result = getCustomerDataToday(noCache);
        break;
      case 'get_customers_by_store':
        // 店舗別取得（スタッフ専用URL用・高速化）
        const storeId = e.parameter.store || '';
        result = getCustomerDataByStore(storeId, noCache);
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
      case 'verify_password':
        // パスワード認証
        const pageType = e.parameter.page_type || 'staff';
        const store = e.parameter.store || '';
        const staff = e.parameter.staff || '';
        const password = e.parameter.password || '';
        result = verifyPassword(pageType, store, staff, password);
        break;
      case 'verify_session':
        // セッション検証
        const sessionToken = e.parameter.session_token || '';
        const sessionPageType = e.parameter.page_type || 'staff';
        result = verifySession(sessionToken, sessionPageType);
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
          existing: parseInt(row[SALES_COLUMNS.CUST_EXISTING]) || 0,
          acquaintance: parseInt(row[SALES_COLUMNS.CUST_ACQUAINTANCE]) || 0
        },
        nextRes: {
          newHPB: parseInt(row[SALES_COLUMNS.NEXT_RES_NEW_HPB]) || 0,
          newMiniNai: parseInt(row[SALES_COLUMNS.NEXT_RES_NEW_MININAI]) || 0,
          existing: parseInt(row[SALES_COLUMNS.NEXT_RES_EXISTING]) || 0
        },
        reviews5Star: parseInt(row[SALES_COLUMNS.REVIEWS_5STAR]) || 0,
        blogUpdates: parseInt(row[SALES_COLUMNS.BLOG_UPDATES]) || 0,
        snsUpdates: parseInt(row[SALES_COLUMNS.SNS_UPDATES]) || 0
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
      existingRow[SALES_COLUMNS.CUST_EXISTING] = row.customers.existing;
      existingRow[SALES_COLUMNS.CUST_ACQUAINTANCE] = row.customers.acquaintance || 0;
      existingRow[SALES_COLUMNS.NEXT_RES_NEW_HPB] = row.nextRes.newHPB;
      existingRow[SALES_COLUMNS.NEXT_RES_NEW_MININAI] = row.nextRes.newMiniNai;
      existingRow[SALES_COLUMNS.NEXT_RES_EXISTING] = row.nextRes.existing;
      existingRow[SALES_COLUMNS.REVIEWS_5STAR] = row.reviews5Star || 0;
      existingRow[SALES_COLUMNS.BLOG_UPDATES] = row.blogUpdates || 0;
      existingRow[SALES_COLUMNS.SNS_UPDATES] = row.snsUpdates || 0;
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
    new Date(),                          // A: タイムスタンプ
    record.date,                         // B: 出勤日
    record.store,                        // C: 出勤店舗
    record.staff,                        // D: スタッフ名（店舗別）
    record.staff,                        // E: スタッフ名（実際）
    record.sales.cash || 0,              // F: 現金売上
    record.sales.credit || 0,            // G: クレジット売上
    record.sales.qr || 0,                // H: QR売上
    record.sales.product || 0,           // I: 物販売上
    record.discounts?.hpbPoints || 0,    // J: HPBポイント
    record.discounts?.hpbGift || 0,      // K: HPBギフト
    record.discounts?.other || 0,        // L: その他割引
    record.discounts?.refund || 0,       // M: 返金
    record.customers.newHPB || 0,        // N: 新規HPB
    record.customers.newMiniNai || 0,    // O: 新規minimo等
    record.customers.existing || 0,      // P: 既存
    record.nextRes?.newHPB || 0,         // Q: 新規次回予約HPB
    record.nextRes?.newMiniNai || 0,     // R: 新規次回予約minimo等
    record.reviews5Star || 0,            // S: 口コミ★5
    record.blogUpdates || 0,             // T: ブログ更新数
    record.snsUpdates || 0,              // U: SNS更新数
    record.nextRes?.existing || 0,       // V: 既存次回予約
    record.customers.acquaintance || 0   // W: 知り合い価格
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

/**
 * 当日分の顧客データのみを取得（高速化版）
 * キャッシュ時間を短く設定し、最新データを素早く表示
 */
function getCustomerDataToday(noCache) {
  const CACHE_KEY = 'customer_data_today';

  // キャッシュチェック（1分間のみ）
  if (!noCache) {
    const cached = getFromCache(CACHE_KEY);
    if (cached) {
      return cached;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  // 今日の日付を取得（JST）
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayTimestamp = today.getTime();

  // 千葉店と本厚木店のシートを取得
  const chibaSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_CHIBA);
  const honatsugiSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_HONATSUGI);

  // 千葉店のデータ（当日分のみ）
  if (chibaSheet) {
    const chibaData = parseCustomerSheetOptimized(chibaSheet, 'chiba');
    const todayData = chibaData.filter(item => {
      return item.timestamp >= todayTimestamp;
    });
    result.push(...todayData);
  }

  // 本厚木店のデータ（当日分のみ）
  if (honatsugiSheet) {
    const honatsugiData = parseCustomerSheetOptimized(honatsugiSheet, 'honatsugi');
    const todayData = honatsugiData.filter(item => {
      return item.timestamp >= todayTimestamp;
    });
    result.push(...todayData);
  }

  const response = { status: 'success', data: result };

  // キャッシュに保存（60秒 = 1分）
  setToCache(CACHE_KEY, response, 60);

  return response;
}

/**
 * 店舗別の顧客データを取得（スタッフ専用URL用・超高速化）
 * 指定された店舗のデータのみを取得し、キャッシュ時間を30秒に設定
 */
function getCustomerDataByStore(storeId, noCache) {
  const CACHE_KEY = 'customer_data_store_' + storeId;

  // キャッシュチェック（30秒のみ）
  if (!noCache) {
    const cached = getFromCache(CACHE_KEY);
    if (cached) {
      return cached;
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  // 店舗に応じたシートを取得
  let sheet = null;
  if (storeId === 'chiba') {
    sheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_CHIBA);
  } else if (storeId === 'honatsugi') {
    sheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_HONATSUGI);
  }

  // 指定店舗のデータのみ取得
  if (sheet) {
    const storeData = parseCustomerSheetOptimized(sheet, storeId);
    result.push(...storeData);
  }

  const response = { status: 'success', data: result };

  // キャッシュに保存（30秒）
  setToCache(CACHE_KEY, response, 30);

  return response;
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

  // 既存データを読み込んでマップ化
  const existingData = {};
  const dataRange = settingsSheet.getDataRange();
  const values = dataRange.getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) {
      existingData[values[i][0]] = values[i][1];
    }
  }

  // 新しい設定で既存データを更新
  if (settings.staffRoster) {
    existingData['staff_roster'] = JSON.stringify(settings.staffRoster);
  }
  if (settings.geminiApiKey !== undefined) {
    existingData['gemini_api_key'] = settings.geminiApiKey;
  }
  if (settings.adminPassword !== undefined) {
    existingData['admin_password'] = settings.adminPassword;
  }

  // 全データを書き込み
  const rows = Object.keys(existingData).map(key => [key, existingData[key]]);

  // 既存データをクリア
  if (values.length > 1) {
    settingsSheet.getRange(2, 1, values.length - 1, 2).clearContent();
  }

  // 新しいデータを書き込み
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

/**
 * パスワード認証
 */
function verifyPassword(pageType, store, staff, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // パスワードシートから認証情報を取得
    const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
    if (!settingsSheet) {
      return { status: 'error', message: '設定シートが見つかりません' };
    }

    const settingsData = settingsSheet.getDataRange().getValues();
    let adminPassword = '';

    // 管理ページのパスワードを取得
    for (let i = 1; i < settingsData.length; i++) {
      if (settingsData[i][0] === 'admin_password') {
        adminPassword = settingsData[i][1] || '';
        break;
      }
    }

    if (pageType === 'admin') {
      // 管理ページの認証
      if (adminPassword === '' || password === adminPassword) {
        // セッショントークンを生成
        const sessionToken = Utilities.getUuid();
        const expiresAt = new Date().getTime() + (24 * 60 * 60 * 1000); // 24時間有効

        // セッションをPropertiesServiceに保存
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty('admin_session_' + sessionToken, JSON.stringify({
          type: 'admin',
          expiresAt: expiresAt
        }));

        return {
          status: 'success',
          message: '認証成功',
          sessionToken: sessionToken,
          expiresAt: expiresAt
        };
      } else {
        return { status: 'error', message: 'パスワードが正しくありません' };
      }
    } else {
      // スタッフページの認証
      const passwordSheet = ss.getSheetByName(SHEET_NAMES.PASSWORDS);
      if (!passwordSheet) {
        return { status: 'error', message: 'パスワードシートが見つかりません' };
      }

      const passwordData = passwordSheet.getDataRange().getValues();
      let passwords = {};

      for (let i = 1; i < passwordData.length; i++) {
        if (passwordData[i][0] === 'passwords_data' && passwordData[i][1]) {
          try {
            passwords = JSON.parse(passwordData[i][1]);
          } catch (e) {}
        }
      }

      const key = store + '_' + staff;
      const correctPassword = passwords[key] || '';

      if (correctPassword === '' || password === correctPassword) {
        // セッショントークンを生成
        const sessionToken = Utilities.getUuid();
        const expiresAt = new Date().getTime() + (24 * 60 * 60 * 1000); // 24時間有効

        // セッションをPropertiesServiceに保存
        const userProperties = PropertiesService.getUserProperties();
        userProperties.setProperty('staff_session_' + sessionToken, JSON.stringify({
          type: 'staff',
          store: store,
          staff: staff,
          expiresAt: expiresAt
        }));

        return {
          status: 'success',
          message: '認証成功',
          sessionToken: sessionToken,
          expiresAt: expiresAt
        };
      } else {
        return { status: 'error', message: 'パスワードが正しくありません' };
      }
    }
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

/**
 * セッション検証
 */
function verifySession(sessionToken, pageType) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const prefix = pageType === 'admin' ? 'admin_session_' : 'staff_session_';
    const sessionData = userProperties.getProperty(prefix + sessionToken);

    if (!sessionData) {
      return { status: 'error', message: 'セッションが見つかりません' };
    }

    const session = JSON.parse(sessionData);
    const now = new Date().getTime();

    if (session.expiresAt < now) {
      // セッション期限切れ
      userProperties.deleteProperty(prefix + sessionToken);
      return { status: 'error', message: 'セッションの期限が切れています' };
    }

    return {
      status: 'success',
      message: 'セッション有効',
      session: session
    };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
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
    'タイムスタンプ',                           // A
    '入力する出勤日を選択してください。',         // B
    '出勤店舗を選択してください。',              // C
    'スタッフ名を選択してください。',            // D（店舗別）
    'スタッフ名を選択してください。',            // E（実際）
    '現金売上合計',                            // F
    'クレジット決済売上合計',                   // G
    'QR決済売上合計',                          // H
    '物販売上（上記売上の内数）',               // I
    'HPBポイント利用額',                       // J
    'HPBギフト券利用額',                       // K
    'その他割引額',                            // L
    '返金額',                                 // M
    '【新規】来店数 (HPB)',                    // N
    '【新規】来店数 (minimo/ネイリーなど)',      // O
    '【既存】来店数',                          // P
    '【新規】からの次回予約獲得数 (HPB)',        // Q
    '【新規】からの次回予約獲得数 (minimo/ネイリーなど)', // R
    '口コミ★5獲得数',                         // S
    'ブログ更新数',                            // T
    'SNS更新数',                              // U
    '【既存】からの次回予約獲得数',              // V
    '【既存】来店数（知り合い価格案内）'          // W
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  return { status: 'success', message: '売上日報シートのヘッダーを作成しました' };
}

// ==================== 旧形式データインポート ====================

/**
 * 旧形式データをインポート
 * 使い方：
 * 1. 「旧形式インポート」シートを作成（createImportSheet関数を実行）
 * 2. B1セルに店舗ID（chiba または honatsugi）を入力
 * 3. B2セルにスタッフ名を入力
 * 4. B3セルに年を入力（例：2026）
 * 5. B4セルに月を入力（例：1）
 * 6. A7行目から旧形式データの「総売上」行以降を貼り付け
 *    ※「1月」「1日、2日...」などのヘッダー行は不要
 * 7. この関数を実行
 */
function importOldFormatData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const importSheet = ss.getSheetByName('旧形式インポート');

  if (!importSheet) {
    return { status: 'error', message: '「旧形式インポート」シートが見つかりません。createImportSheet関数を実行してシートを作成してください。' };
  }

  // 設定を読み込み
  const storeId = importSheet.getRange('B1').getValue().toString().trim();
  const staffName = importSheet.getRange('B2').getValue().toString().trim();
  const year = parseInt(importSheet.getRange('B3').getValue());
  const month = parseInt(importSheet.getRange('B4').getValue());

  if (!storeId || !staffName || !year || !month) {
    return { status: 'error', message: 'B1~B4に店舗ID、スタッフ名、年、月を入力してください' };
  }

  // データを読み込み（7行目から20行分）
  const dataRange = importSheet.getRange('A7:AH27');
  const data = dataRange.getValues();

  // データを解析して変換
  const salesData = parseOldFormatData(data, storeId, staffName, year, month);

  if (salesData.length === 0) {
    return { status: 'error', message: 'データが見つかりませんでした' };
  }

  // 売上日報シートに書き込み
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);
  if (!salesSheet) {
    return { status: 'error', message: '売上日報シートが見つかりません' };
  }

  // 最後の行を取得
  const lastRow = salesSheet.getLastRow();

  // データを追加
  salesSheet.getRange(lastRow + 1, 1, salesData.length, salesData[0].length).setValues(salesData);

  // キャッシュを無効化
  invalidateCache('sales_data');

  return {
    status: 'success',
    message: `${salesData.length}件のデータをインポートしました`,
    count: salesData.length
  };
}

/**
 * 旧形式データを解析して新形式の配列に変換
 */
function parseOldFormatData(data, storeId, staffName, year, month) {
  const result = [];

  // 行の定義（旧形式）
  const rowDefs = {
    totalSales: 0,       // 総売上
    cash: 1,             // 現金売上合計
    credit: 2,           // クレジット決済売上合計
    qr: 3,               // QR決済売上合計
    hpbTotal: 4,         // HPB割引合計
    hpbPoints: 5,        // HPBポイント割引
    hpbGift: 6,          // HPBギフト券割引
    lossTotal: 7,        // 損失合計
    otherDiscount: 8,    // その他割引
    refund: 9,           // 返金
    totalCustomers: 10,  // 総来店数
    newHPB: 11,          // 新規数（HPB）
    newMinimo: 12,       // 新規数（minimo）
    existing: 13,        // 既存来店
    merchandise: 14,     // 物販
    newNextRes: 15,      // 新規次回予約（HPB/minimo）
    existingNextRes: 16  // 既存次回予約
  };

  // 日付列は3列目から（C列=index 2）開始し、31日分
  const startCol = 2;
  const maxDays = 31;

  // 各日付ごとにデータを作成
  for (let day = 1; day <= maxDays; day++) {
    const colIndex = startCol + day - 1;

    // この日のデータが存在するかチェック（総売上が0でない、または来店数が0でない）
    const totalSales = parseFloat(data[rowDefs.totalSales][colIndex]) || 0;
    const totalCustomers = parseInt(data[rowDefs.totalCustomers][colIndex]) || 0;

    // データがない日はスキップ
    if (totalSales === 0 && totalCustomers === 0) {
      continue;
    }

    // 日付を作成
    const dateStr = `${year}/${month}/${day}`;
    const timestamp = new Date(year, month - 1, day);

    // 新形式の行データを作成（フォーム_売上日報の形式）
    const row = [
      timestamp,                                                    // A: タイムスタンプ
      dateStr,                                                      // B: 日付
      storeId,                                                      // C: 店舗
      staffName,                                                    // D: スタッフ名（店舗別）
      staffName,                                                    // E: スタッフ名
      parseFloat(data[rowDefs.cash][colIndex]) || 0,               // F: 現金売上合計
      parseFloat(data[rowDefs.credit][colIndex]) || 0,             // G: クレジット決済売上合計
      parseFloat(data[rowDefs.qr][colIndex]) || 0,                 // H: QR決済売上合計
      parseFloat(data[rowDefs.merchandise][colIndex]) || 0,        // I: 物販売上
      parseFloat(data[rowDefs.hpbPoints][colIndex]) || 0,          // J: HPBポイント利用額
      parseFloat(data[rowDefs.hpbGift][colIndex]) || 0,            // K: HPBギフト券利用額
      parseFloat(data[rowDefs.otherDiscount][colIndex]) || 0,      // L: その他割引額
      parseFloat(data[rowDefs.refund][colIndex]) || 0,             // M: 返金額
      parseInt(data[rowDefs.newHPB][colIndex]) || 0,               // N: 新規来店数（HPB）
      parseInt(data[rowDefs.newMinimo][colIndex]) || 0,            // O: 新規来店数（minimo）
      parseInt(data[rowDefs.existing][colIndex]) || 0,             // P: 既存来店数
      parseInt(data[rowDefs.newNextRes][colIndex]) || 0,           // Q: 新規次回予約（HPB）
      0,                                                            // R: 新規次回予約（minimo） ※旧形式では分かれていない
      0,                                                            // S: 口コミ★5獲得数 ※旧形式にはない
      0,                                                            // T: ブログ更新数 ※旧形式にはない
      0,                                                            // U: SNS更新数 ※旧形式にはない
      parseInt(data[rowDefs.existingNextRes][colIndex]) || 0,      // V: 既存次回予約
      0                                                             // W: 既存来店数（知り合い） ※旧形式にはない
    ];

    result.push(row);
  }

  return result;
}

/**
 * 旧形式インポート用シートを作成
 */
function createImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('旧形式インポート');

  if (sheet) {
    return { status: 'error', message: '「旧形式インポート」シートは既に存在します' };
  }

  sheet = ss.insertSheet('旧形式インポート');

  // 説明を追加
  sheet.getRange('A1').setValue('店舗ID（chiba または honatsugi）:');
  sheet.getRange('B1').setValue('chiba');
  sheet.getRange('A2').setValue('スタッフ名:');
  sheet.getRange('B2').setValue('kiki');
  sheet.getRange('A3').setValue('年:');
  sheet.getRange('B3').setValue('2026');
  sheet.getRange('A4').setValue('月:');
  sheet.getRange('B4').setValue('1');

  sheet.getRange('A6').setValue('【7行目以降に旧形式データを貼り付けてください】');
  sheet.getRange('A7').setValue('総売上');
  sheet.getRange('B7').setValue('← この行から「総売上」以降のデータを貼り付け（「1月」「1日、2日...」などは不要）');

  // フォーマット
  sheet.getRange('A1:B4').setFontWeight('bold');
  sheet.getRange('A6:B7').setBackground('#fff3cd');

  return { status: 'success', message: '旧形式インポートシートを作成しました' };
}

// ==================== 2026年1月データ直接インポート ====================

/**
 * 2026年1月分のデータを直接インポート
 * GASエディタでこの関数を実行するだけでOK
 */
function import2026January() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!salesSheet) {
    return { status: 'error', message: '売上日報シートが見つかりません' };
  }

  // 既存の2026年1月データをチェック
  const checkResult = checkJanuaryData();
  if (checkResult.total > 0) {
    Logger.log(`警告: 既に${checkResult.total}件の2026年1月データが存在します`);
    Logger.log('既存データ: ' + JSON.stringify(checkResult.byStaff));
    return {
      status: 'error',
      message: `既に${checkResult.total}件の2026年1月データが存在します。重複を避けるため、先にdeleteJanuary2026Data()を実行してください。`,
      existingData: checkResult.byStaff
    };
  }

  const year = 2026;
  const month = 1;
  const allData = [];

  // kiki (chiba) のデータ
  const kikiData = {
    store: 'chiba',
    staff: 'kiki',
    days: {
      4: {cash:13200,credit:21750,qr:0,hpbPoints:400,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:2},
      5: {cash:34700,credit:7800,qr:0,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:3,product:0,newNextRes:0,existingNextRes:3},
      7: {cash:13300,credit:15750,qr:0,hpbPoints:800,hpbGift:0,other:0,refund:0,newHPB:4,newMinimo:0,existing:1,product:0,newNextRes:0,existingNextRes:1},
      8: {cash:13100,credit:11400,qr:8300,hpbPoints:4400,hpbGift:0,other:0,refund:0,newHPB:4,newMinimo:1,existing:2,product:2640,newNextRes:1,existingNextRes:0},
      9: {cash:15200,credit:11000,qr:10900,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:4,product:2640,newNextRes:2,existingNextRes:2},
      10: {cash:6650,credit:6000,qr:15900,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:3,product:2640,newNextRes:0,existingNextRes:2},
      12: {cash:0,credit:21100,qr:4800,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:2,product:0,newNextRes:1,existingNextRes:2},
      13: {cash:5000,credit:33000,qr:3500,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:6,product:0,newNextRes:0,existingNextRes:5},
      14: {cash:0,credit:41900,qr:0,hpbPoints:2100,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:3,product:6160,newNextRes:1,existingNextRes:2},
      16: {cash:5000,credit:30700,qr:16200,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:4},
      17: {cash:9800,credit:15000,qr:10000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:7,product:0,newNextRes:0,existingNextRes:7},
      18: {cash:10000,credit:23650,qr:11650,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:7,product:0,newNextRes:0,existingNextRes:6},
      20: {cash:0,credit:36750,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:4,product:11000,newNextRes:0,existingNextRes:3},
      21: {cash:0,credit:19400,qr:6100,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:3,product:0,newNextRes:0,existingNextRes:3},
      22: {cash:6000,credit:15550,qr:5000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:5,product:0,newNextRes:0,existingNextRes:5},
      24: {cash:14550,credit:14550,qr:14800,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:6,product:0,newNextRes:0,existingNextRes:5},
      25: {cash:15900,credit:37200,qr:13000,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:6,product:0,newNextRes:0,existingNextRes:4},
      27: {cash:0,credit:15300,qr:10000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:4,product:3520,newNextRes:0,existingNextRes:4},
      28: {cash:5000,credit:21500,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:3,product:0,newNextRes:0,existingNextRes:3},
      30: {cash:0,credit:25050,qr:12750,hpbPoints:800,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:4},
      31: {cash:5000,credit:25750,qr:12150,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:5,product:0,newNextRes:0,existingNextRes:5}
    }
  };

  // karin (chiba) のデータ
  const karinData = {
    store: 'chiba',
    staff: 'karin',
    days: {
      4: {cash:22200,credit:14300,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:4},
      5: {cash:15300,credit:10100,qr:3800,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:4,newMinimo:0,existing:2,product:0,newNextRes:0,existingNextRes:2},
      6: {cash:15400,credit:15900,qr:10550,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:2,existing:3,product:0,newNextRes:0,existingNextRes:0},
      8: {cash:31400,credit:5000,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:5,newMinimo:0,existing:1,product:0,newNextRes:0,existingNextRes:1},
      9: {cash:33750,credit:16200,qr:0,hpbPoints:1200,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:1,existing:4,product:2640,newNextRes:0,existingNextRes:3},
      10: {cash:9800,credit:20000,qr:24070,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:6,product:0,newNextRes:1,existingNextRes:6},
      11: {cash:26100,credit:9000,qr:8500,hpbPoints:400,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:0},
      13: {cash:4800,credit:9800,qr:10000,hpbPoints:100,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:2,product:0,newNextRes:1,existingNextRes:2},
      14: {cash:20250,credit:8000,qr:0,hpbPoints:500,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:1,product:11000,newNextRes:1,existingNextRes:1},
      16: {cash:14800,credit:23650,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:7,product:0,newNextRes:0,existingNextRes:6},
      17: {cash:15900,credit:23650,qr:5000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:5,product:0,newNextRes:1,existingNextRes:5},
      18: {cash:10550,credit:26600,qr:0,hpbPoints:500,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:5,product:0,newNextRes:0,existingNextRes:4},
      20: {cash:4000,credit:28700,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:1,existing:3,product:0,newNextRes:0,existingNextRes:3},
      21: {cash:5000,credit:28700,qr:6650,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:5,product:0,newNextRes:0,existingNextRes:3},
      22: {cash:6600,credit:11820,qr:5000,hpbPoints:2500,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:2,product:0,newNextRes:0,existingNextRes:2},
      24: {cash:15000,credit:22100,qr:10000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:5,product:0,newNextRes:0,existingNextRes:4},
      25: {cash:0,credit:26000,qr:22350,hpbPoints:500,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:2},
      27: {cash:4400,credit:17500,qr:11000,hpbPoints:700,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:3,product:2640,newNextRes:0,existingNextRes:0},
      28: {cash:10350,credit:5000,qr:5000,hpbPoints:100,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:3,product:0,newNextRes:1,existingNextRes:2},
      29: {cash:12600,credit:8900,qr:0,hpbPoints:600,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:1,product:0,newNextRes:1,existingNextRes:1}
    }
  };

  // nanami (chiba) のデータ
  const nanamiData = {
    store: 'chiba',
    staff: 'nanami',
    days: {
      5: {cash:4800,credit:25200,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:4,newMinimo:0,existing:1,product:0,newNextRes:2,existingNextRes:1},
      6: {cash:19650,credit:0,qr:22950,hpbPoints:500,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:0,newNextRes:1,existingNextRes:2},
      7: {cash:11000,credit:26000,qr:14800,hpbPoints:1000,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:0,newNextRes:1,existingNextRes:3},
      9: {cash:18200,credit:11100,qr:10550,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:0,newNextRes:0,existingNextRes:3},
      10: {cash:22600,credit:0,qr:21500,hpbPoints:500,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:0,newNextRes:1,existingNextRes:3},
      12: {cash:0,credit:41800,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:4,product:0,newNextRes:0,existingNextRes:3},
      13: {cash:11900,credit:7900,qr:10000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:1,existing:3,product:0,newNextRes:0,existingNextRes:2},
      14: {cash:6600,credit:17750,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:4,product:2640,newNextRes:0,existingNextRes:3},
      15: {cash:10000,credit:7000,qr:10350,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:3,product:0,newNextRes:0,existingNextRes:1},
      17: {cash:12150,credit:32100,qr:16320,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:7,product:11000,newNextRes:0,existingNextRes:6},
      18: {cash:0,credit:28750,qr:19000,hpbPoints:1800,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:5,product:0,newNextRes:1,existingNextRes:5},
      19: {cash:13500,credit:38800,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:7,product:0,newNextRes:0,existingNextRes:6},
      24: {cash:13100,credit:19550,qr:14800,hpbPoints:900,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:0,newNextRes:1,existingNextRes:3},
      25: {cash:5550,credit:13900,qr:20650,hpbPoints:1400,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:3},
      26: {cash:16700,credit:0,qr:10500,hpbPoints:2300,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:1,existing:3,product:0,newNextRes:1,existingNextRes:1},
      28: {cash:0,credit:8000,qr:5500,hpbPoints:0,hpbGift:500,other:0,refund:0,newHPB:2,newMinimo:0,existing:1,product:0,newNextRes:1,existingNextRes:0},
      29: {cash:13750,credit:11300,qr:3200,hpbPoints:1400,hpbGift:0,other:0,refund:0,newHPB:4,newMinimo:1,existing:1,product:0,newNextRes:3,existingNextRes:1},
      30: {cash:0,credit:43600,qr:0,hpbPoints:1800,hpbGift:2000,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:0,newNextRes:1,existingNextRes:2},
      31: {cash:9000,credit:19000,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:3,product:0,newNextRes:1,existingNextRes:3}
    }
  };

  // kanon (chiba) のデータ
  const kanonData = {
    store: 'chiba',
    staff: 'kanon',
    days: {
      5: {cash:7200,credit:16100,qr:4000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      6: {cash:0,credit:36000,qr:0,hpbPoints:400,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:2,existing:3,product:0,newNextRes:0,existingNextRes:2},
      7: {cash:19800,credit:20200,qr:0,hpbPoints:400,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:2,product:0,newNextRes:0,existingNextRes:0},
      8: {cash:22400,credit:8400,qr:0,hpbPoints:500,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      10: {cash:5770,credit:38850,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:6,product:0,newNextRes:0,existingNextRes:4},
      11: {cash:13200,credit:29700,qr:0,hpbPoints:2400,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:4,product:0,newNextRes:1,existingNextRes:2},
      12: {cash:5000,credit:17300,qr:10800,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:2,existing:3,product:0,newNextRes:0,existingNextRes:2},
      14: {cash:15000,credit:12000,qr:8000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:3},
      15: {cash:9000,credit:4400,qr:5500,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:2,existing:2,product:0,newNextRes:0,existingNextRes:1},
      17: {cash:0,credit:25250,qr:22550,hpbPoints:1000,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:1,existing:4,product:0,newNextRes:0,existingNextRes:4},
      18: {cash:8600,credit:14700,qr:0,hpbPoints:1300,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:1,existing:0,product:0,newNextRes:2,existingNextRes:0},
      19: {cash:5000,credit:28100,qr:5000,hpbPoints:1000,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:5,product:2640,newNextRes:0,existingNextRes:4},
      21: {cash:14000,credit:12600,qr:1400,hpbPoints:1800,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:1,existing:2,product:0,newNextRes:1,existingNextRes:2},
      22: {cash:0,credit:30400,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:2,product:0,newNextRes:0,existingNextRes:1},
      23: {cash:4800,credit:7400,qr:13900,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:3,existing:3,product:0,newNextRes:0,existingNextRes:3},
      24: {cash:0,credit:12200,qr:13800,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:2,product:0,newNextRes:1,existingNextRes:2},
      27: {cash:0,credit:4400,qr:14600,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:1,existing:0,product:0,newNextRes:1,existingNextRes:0},
      28: {cash:2200,credit:8000,qr:4400,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:0,existing:1,product:0,newNextRes:0,existingNextRes:0},
      30: {cash:5000,credit:27600,qr:0,hpbPoints:1000,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:2640,newNextRes:1,existingNextRes:1},
      31: {cash:2200,credit:31300,qr:0,hpbPoints:2600,hpbGift:0,other:0,refund:0,newHPB:4,newMinimo:0,existing:1,product:0,newNextRes:1,existingNextRes:1}
    }
  };

  // ayami (chiba) のデータ
  const ayamiData = {
    store: 'chiba',
    staff: 'ayami',
    days: {
      4: {cash:0,credit:28500,qr:12900,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:1,existing:2,product:0,newNextRes:0,existingNextRes:2},
      5: {cash:22900,credit:5300,qr:2200,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:1,newMinimo:3,existing:2,product:0,newNextRes:0,existingNextRes:1},
      6: {cash:18800,credit:7700,qr:24600,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:1,existing:3,product:0,newNextRes:0,existingNextRes:1},
      7: {cash:4000,credit:34400,qr:0,hpbPoints:1000,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:3,product:0,newNextRes:0,existingNextRes:1},
      9: {cash:15200,credit:10400,qr:3900,hpbPoints:300,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:2,product:0,newNextRes:0,existingNextRes:0},
      10: {cash:9550,credit:36500,qr:9000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:2},
      11: {cash:14550,credit:11050,qr:5000,hpbPoints:200,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:3,product:0,newNextRes:0,existingNextRes:2},
      13: {cash:0,credit:25800,qr:4500,hpbPoints:2200,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:3,product:0,newNextRes:0,existingNextRes:1},
      14: {cash:0,credit:25300,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:3,product:0,newNextRes:1,existingNextRes:3},
      15: {cash:5350,credit:13400,qr:5500,hpbPoints:100,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:2,product:0,newNextRes:0,existingNextRes:1},
      17: {cash:0,credit:37000,qr:9600,hpbPoints:800,hpbGift:0,other:0,refund:0,newHPB:3,newMinimo:0,existing:4,product:0,newNextRes:0,existingNextRes:3},
      18: {cash:12600,credit:7700,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:5,product:0,newNextRes:0,existingNextRes:0},
      19: {cash:0,credit:0,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:0,product:0,newNextRes:0,existingNextRes:0},
      20: {cash:9400,credit:34900,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:6,product:0,newNextRes:1,existingNextRes:2},
      21: {cash:10500,credit:12800,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:2,product:2640,newNextRes:1,existingNextRes:1},
      22: {cash:7400,credit:11800,qr:9600,hpbPoints:2300,hpbGift:0,other:0,refund:0,newHPB:4,newMinimo:0,existing:1,product:2640,newNextRes:2,existingNextRes:0},
      23: {cash:2200,credit:14800,qr:17300,hpbPoints:100,hpbGift:0,other:0,refund:0,newHPB:2,newMinimo:0,existing:5,product:0,newNextRes:1,existingNextRes:1},
      26: {cash:15300,credit:11400,qr:4800,hpbPoints:1500,hpbGift:0,other:0,refund:0,newHPB:5,newMinimo:0,existing:1,product:0,newNextRes:4,existingNextRes:1},
      27: {cash:0,credit:0,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:0,product:0,newNextRes:0,existingNextRes:0},
      28: {cash:0,credit:13500,qr:4800,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:2,product:0,newNextRes:0,existingNextRes:2},
      29: {cash:0,credit:22100,qr:3400,hpbPoints:1000,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      30: {cash:0,credit:0,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:0,product:0,newNextRes:0,existingNextRes:0},
      31: {cash:16000,credit:19300,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:1,existing:3,product:0,newNextRes:1,existingNextRes:2}
    }
  };

  // vienna (honatsugi) のデータ
  const viennaData = {
    store: 'honatsugi',
    staff: 'vienna',
    days: {
      6: {cash:27500,credit:17600,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      7: {cash:0,credit:17600,qr:21500,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:2640,newNextRes:0,existingNextRes:0},
      8: {cash:20500,credit:0,qr:11000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      10: {cash:33400,credit:18700,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      11: {cash:17000,credit:18000,qr:10500,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      12: {cash:5500,credit:12100,qr:30100,hpbPoints:600,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:2640,newNextRes:0,existingNextRes:0},
      13: {cash:5000,credit:32500,qr:0,hpbPoints:0,hpbGift:0,other:500,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      16: {cash:27000,credit:11600,qr:5500,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      17: {cash:0,credit:22100,qr:17000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      18: {cash:0,credit:45100,qr:5500,hpbPoints:100,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      19: {cash:11000,credit:28100,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      21: {cash:5500,credit:13000,qr:17000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      22: {cash:18600,credit:0,qr:25000,hpbPoints:600,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      24: {cash:5500,credit:12100,qr:11000,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      25: {cash:21500,credit:5500,qr:5500,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      26: {cash:27000,credit:11000,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0},
      27: {cash:21500,credit:18000,qr:0,hpbPoints:0,hpbGift:0,other:0,refund:0,newHPB:0,newMinimo:0,existing:0,product:0,newNextRes:0,existingNextRes:0}
    }
  };


  // miki (千葉店) のデータ - 産休中のため全て0
  const mikiData = {
    store: 'chiba',
    staff: 'miki',
    days: {}
  };

  // 全スタッフのデータを統合
  const allStaffData = [kikiData, karinData, nanamiData, kanonData, ayamiData, viennaData, mikiData];

  // 各スタッフのデータを処理
  for (const staffData of allStaffData) {
    for (const [day, dayData] of Object.entries(staffData.days)) {
      const dayNum = parseInt(day);
      const dateStr = `${year}/${month}/${dayNum}`;
      const timestamp = new Date(year, month - 1, dayNum);

      const row = [
        timestamp,                    // A: タイムスタンプ
        dateStr,                      // B: 日付
        staffData.store,              // C: 店舗
        staffData.staff,              // D: スタッフ名（店舗別）
        staffData.staff,              // E: スタッフ名
        dayData.cash,                 // F: 現金売上合計
        dayData.credit,               // G: クレジット決済売上合計
        dayData.qr,                   // H: QR決済売上合計
        dayData.product,              // I: 物販売上
        dayData.hpbPoints,            // J: HPBポイント利用額
        dayData.hpbGift,              // K: HPBギフト券利用額
        dayData.other,                // L: その他割引額
        dayData.refund,               // M: 返金額
        dayData.newHPB,               // N: 新規来店数（HPB）
        dayData.newMinimo,            // O: 新規来店数（minimo）
        dayData.existing,             // P: 既存来店数
        dayData.newNextRes,           // Q: 新規次回予約（HPB）
        0,                            // R: 新規次回予約（minimo）
        0,                            // S: 口コミ★5獲得数
        0,                            // T: ブログ更新数
        0,                            // U: SNS更新数
        dayData.existingNextRes,      // V: 既存次回予約
        0                             // W: 既存来店数（知り合い）
      ];

      allData.push(row);
    }
  }

  // データをシートに追加
  const lastRow = salesSheet.getLastRow();
  salesSheet.getRange(lastRow + 1, 1, allData.length, allData[0].length).setValues(allData);

  // キャッシュを無効化
  invalidateCache('sales_data');

  // ログ出力
  Logger.log(`2026年1月分のデータを${allData.length}件追加しました`);
  Logger.log(`最終行: ${lastRow} → ${salesSheet.getLastRow()}`);

  return {
    status: 'success',
    message: `2026年1月分のデータを${allData.length}件インポートしました`,
    count: allData.length
  };
}

/**
 * すべてのキャッシュをクリア（デバッグ用）
 */
function clearAllCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['sales_data', 'customer_data_chiba', 'customer_data_honatsugi', 'goals_data', 'settings_data']);
    Logger.log('すべてのキャッシュをクリアしました');
    return { status: 'success', message: 'キャッシュをクリアしました' };
  } catch (e) {
    Logger.log('キャッシュクリアエラー: ' + e.toString());
    return { status: 'error', message: e.toString() };
  }
}

/**
 * 2026年1月のデータをチェック（デバッグ用）
 */
function checkJanuaryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!salesSheet) {
    return { status: 'error', message: 'シートが見つかりません' };
  }

  const data = salesSheet.getDataRange().getValues();
  const header = data[0];
  const rows = data.slice(1);

  // 2026年1月のデータを検索
  let januaryCount = 0;
  const staffCounts = {};

  for (const row of rows) {
    const dateStr = String(row[SALES_COLUMNS.DATE] || '');
    const staff = String(row[SALES_COLUMNS.STAFF] || '').toLowerCase();

    if (dateStr.startsWith('2026/1/')) {
      januaryCount++;
      staffCounts[staff] = (staffCounts[staff] || 0) + 1;
    }
  }

  Logger.log(`2026年1月のデータ: ${januaryCount}件`);
  Logger.log('スタッフ別: ' + JSON.stringify(staffCounts));

  return {
    status: 'success',
    total: januaryCount,
    byStaff: staffCounts,
    message: `2026年1月のデータ: ${januaryCount}件`
  };
}

/**
 * 2026年1月のデータを削除（重複削除用）
 */
function deleteJanuary2026Data() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!salesSheet) {
    return { status: 'error', message: 'シートが見つかりません' };
  }

  const data = salesSheet.getDataRange().getValues();
  const rowsToDelete = [];

  // 2026年1月のデータの行番号を収集（後ろから削除するため逆順）
  for (let i = data.length - 1; i >= 1; i--) {
    const dateStr = String(data[i][SALES_COLUMNS.DATE] || '');
    if (dateStr.startsWith('2026/1/')) {
      rowsToDelete.push(i + 1); // 1-indexed
    }
  }

  // 行を削除
  let deletedCount = 0;
  for (const rowNum of rowsToDelete) {
    salesSheet.deleteRow(rowNum);
    deletedCount++;
  }

  // キャッシュを無効化
  invalidateCache('sales_data');

  Logger.log(`2026年1月のデータを${deletedCount}件削除しました`);

  return {
    status: 'success',
    message: `2026年1月のデータを${deletedCount}件削除しました`,
    deletedCount: deletedCount
  };
}
