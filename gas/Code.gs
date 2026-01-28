/**
 * Mavie Dashboard - Google Apps Script
 * スプレッドシートとダッシュボードを連携するためのWeb API
 *
 * シート構成:
 * - フォーム_売上日報: 日々の売上データ
 * - フォーム回答_千葉店: 千葉店の顧客データ
 * - フォーム回答_厚木店: 本厚木店の顧客データ
 * - 目標設定: 月別・スタッフ別の目標データ（自動作成）
 * - 基本給設定: スタッフ別の基本給データ（自動作成）
 *
 * 使い方:
 * 1. このスクリプトをGoogle スプレッドシートのApps Scriptに貼り付け
 * 2. 「デプロイ」→「新しいデプロイ」→「ウェブアプリ」を選択
 * 3. 「アクセスできるユーザー」を「全員」に設定してデプロイ
 * 4. 生成されたURLをダッシュボードの設定に入力
 */

// ==================== 設定 ====================
const SHEET_NAMES = {
  SALES_REPORT: 'フォーム_売上日報',
  CUSTOMER_CHIBA: 'フォーム回答_千葉店',
  CUSTOMER_HONATSUGI: 'フォーム回答_厚木店',
  GOALS: '目標設定',
  SALARIES: '基本給設定'
};

// 売上日報シートのカラム定義（A列から順番に）
const SALES_COLUMNS = {
  TIMESTAMP: 0,      // タイムスタンプ
  DATE: 1,           // 日付
  STORE: 2,          // 店舗
  STAFF: 3,          // スタッフ名
  SALES_CASH: 4,     // 現金売上
  SALES_CREDIT: 5,   // クレジット売上
  SALES_QR: 6,       // QR売上
  SALES_PRODUCT: 7,  // 物販売上
  DISCOUNT_HPB_POINTS: 8,  // HPBポイント値引き
  DISCOUNT_HPB_GIFT: 9,    // HPBギフト値引き
  DISCOUNT_OTHER: 10,      // その他値引き
  DISCOUNT_REFUND: 11,     // 返金
  CUST_NEW_HPB: 12,        // 新規HPB
  CUST_NEW_MININAI: 13,    // 新規ミニナイ
  CUST_REFERRAL: 14,       // 紹介客
  CUST_ACQUAINTANCE: 15,   // 知人
  CUST_EXISTING: 16,       // 既存
  NEXT_RES_NEW_HPB: 17,    // 次回予約_新規HPB
  NEXT_RES_NEW_MININAI: 18, // 次回予約_新規ミニナイ
  NEXT_RES_EXISTING: 19,   // 次回予約_既存
  REVIEWS_5STAR: 20        // 5つ星レビュー
};

// ==================== メイン処理 ====================

/**
 * GETリクエストの処理
 */
function doGet(e) {
  const action = e.parameter.action || 'get_data';

  try {
    let result;

    switch (action) {
      case 'get_data':
        result = getSalesData();
        break;
      case 'get_customers':
        result = getCustomerData();
        break;
      case 'load_goals':
        result = loadGoals();
        break;
      default:
        result = getSalesData();
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

// ==================== 売上データ処理 ====================

/**
 * 売上日報データを取得
 */
function getSalesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!sheet) {
    return { status: 'error', message: `シート「${SHEET_NAMES.SALES_REPORT}」が見つかりません` };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const result = rows.map((row, index) => {
    // 日付のフォーマット処理
    let dateStr = '';
    if (row[SALES_COLUMNS.DATE]) {
      const dateVal = row[SALES_COLUMNS.DATE];
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/M/d');
      } else {
        dateStr = String(dateVal);
      }
    }

    // 店舗名の正規化
    let store = String(row[SALES_COLUMNS.STORE] || '').toLowerCase();
    if (store.includes('千葉') || store.includes('chiba')) {
      store = 'chiba';
    } else if (store.includes('厚木') || store.includes('honatsugi')) {
      store = 'honatsugi';
    }

    return {
      id: index + 1,
      date: dateStr,
      store: store,
      storeName: store === 'chiba' ? '千葉店' : store === 'honatsugi' ? '本厚木店' : store,
      staff: String(row[SALES_COLUMNS.STAFF] || '').toLowerCase(),
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
    };
  }).filter(row => row.date && row.store && row.staff); // 有効なデータのみ

  return result;
}

/**
 * 売上データを更新
 */
function updateSalesData(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SALES_REPORT);

  if (!sheet) {
    return { status: 'error', message: `シート「${SHEET_NAMES.SALES_REPORT}」が見つかりません` };
  }

  rows.forEach(row => {
    const rowIndex = row.id + 1; // 1-indexed, ヘッダー行をスキップ

    // 各フィールドを更新
    sheet.getRange(rowIndex, SALES_COLUMNS.SALES_CASH + 1).setValue(row.sales.cash);
    sheet.getRange(rowIndex, SALES_COLUMNS.SALES_CREDIT + 1).setValue(row.sales.credit);
    sheet.getRange(rowIndex, SALES_COLUMNS.SALES_QR + 1).setValue(row.sales.qr);
    sheet.getRange(rowIndex, SALES_COLUMNS.SALES_PRODUCT + 1).setValue(row.sales.product);
    sheet.getRange(rowIndex, SALES_COLUMNS.CUST_NEW_HPB + 1).setValue(row.customers.newHPB);
    sheet.getRange(rowIndex, SALES_COLUMNS.CUST_NEW_MININAI + 1).setValue(row.customers.newMiniNai);
    sheet.getRange(rowIndex, SALES_COLUMNS.CUST_REFERRAL + 1).setValue(row.customers.referral);
    sheet.getRange(rowIndex, SALES_COLUMNS.CUST_ACQUAINTANCE + 1).setValue(row.customers.acquaintance);
    sheet.getRange(rowIndex, SALES_COLUMNS.CUST_EXISTING + 1).setValue(row.customers.existing);
    sheet.getRange(rowIndex, SALES_COLUMNS.NEXT_RES_NEW_HPB + 1).setValue(row.nextRes.newHPB);
    sheet.getRange(rowIndex, SALES_COLUMNS.NEXT_RES_NEW_MININAI + 1).setValue(row.nextRes.newMiniNai);
    sheet.getRange(rowIndex, SALES_COLUMNS.NEXT_RES_EXISTING + 1).setValue(row.nextRes.existing);
    sheet.getRange(rowIndex, SALES_COLUMNS.REVIEWS_5STAR + 1).setValue(row.reviews5Star || 0);
  });

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
    new Date(), // タイムスタンプ
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

  return { status: 'success', message: 'レコードを追加しました' };
}

// ==================== 顧客データ処理 ====================

/**
 * 顧客データを取得（両店舗）
 */
function getCustomerData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = [];

  // 千葉店のデータ
  const chibaSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_CHIBA);
  if (chibaSheet) {
    const chibaData = parseCustomerSheet(chibaSheet, 'chiba');
    result.push(...chibaData);
  }

  // 本厚木店のデータ
  const honatsugiSheet = ss.getSheetByName(SHEET_NAMES.CUSTOMER_HONATSUGI);
  if (honatsugiSheet) {
    const honatsugiData = parseCustomerSheet(honatsugiSheet, 'honatsugi');
    result.push(...honatsugiData);
  }

  return { status: 'success', data: result };
}

/**
 * 顧客シートをパース
 * フォーム回答_千葉店 / フォーム回答_厚木店 の列構成に対応
 */
function parseCustomerSheet(sheet, store) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // ヘッダーから列インデックスを特定（柔軟な対応）
  const getColumnIndex = (keywords) => {
    return headers.findIndex(h => {
      const headerStr = String(h).toLowerCase();
      return keywords.some(k => headerStr.includes(k.toLowerCase()));
    });
  };

  // 基本情報（実際のフォーム列に対応）
  const colTimestamp = getColumnIndex(['タイムスタンプ', 'timestamp']);
  const colName = getColumnIndex(['お名前', 'フルネーム', '名前', '氏名']);
  const colNameKana = getColumnIndex(['フリガナ', 'ふりがな', 'カナ']);
  const colAddress = getColumnIndex(['住所']);
  const colPhone = getColumnIndex(['電話番号', '携帯電話', '電話']);
  const colBirthday = getColumnIndex(['生年月日']);
  const colJob = getColumnIndex(['職業']);
  const colSnsOk = getColumnIndex(['sns', 'ブログ', '写真']);
  const colVisitReason = getColumnIndex(['ご来店いただいた理由', '来店理由']);
  const colFromOtherSalon = getColumnIndex(['他サロンから']);
  const colDissatisfaction = getColumnIndex(['満足しなかった', '不満']);
  const colAllergy = getColumnIndex(['アレルギー']);

  // 眉毛メニュー関連
  const colEyebrowFreq = getColumnIndex(['眉毛サロンのご利用頻度', '眉毛メニュー】眉毛サロン']);
  const colEyebrowLastCare = getColumnIndex(['眉毛のお手入れ', '最後に眉毛']);
  const colEyebrowConcern = getColumnIndex(['眉毛のお悩み']);
  const colEyebrowDesign = getColumnIndex(['眉毛メニュー】ご希望に一番近いデザイン', '眉毛】ご希望']);
  const colEyebrowImpression = getColumnIndex(['印象に見られたい']);
  const colEyebrowTrouble = getColumnIndex(['眉毛メニュー】施術後の肌トラブル']);

  // まつ毛メニュー関連
  const colLashFreq = getColumnIndex(['まつ毛パーマサロンのご利用頻度', 'まつ毛メニュー】まつ毛パーマ']);
  const colLashDesign = getColumnIndex(['まつ毛メニュー】ご希望のデザイン']);
  const colLashEyeLook = getColumnIndex(['目の見え方']);
  const colLashContact = getColumnIndex(['コンタクトレンズ']);
  const colLashTrouble = getColumnIndex(['まつ毛メニュー】施術後の肌トラブル']);

  // 同意
  const colAgreement = getColumnIndex(['注意事項']);

  return rows.map((row, index) => {
    let dateStr = '';
    let dateObj = null;
    if (colTimestamp >= 0 && row[colTimestamp]) {
      const dateVal = row[colTimestamp];
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, 'Asia/Tokyo', 'yyyy/M/d');
        dateObj = dateVal;
      } else {
        dateStr = String(dateVal);
      }
    }

    return {
      id: `${store}_${index + 1}`,
      store: store,
      storeName: store === 'chiba' ? '千葉店' : '本厚木店',
      date: dateStr,
      timestamp: dateObj ? dateObj.getTime() : 0,

      // 基本情報
      name: colName >= 0 ? String(row[colName] || '') : '',
      nameKana: colNameKana >= 0 ? String(row[colNameKana] || '') : '',
      address: colAddress >= 0 ? String(row[colAddress] || '') : '',
      phone: colPhone >= 0 ? String(row[colPhone] || '') : '',
      birthday: colBirthday >= 0 ? formatDateValue(row[colBirthday]) : '',
      job: colJob >= 0 ? String(row[colJob] || '') : '',
      snsOk: colSnsOk >= 0 ? String(row[colSnsOk] || '') : '',
      visitReason: colVisitReason >= 0 ? String(row[colVisitReason] || '') : '',
      fromOtherSalon: colFromOtherSalon >= 0 ? String(row[colFromOtherSalon] || '') : '',
      dissatisfaction: colDissatisfaction >= 0 ? String(row[colDissatisfaction] || '') : '',
      allergy: colAllergy >= 0 ? String(row[colAllergy] || '') : '',

      // 眉毛メニュー情報
      eyebrowFrequency: colEyebrowFreq >= 0 ? String(row[colEyebrowFreq] || '') : '',
      eyebrowLastCare: colEyebrowLastCare >= 0 ? String(row[colEyebrowLastCare] || '') : '',
      eyebrowConcern: colEyebrowConcern >= 0 ? String(row[colEyebrowConcern] || '') : '',
      eyebrowDesign: colEyebrowDesign >= 0 ? String(row[colEyebrowDesign] || '') : '',
      eyebrowImpression: colEyebrowImpression >= 0 ? String(row[colEyebrowImpression] || '') : '',
      eyebrowTrouble: colEyebrowTrouble >= 0 ? String(row[colEyebrowTrouble] || '') : '',

      // まつ毛メニュー情報
      lashFrequency: colLashFreq >= 0 ? String(row[colLashFreq] || '') : '',
      lashDesign: colLashDesign >= 0 ? String(row[colLashDesign] || '') : '',
      lashEyeLook: colLashEyeLook >= 0 ? String(row[colLashEyeLook] || '') : '',
      lashContact: colLashContact >= 0 ? String(row[colLashContact] || '') : '',
      lashTrouble: colLashTrouble >= 0 ? String(row[colLashTrouble] || '') : '',

      // 同意
      agreement: colAgreement >= 0 ? String(row[colAgreement] || '') : ''
    };
  }).filter(row => row.date || row.name); // 有効なデータのみ
}

/**
 * 日付値をフォーマット
 */
function formatDateValue(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy/M/d');
  }
  return String(val);
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

  return { status: 'success', message: '目標データを保存しました' };
}

/**
 * 目標データを読み込み
 */
function loadGoals() {
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
        } catch (e) {
          console.error('目標データのパースに失敗:', e);
        }
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
        } catch (e) {
          console.error('基本給データのパースに失敗:', e);
        }
      }
    }
  }

  return { status: 'success', goals: goals, salaries: salaries };
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
    'タイムスタンプ',
    '日付',
    '店舗',
    'スタッフ名',
    '現金売上',
    'クレジット売上',
    'QR売上',
    '物販売上',
    'HPBポイント値引き',
    'HPBギフト値引き',
    'その他値引き',
    '返金',
    '新規HPB',
    '新規ミニナイ',
    '紹介客',
    '知人',
    '既存',
    '次回予約_新規HPB',
    '次回予約_新規ミニナイ',
    '次回予約_既存',
    '5つ星レビュー'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  // 列幅を調整
  sheet.setColumnWidth(1, 150); // タイムスタンプ
  sheet.setColumnWidth(2, 100); // 日付
  sheet.setColumnWidth(3, 80);  // 店舗
  sheet.setColumnWidth(4, 100); // スタッフ名

  return { status: 'success', message: '売上日報シートのヘッダーを作成しました' };
}
