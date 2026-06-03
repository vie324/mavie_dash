/**
 * フォーム → Supabase 同期（「フォームは残してデータだけ移行」用）
 *
 * 既存の Google フォーム回答をそのまま使い続けながら、回答が追加された瞬間に
 * Supabase にも1行コピーします。スタッフの入力操作は一切変わりません。
 *
 * 【セットアップ】
 *  1. 売上フォーム／各顧客フォームに紐づくスプレッドシートの Apps Script に本ファイルを貼り付け
 *  2. setSupabaseKey() を一度実行し、service_role キーをスクリプトプロパティに保存
 *     （コードに直書きしない）
 *  3. SUPABASE_URL を自分のプロジェクトURLに変更
 *  4. トリガーを設定：
 *      - 売上シート側：onSalesFormSubmit を「フォーム送信時」installable トリガーで登録
 *      - 顧客シート側：onCustomerFormSubmit を「フォーム送信時」installable トリガーで登録
 *  ※ 既存の onFormSubmit（あれば）と競合しないよう関数名を分けています。
 */

const SUPABASE_URL = 'https://xxxx.supabase.co'; // ← 自分のプロジェクトURLに変更

// service_role キーをスクリプトプロパティから取得
function _supaKey() {
  return PropertiesService.getScriptProperties().getProperty('SUPABASE_SERVICE_ROLE_KEY');
}
// 初回のみ実行してキーを保存（実行後はこの行のキーは消してOK）
function setSupabaseKey() {
  PropertiesService.getScriptProperties()
    .setProperty('SUPABASE_SERVICE_ROLE_KEY', 'ここに service_role キーを貼り付けて1回だけ実行');
}

// Supabase REST へ INSERT
function _supaInsert(table, payload) {
  const key = _supaKey();
  const res = UrlFetchApp.fetch(`${SUPABASE_URL}/rest/v1/${table}`, {
    method: 'post',
    contentType: 'application/json',
    headers: { apikey: key, Authorization: 'Bearer ' + key, Prefer: 'return=minimal' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code >= 300) Logger.log(`Supabase INSERT 失敗(${code}): ${res.getContentText()}`);
  return code;
}

// 店舗名 → ID 正規化
function _normStore(v) {
  const s = String(v || '').toLowerCase();
  if (s.includes('千葉') || s.includes('chiba')) return 'chiba';
  if (s.includes('厚木') || s.includes('honatsugi')) return 'honatsugi';
  if (s.includes('大和') || s.includes('yamato')) return 'yamato';
  return s;
}
function _toISODate(v) {            // 'yyyy/M/d' や Date → 'YYYY-MM-DD'
  if (!v) return null;
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  const s = String(v).replace(/-/g, '/');
  const p = s.split('/');
  if (p.length < 3) return null;
  return `${p[0]}-${String(p[1]).padStart(2, '0')}-${String(p[2]).padStart(2, '0')}`;
}
function _num(v) { const n = parseInt(v, 10); return isNaN(n) ? 0 : n; }

/**
 * 売上日報フォーム送信時（フォーム_売上日報 と同じ列順を前提）
 * 列順は gas/Code.gs の SALES_COLUMNS と一致。
 */
function onSalesFormSubmit(e) {
  const v = e.values; // A列から順の配列
  const store = _normStore(v[2]);
  // D=本厚木 / E=千葉 / F=大和。店舗に応じた列を採用（空ならフォールバック）
  let staff = '';
  const sH = String(v[3] || '').toLowerCase().trim();
  const sC = String(v[4] || '').toLowerCase().trim();
  const sY = String(v[5] || '').toLowerCase().trim();
  if (store === 'honatsugi') staff = sH; else if (store === 'chiba') staff = sC; else if (store === 'yamato') staff = sY;
  if (!staff) staff = sH || sC || sY;

  const payload = {
    report_date: _toISODate(v[1]), store: store, staff: staff,
    sales_cash: _num(v[6]), sales_credit: _num(v[7]), sales_qr: _num(v[8]), sales_product: _num(v[9]),
    discount_hpb_points: _num(v[10]), discount_hpb_gift: _num(v[11]), discount_other: _num(v[12]), discount_refund: _num(v[13]),
    cust_new_hpb: _num(v[14]), cust_new_mininai: _num(v[15]), cust_existing: _num(v[16]),
    nextres_new_hpb: _num(v[17]), nextres_new_mininai: _num(v[18]),
    reviews_5star: _num(v[19]), blog_updates: _num(v[20]), sns_updates: _num(v[21]),
    nextres_existing: _num(v[22]), cust_acquaintance: _num(v[23])
  };
  if (!payload.report_date || !payload.store || !payload.staff) {
    Logger.log('売上同期スキップ（必須項目不足）: ' + JSON.stringify(payload));
    return;
  }
  _supaInsert('sales_reports', payload);
}

/**
 * 顧客カウンセリングフォーム送信時
 * 店舗はシート名（フォーム回答_千葉店 等）から判定。
 * 項目は namedValues をキーワードで自動検出（gas/Code.gs の findCol と同じ方針）。
 */
function onCustomerFormSubmit(e) {
  const sheetName = e.range ? e.range.getSheet().getName() : '';
  const store = _normStore(sheetName);
  const nv = e.namedValues || {};

  const find = (keywords) => {
    for (const k of Object.keys(nv)) {
      const kl = k.toLowerCase();
      for (const kw of keywords) if (kl.includes(kw.toLowerCase())) return (nv[k] || [''])[0];
    }
    return '';
  };

  const payload = {
    store: store,
    submitted_at: new Date().toISOString(),
    name: find(['お名前', 'フルネーム', '名前', '氏名']),
    name_kana: find(['フリガナ', 'ふりがな', 'カナ']),
    address: find(['住所']),
    phone: find(['電話番号', '携帯電話', '電話']),
    birthday: _toISODate(find(['生年月日'])),
    job: find(['職業']),
    sns_ok: find(['sns', 'ブログ', '写真']),
    visit_reason: find(['ご来店いただいた理由', '来店理由']),
    from_other_salon: find(['他サロンから']),
    dissatisfaction: find(['満足しなかった', '不満']),
    allergy: find(['アレルギー']),
    eyebrow_frequency: find(['眉毛サロンのご利用頻度', '眉毛メニュー】眉毛サロン']),
    eyebrow_last_care: find(['眉毛のお手入れ', '最後に眉毛']),
    eyebrow_concern: find(['眉毛のお悩み']),
    eyebrow_design: find(['眉毛メニュー】ご希望に一番近いデザイン', '眉毛】ご希望']),
    eyebrow_design_image: find(['眉毛メニュー】ご希望のデザインイメージ', '眉毛】デザインイメージ']),
    eyebrow_impression: find(['印象に見られたい']),
    eyebrow_trouble: find(['眉毛メニュー】施術後の肌トラブル']),
    lash_frequency: find(['まつ毛パーマサロンのご利用頻度', 'まつ毛メニュー】まつ毛パーマ']),
    lash_design: find(['まつ毛メニュー】ご希望のデザイン']),
    lash_design_image: find(['まつ毛メニュー】ご希望のデザインイメージ', 'まつ毛】デザインイメージ']),
    lash_eye_look: find(['目の見え方']),
    lash_contact: find(['コンタクトレンズ']),
    lash_trouble: find(['まつ毛メニュー】施術後の肌トラブル']),
    agreement: find(['注意事項'])
  };
  if (!payload.store) { Logger.log('顧客同期スキップ（店舗不明）: ' + sheetName); return; }
  _supaInsert('customers', payload);
}
