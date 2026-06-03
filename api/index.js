// ============================================================
// Mavie Dashboard 互換API（Vercel Serverless Function）
// 現行 GAS(gas/Code.gs) と同じ ?action= 契約を Supabase で再現。
// フロントは接続先URLを差し替えるだけで動作します。
//   GET  /api?action=get_data | get_customers | load_goals | ...
//   POST /api    body(JSON): { action: 'update' | 'save_goals' | ... }
// ============================================================
import {
  supabase, cors, json, readBody,
  salesRowToApi, salesApiToCols, customerRowToApi,
  getConfig, setConfig, signToken, verifyTokenStr
} from './_lib.js';

export default async function handler(req, res) {
  cors(res);
  if (req.method === 'OPTIONS') { res.status(204).end(); return; }

  try {
    if (req.method === 'GET') return await handleGet(req, res);
    if (req.method === 'POST') return await handlePost(req, res);
    return json(res, { status: 'error', message: 'Method not allowed' }, 405);
  } catch (err) {
    return json(res, { status: 'error', message: String(err && err.message || err) }, 500);
  }
}

// ==================== GET ====================
async function handleGet(req, res) {
  const q = req.query || {};
  const action = q.action || 'get_data';

  switch (action) {
    case 'health':
      return json(res, { status: 'ok', timestamp: new Date().toISOString(), version: 'supabase-1.0' });

    case 'get_data': {
      let query = supabase.from('sales_reports').select('*').order('report_date', { ascending: true });
      if (q.start_date && q.end_date) query = query.gte('report_date', q.start_date).lte('report_date', q.end_date);
      else if (q.months) {
        const d = new Date(); d.setMonth(d.getMonth() - parseInt(q.months, 10));
        query = query.gte('report_date', d.toISOString().slice(0, 10));
      }
      const { data, error } = await query;
      if (error) throw error;
      return json(res, (data || []).map(salesRowToApi)); // ← 素の配列（GAS と同じ）
    }

    case 'get_customers':
    case 'get_customers_today':
    case 'get_customers_by_store': {
      let query = supabase.from('customers').select('*').order('submitted_at', { ascending: false });
      if (action === 'get_customers_by_store' && q.store) query = query.eq('store', q.store);
      if (action === 'get_customers_today') {
        const start = new Date(); start.setHours(0, 0, 0, 0);
        query = query.gte('submitted_at', start.toISOString());
      }
      const { data, error } = await query;
      if (error) throw error;
      return json(res, { status: 'success', data: (data || []).map(customerRowToApi) });
    }

    case 'load_goals': {
      const m = await getConfig(['goals_data', 'salaries_data']);
      return json(res, { status: 'success', goals: m.goals_data || {}, salaries: m.salaries_data || {} });
    }

    case 'load_settings': {
      const m = await getConfig(['staff_roster', 'gemini_api_key']);
      return json(res, { status: 'success', settings: { staffRoster: m.staff_roster || null, geminiApiKey: m.gemini_api_key || null } });
    }

    case 'load_passwords': {
      const m = await getConfig(['passwords_data']);
      return json(res, { status: 'success', passwords: m.passwords_data || {} });
    }

    case 'verify_password': {
      const pageType = q.page_type || 'staff';
      const password = q.password || '';
      const m = await getConfig(['admin_password', 'passwords_data']);
      let correct = '';
      if (pageType === 'admin') {
        correct = m.admin_password || '';
      } else {
        const pwds = m.passwords_data || {};
        correct = pwds[`${q.store || ''}_${q.staff || ''}`] || '';
      }
      if (correct === '' || password === correct) {
        const payload = pageType === 'admin'
          ? { type: 'admin' }
          : { type: 'staff', store: q.store || '', staff: q.staff || '' };
        const sessionToken = signToken(payload);
        return json(res, { status: 'success', message: '認証成功', sessionToken, expiresAt: Date.now() + 24 * 60 * 60 * 1000 });
      }
      return json(res, { status: 'error', message: 'パスワードが正しくありません' });
    }

    case 'verify_session': {
      const payload = verifyTokenStr(q.session_token || '');
      const want = q.page_type === 'admin' ? 'admin' : 'staff';
      if (payload && payload.type === want) return json(res, { status: 'success', message: 'セッション有効', session: payload });
      return json(res, { status: 'error', message: 'セッションが無効です' });
    }

    case 'get_all': {
      const [{ data: sales }, cfg] = await Promise.all([
        supabase.from('sales_reports').select('*').order('report_date', { ascending: true }),
        getConfig(['goals_data', 'salaries_data', 'staff_roster', 'gemini_api_key'])
      ]);
      return json(res, {
        status: 'success',
        sales: (sales || []).map(salesRowToApi),
        goals: cfg.goals_data || {}, salaries: cfg.salaries_data || {},
        settings: { staffRoster: cfg.staff_roster || null, geminiApiKey: cfg.gemini_api_key || null }
      });
    }

    default:
      return json(res, { status: 'error', message: `不明なaction: ${action}` }, 400);
  }
}

// ==================== POST ====================
async function handlePost(req, res) {
  const body = await readBody(req);
  const action = body.action || 'update';

  switch (action) {
    case 'update': {
      const rows = body.rows || [];
      for (const row of rows) {
        const { error } = await supabase.from('sales_reports').update(salesApiToCols(row)).eq('id', row.id);
        if (error) throw error;
      }
      return json(res, { status: 'success', message: `${rows.length}件のデータを更新しました` });
    }

    case 'add_record': {
      const r = body.record || {};
      const s = r.sales || {}, d = r.discounts || {}, c = r.customers || {}, n = r.nextRes || {};
      const { error } = await supabase.from('sales_reports').insert({
        report_date: r.date, store: r.store, staff: r.staff,
        sales_cash: s.cash || 0, sales_credit: s.credit || 0, sales_qr: s.qr || 0, sales_product: s.product || 0,
        discount_hpb_points: d.hpbPoints || 0, discount_hpb_gift: d.hpbGift || 0, discount_other: d.other || 0, discount_refund: d.refund || 0,
        cust_new_hpb: c.newHPB || 0, cust_new_mininai: c.newMiniNai || 0, cust_existing: c.existing || 0, cust_acquaintance: c.acquaintance || 0,
        nextres_new_hpb: n.newHPB || 0, nextres_new_mininai: n.newMiniNai || 0, nextres_existing: n.existing || 0,
        reviews_5star: r.reviews5Star || 0, blog_updates: r.blogUpdates || 0, sns_updates: r.snsUpdates || 0
      });
      if (error) throw error;
      return json(res, { status: 'success', message: 'レコードを追加しました' });
    }

    case 'save_goals': {
      await setConfig('goals_data', body.goals || {});
      await setConfig('salaries_data', body.salaries || {});
      return json(res, { status: 'success', message: '目標データを保存しました' });
    }

    case 'save_settings': {
      const s = body.settings || {};
      if (s.staffRoster !== undefined) await setConfig('staff_roster', s.staffRoster);
      if (s.geminiApiKey !== undefined) await setConfig('gemini_api_key', s.geminiApiKey);
      if (s.adminPassword !== undefined) await setConfig('admin_password', s.adminPassword);
      return json(res, { status: 'success', message: '設定を保存しました' });
    }

    case 'save_passwords': {
      await setConfig('passwords_data', body.passwords || {});
      return json(res, { status: 'success', message: 'パスワードを保存しました' });
    }

    case 'clear_cache':
      return json(res, { status: 'success', message: 'キャッシュ不要（PostgreSQL直結）' });

    default:
      return json(res, { status: 'error', message: `不明なaction: ${action}` }, 400);
  }
}
