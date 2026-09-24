// /api/cashbook — 現金出納帳（店舗ごと・月単位）と SalonOne の現金売上との突合
//
// GET  ?shop=<id>&month=YYYY-MM[&log=1]   月の出納帳（日別の繰越・現金売上・入出金・帳簿残高・実査・過不足）
// GET  ?status=1                          ホーム用: 店舗ごとの未締め・締め後の変更・直近の実際の現金売上
// POST { action, shopId, ... }
//   entry.add    { date, type, cat, amount, memo, payee, receipt, opId }  入金・出金の記録
//   entry.edit   { date, entryId, patch: {...}, opId }                    記録の修正（締めた日は不可）
//   entry.void   { date, entryId, reason, opId }                          記録の取消（削除はせず取消線で残す）
//   count.save   { date, denoms: { "10000": 枚数, ... }, opId }          レジの実査（金種ごとの枚数）
//   day.close    { date, reason, opId }                                   日次締め（過不足がある場合は理由が必須）
//   day.reopen   { date, reason, opId }                                   締めの取消（理由が必須・履歴に残る）
//   settings.shop  { startDate, initialBalance, float }                   開始日・開始時の残高・釣り銭準備金
//   settings.brand { categories, cashMethodIds }                          科目・現金とみなす支払い方法（オーナー/マネージャー）
//
// 操作はすべて操作ログ（誰が・いつ・何を・変更前後・理由）に追記され、ログは修正・削除できない。
// 権限: オーナー/マネージャー = 全店舗 / 店長 = 自店舗 / スタッフ = 利用不可

'use strict';

const { getSession, readJsonBody } = require('./_lib/auth');
const { kvAvailable, kvGet, kvUpdate } = require('./_lib/kv');
const { fetchSalonOne } = require('./_lib/salonone');
const C = require('./_lib/cashbook');

const DOC_LIMIT = 900 * 1024;

function bad(res, status, error, extra) {
    res.statusCode = status;
    res.end(JSON.stringify({ error, ...extra }));
}

class Reject extends Error {
    constructor(status, code, detail) {
        super(code);
        this.status = status;
        this.code = code;
        this.detail = detail;
    }
}

async function shopStaffs(shopId) {
    const raw = await fetchSalonOne('staffs', { shop_id: shopId });
    const list = Array.isArray(raw) ? raw : (raw.data || []);
    return list.filter(s => !s.deleted_at && String(s.shop_id ?? shopId) === String(shopId));
}

// 担当者（実際に作業したスタッフ）の名前。店舗のスタッフ以外は無視する
async function operatorName(shopId, opId) {
    if (opId === undefined || opId === null || opId === '') return null;
    try {
        const st = (await shopStaffs(shopId)).find(s => String(s.id) === String(opId));
        return st ? String(st.name || '') || null : null;
    } catch (_) {
        return null;
    }
}

async function accessibleShops(session) {
    if (session.role === 'store') return [String(session.shopId)];
    const raw = await fetchSalonOne('shops', {});
    const list = Array.isArray(raw) ? raw : (raw.data || []);
    return list.filter(s => !s.deleted_at).map(s => String(s.id));
}

function logItem(session, op, fields) {
    return { id: C.newId('l_'), at: new Date().toISOString(), by: C.actorOf(session), role: session.role, op: op || null, ...fields };
}

function entrySummary(e) {
    return { type: e.type, cat: e.cat, amount: e.amount, memo: e.memo || '', payee: e.payee || '', receipt: !!e.receipt, date: e.date };
}

function validateEntryFields(src, cfg, { partial }) {
    const out = {};
    if (!partial || src.type !== undefined) {
        if (src.type !== 'in' && src.type !== 'out') throw new Reject(400, 'invalid_request', '入金/出金の区分が不正です');
        out.type = src.type;
    }
    if (!partial || src.cat !== undefined) {
        const cat = cfg.categories.find(c => c.id === src.cat);
        if (!cat || cat.disabled) throw new Reject(400, 'invalid_request', '科目が不正です');
        out.cat = cat.id;
    }
    if (!partial || src.amount !== undefined) {
        const n = C.intIn(src.amount, 1, C.MAX_AMOUNT);
        if (n === null) throw new Reject(400, 'invalid_request', '金額は1円以上の整数で入力してください');
        out.amount = n;
    }
    for (const [f, max] of [['memo', 100], ['payee', 40]]) {
        if (!partial || src[f] !== undefined) {
            const t = C.cleanText(src[f], max);
            if (t === null) throw new Reject(400, 'invalid_request', `${f === 'memo' ? '摘要' : '支払先'}は${max}文字以内で入力してください`);
            out[f] = t;
        }
    }
    if (!partial || src.receipt !== undefined) out.receipt = !!src.receipt;
    return out;
}

function checkDate(date, sc, today) {
    if (!C.DAY_RE.test(String(date || ''))) throw new Reject(400, 'invalid_request', '日付が不正です');
    if (date > today) throw new Reject(400, 'future_date', '未来の日付には記録できません');
    if (sc && date < sc.startDate) throw new Reject(400, 'before_start', '出納帳の開始日より前の日付です');
}

function assertOpen(doc, date) {
    if (doc.days[date]?.closed) throw new Reject(409, 'day_closed', 'この日は締め済みです。修正するには締めを取り消してください');
}

// ---- 書き込み（アクションごと）----
async function handlePost(session, body, res) {
    const action = String(body.action || '');
    const today = C.todayJst();
    const shopId = String(body.shopId ?? '');

    if (action === 'settings.brand') {
        if (!C.isAdminLike(session)) throw new Reject(403, 'forbidden', '科目・支払い方法の設定はオーナー・マネージャーのみ変更できます');
        await kvUpdate(C.CFG_KEY, current => {
            const cfg = C.sanitizeCfg(current);
            const before = { categories: cfg.categories, cashMethodIds: cfg.cashMethodIds };
            if (body.categories !== undefined) {
                if (!Array.isArray(body.categories)) throw new Reject(400, 'invalid_request', '科目が不正です');
                const list = body.categories.map(c => (c && c.id === 'new')
                    ? { ...c, id: C.newId('c_') }
                    : c);
                cfg.categories = C.sanitizeCategories(list);
            }
            if (body.cashMethodIds !== undefined) {
                if (!Array.isArray(body.cashMethodIds)) throw new Reject(400, 'invalid_request', '支払い方法が不正です');
                cfg.cashMethodIds = C.sanitizeCfg({ cashMethodIds: body.cashMethodIds }).cashMethodIds;
            }
            cfg.log.push(logItem(session, null, {
                action: 'settings.brand',
                before, after: { categories: cfg.categories, cashMethodIds: cfg.cashMethodIds },
            }));
            cfg.log = cfg.log.slice(-C.CFG_LOG_LIMIT);
            return cfg;
        });
        const cfg = await C.loadCfg();
        return res.end(JSON.stringify({ ok: true, settings: { categories: cfg.categories, cashMethodIds: cfg.cashMethodIds } }));
    }

    if (!/^\d+$/.test(shopId)) throw new Reject(400, 'invalid_request', '店舗が指定されていません');
    if (!C.canAccessShop(session, shopId)) throw new Reject(403, 'forbidden', '他店舗の出納帳は操作できません');

    if (action === 'settings.shop') {
        await kvUpdate(C.CFG_KEY, current => {
            const cfg = C.sanitizeCfg(current);
            const cur = cfg.shops[shopId] || null;
            const next = { ...(cur || { startDate: null, initialBalance: 0, float: 0 }) };
            const touchesStart = body.startDate !== undefined || body.initialBalance !== undefined;
            // 店長は初期設定（未設定のとき）と釣り銭準備金のみ。開始日・開始残高の変更はオーナー/マネージャー
            if (touchesStart && cur && !C.isAdminLike(session)) {
                throw new Reject(403, 'forbidden', '開始日・開始時の残高の変更はオーナー・マネージャーのみできます');
            }
            if (body.startDate !== undefined) {
                if (!C.DAY_RE.test(String(body.startDate)) || body.startDate > today || body.startDate < '2020-01-01') {
                    throw new Reject(400, 'invalid_request', '開始日が不正です（今日以前の日付）');
                }
                next.startDate = body.startDate;
            }
            if (body.initialBalance !== undefined) {
                const n = C.intIn(body.initialBalance, 0, C.MAX_AMOUNT);
                if (n === null) throw new Reject(400, 'invalid_request', '開始時の残高が不正です');
                next.initialBalance = n;
            }
            if (body.float !== undefined) {
                const n = C.intIn(body.float, 0, C.MAX_AMOUNT);
                if (n === null) throw new Reject(400, 'invalid_request', '釣り銭準備金が不正です');
                next.float = n;
            }
            if (!next.startDate) throw new Reject(400, 'invalid_request', '開始日を入力してください');
            next.updatedAt = new Date().toISOString();
            next.updatedBy = C.actorOf(session);
            cfg.shops[shopId] = next;
            cfg.log.push(logItem(session, null, {
                action: 'settings.shop', shopId: Number(shopId),
                before: cur ? { startDate: cur.startDate, initialBalance: cur.initialBalance, float: cur.float } : null,
                after: { startDate: next.startDate, initialBalance: next.initialBalance, float: next.float },
            }));
            cfg.log = cfg.log.slice(-C.CFG_LOG_LIMIT);
            return cfg;
        });
        const month = C.MONTH_RE.test(String(body.month || '')) ? body.month : C.monthOf(today);
        const cfg = await C.loadCfg();
        const view = await C.buildView(shopId, month, cfg, { today, docs: new Map() });
        return res.end(JSON.stringify({ ok: true, storage: 'kv', ...decorate(view, session) }));
    }

    const cfg = await C.loadCfg();
    const sc = cfg.shops[shopId];
    if (!sc) throw new Reject(409, 'setup_required', '先に出納帳の開始日と開始時の残高を設定してください');
    const date = String(body.date || '');
    checkDate(date, sc, today);
    const month = C.monthOf(date);
    const key = C.docKey(shopId, month);
    const op = await operatorName(shopId, body.opId);

    // 締めの計算に使う前月以前の帳簿・SalonOneのデータは先に温めておく（ロック中の処理を短くする）
    const ctx = { today, docs: new Map() };
    if (action === 'day.close') await C.buildView(shopId, month, cfg, ctx);

    await kvUpdate(key, async current => {
        const doc = C.normDoc(current);
        if (action === 'entry.add') {
            assertOpen(doc, date);
            const f = validateEntryFields(body, cfg, { partial: false });
            const cat = cfg.categories.find(c => c.id === f.cat);
            if (cat.type !== f.type) throw new Reject(400, 'invalid_request', '科目と入金/出金の区分が一致しません');
            if (doc.entries.length >= 3000) throw new Reject(413, 'too_large', '1ヶ月の記録件数の上限に達しました');
            const e = {
                id: C.newId('e_'), date, ...f, op,
                createdAt: new Date().toISOString(), createdBy: C.actorOf(session), rev: 1,
            };
            doc.entries.push(e);
            doc.log.push(logItem(session, op, { action: 'entry.add', date, entryId: e.id, after: entrySummary(e) }));
        } else if (action === 'entry.edit') {
            const e = doc.entries.find(x => x.id === body.entryId);
            if (!e) throw new Reject(404, 'not_found', '記録が見つかりません');
            if (e.voided) throw new Reject(409, 'voided', '取り消し済みの記録は修正できません');
            assertOpen(doc, e.date);
            const patch = body.patch && typeof body.patch === 'object' ? body.patch : {};
            const f = validateEntryFields(patch, cfg, { partial: true });
            if (patch.date !== undefined && patch.date !== e.date) {
                checkDate(patch.date, sc, today);
                if (C.monthOf(patch.date) !== month) throw new Reject(400, 'invalid_request', '別の月への移動はできません（取り消して記録し直してください）');
                assertOpen(doc, patch.date);
                f.date = patch.date;
            }
            const next = { ...e, ...f };
            const cat = cfg.categories.find(c => c.id === next.cat);
            if (!cat || cat.type !== next.type) throw new Reject(400, 'invalid_request', '科目と入金/出金の区分が一致しません');
            const before = {}, after = {};
            for (const k of ['date', 'type', 'cat', 'amount', 'memo', 'payee', 'receipt']) {
                if (JSON.stringify(e[k]) !== JSON.stringify(next[k])) { before[k] = e[k]; after[k] = next[k]; }
            }
            if (!Object.keys(after).length) return null; // 変更なし
            Object.assign(e, f, { updatedAt: new Date().toISOString(), updatedBy: C.actorOf(session), rev: (e.rev || 1) + 1 });
            if (op) e.op = op;
            doc.log.push(logItem(session, op, { action: 'entry.edit', date: e.date, entryId: e.id, before, after, amount: e.amount, cat: e.cat, type: e.type }));
        } else if (action === 'entry.void') {
            const e = doc.entries.find(x => x.id === body.entryId);
            if (!e) throw new Reject(404, 'not_found', '記録が見つかりません');
            if (e.voided) throw new Reject(409, 'voided', 'すでに取り消されています');
            assertOpen(doc, e.date);
            const reason = C.cleanText(body.reason, 200);
            if (!reason) throw new Reject(400, 'reason_required', '取り消しの理由を入力してください');
            e.voided = { at: new Date().toISOString(), by: C.actorOf(session), op, reason };
            doc.log.push(logItem(session, op, { action: 'entry.void', date: e.date, entryId: e.id, before: entrySummary(e), reason }));
        } else if (action === 'count.save') {
            assertOpen(doc, date);
            const denoms = {};
            let total = 0;
            for (const d of C.DENOMS) {
                const n = C.intIn(body.denoms?.[d] ?? body.denoms?.[String(d)] ?? 0, 0, 99999);
                if (n === null) throw new Reject(400, 'invalid_request', `${d}円の枚数が不正です`);
                if (n) denoms[d] = n;
                total += d * n;
            }
            if (total > C.MAX_AMOUNT) throw new Reject(400, 'invalid_request', '合計が大きすぎます');
            const prev = doc.days[date]?.count || null;
            doc.days[date] = { ...(doc.days[date] || {}), count: { denoms, total, at: new Date().toISOString(), by: C.actorOf(session), op } };
            doc.log.push(logItem(session, op, { action: 'count.save', date, before: prev ? { total: prev.total } : null, after: { total } }));
        } else if (action === 'day.close') {
            assertOpen(doc, date);
            ctx.docs.set(month, doc); // ロック中の最新の帳簿で計算し直す
            const view = await C.buildView(shopId, month, cfg, ctx);
            const row = view.rows.find(r => r.date === date);
            if (!row) throw new Reject(400, 'invalid_request', 'この日は締められません');
            if (!row.count) throw new Reject(409, 'count_required', '先にレジの現金を数えて（実査）保存してください');
            if (view.salonError || row.expected === null) throw new Reject(503, 'salon_unavailable', 'SalonOneの売上を取得できないため締められません。時間をおいて再度お試しください');
            const diff = row.count.total - row.expected;
            const reason = C.cleanText(body.reason, 200);
            if (reason === null) throw new Reject(400, 'invalid_request', '理由は200文字以内で入力してください');
            if (diff !== 0 && !reason) throw new Reject(400, 'reason_required', '過不足があります。理由を入力してください');
            const closed = {
                at: new Date().toISOString(), by: C.actorOf(session), op,
                opening: row.opening, salonCash: row.salonCash, cashMethods: row.cashMethods,
                manualIn: row.manualIn, manualOut: row.manualOut,
                expected: row.expected, counted: row.count.total, diff,
                actualCash: row.count.total - row.opening - row.manualIn + row.manualOut,
                denoms: row.count.denoms, reason: reason || '',
            };
            doc.days[date] = { ...(doc.days[date] || {}), closed };
            doc.log.push(logItem(session, op, {
                action: 'day.close', date,
                after: { opening: closed.opening, salonCash: closed.salonCash, manualIn: closed.manualIn, manualOut: closed.manualOut, expected: closed.expected, counted: closed.counted, diff },
                reason: reason || '',
            }));
        } else if (action === 'day.reopen') {
            const rec = doc.days[date];
            if (!rec?.closed) throw new Reject(409, 'not_closed', 'この日は締められていません');
            const reason = C.cleanText(body.reason, 200);
            if (!reason) throw new Reject(400, 'reason_required', '締めを取り消す理由を入力してください');
            const history = Array.isArray(rec.history) ? rec.history : [];
            history.push({ closed: rec.closed, reopenedAt: new Date().toISOString(), reopenedBy: C.actorOf(session), op, reason });
            doc.days[date] = { ...rec, closed: null, history: history.slice(-20) };
            delete doc.days[date].closed;
            doc.log.push(logItem(session, op, {
                action: 'day.reopen', date,
                before: { counted: rec.closed.counted, diff: rec.closed.diff, expected: rec.closed.expected },
                reason,
            }));
        } else {
            throw new Reject(400, 'invalid_request', `unknown action: ${action}`);
        }
        if (JSON.stringify(doc).length > DOC_LIMIT) throw new Reject(413, 'too_large', 'この月の出納帳のデータ量が上限に達しました');
        return doc;
    });

    const view = await C.buildView(shopId, month, cfg, { today, docs: new Map() });
    return res.end(JSON.stringify({ ok: true, storage: 'kv', ...decorate(view, session) }));
}

// 画面の出し分け用の権限フラグ
function decorate(view, session) {
    const setupDone = !!view.settings.shop;
    return {
        ...view,
        can: {
            manage: C.isAdminLike(session),                        // 科目・支払い方法・開始日の変更
            setup: C.isAdminLike(session) || !setupDone,            // 開始日・開始残高の設定
            write: true,
        },
    };
}

// ホーム用: 未締めの日・締め後の変更・締めた日の実際の現金売上（入金突合の自動入力用。今月と、7日前が前月なら前月も）
async function statusFor(session) {
    const today = C.todayJst();
    const cfg = await C.loadCfg();
    const shops = await accessibleShops(session);
    const months = [C.monthOf(today)];
    const weekAgo = C.addDays(today, -7);
    if (C.monthOf(weekAgo) !== months[0]) months.unshift(C.monthOf(weekAgo));
    const out = [];
    for (const shopId of shops) {
        if (!cfg.shops[shopId]) { out.push({ shopId: Number(shopId), configured: false }); continue; }
        const ctx = { today, docs: new Map() };
        const item = { shopId: Number(shopId), configured: true, pending: [], changed: [], days: {}, diffTotal: 0, diffDays: 0, closedDays: 0, salonError: false };
        for (const m of months) {
            const v = await C.buildView(shopId, m, cfg, ctx);
            if (!v.summary) continue;
            item.salonError = item.salonError || !!v.salonError;
            item.pending.push(...v.summary.pendingDays);
            item.changed.push(...v.summary.changedDays);
            if (m === C.monthOf(today)) {
                item.diffTotal = v.summary.diffTotal;
                item.diffDays = v.summary.diffDays;
                item.closedDays = v.summary.closedDays;
            }
            for (const r of v.rows) {
                if (!r.closed) continue;
                item.days[r.date] = { actualCash: r.closed.actualCash, cashMethodIds: (r.closed.cashMethods || []).map(x => x.id) };
            }
        }
        out.push(item);
    }
    return { today, shops: out };
}

module.exports = async (req, res) => {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Cache-Control', 'no-store');

    const session = getSession(req);
    if (!session) return bad(res, 401, 'auth_required');
    if (session.role === 'staff') return bad(res, 403, 'forbidden', { detail: '出納帳は店長以上のみ利用できます' });

    try {
        if (req.method === 'GET') {
            const url = new URL(req.url, 'http://localhost');
            if (!kvAvailable()) {
                return res.end(JSON.stringify({ storage: 'none' }));
            }
            if (url.searchParams.get('status') === '1') {
                return res.end(JSON.stringify({ storage: 'kv', ...(await statusFor(session)) }));
            }
            const shopId = url.searchParams.get('shop') || '';
            const month = url.searchParams.get('month') || '';
            if (!/^\d+$/.test(shopId) || !C.MONTH_RE.test(month)) return bad(res, 400, 'invalid_request', { fields: ['shop', 'month'] });
            if (!C.canAccessShop(session, shopId)) return bad(res, 403, 'forbidden');
            const cfg = await C.loadCfg();
            const ctx = { today: C.todayJst(), docs: new Map() };
            const view = await C.buildView(shopId, month, cfg, ctx);
            const body = { storage: 'kv', ...decorate(view, session) };
            if (url.searchParams.get('log') === '1') {
                const doc = ctx.docs.get(month) || C.normDoc(await kvGet(C.docKey(shopId, month)));
                // 設定の変更（この店舗分と科目の変更）は、操作した日本時間の月で絞り込んで合わせて返す
                const inMonth = l => C.monthOf(new Date(Date.parse(l.at) + 9 * 3600e3).toISOString()) === month;
                const cfgLog = cfg.log.filter(l => (l.action === 'settings.brand' || String(l.shopId) === shopId) && inMonth(l));
                body.log = [...doc.log, ...cfgLog].sort((a, b) => String(b.at).localeCompare(String(a.at)));
            }
            return res.end(JSON.stringify(body));
        }

        if (req.method !== 'POST') return bad(res, 405, 'method_not_allowed');
        if (!kvAvailable()) return bad(res, 501, 'storage_unconfigured', {
            detail: '出納帳にはサーバー保存（Supabase）が必要です（docs/SALONONE_INTEGRATION.md 参照）',
        });
        const body = await readJsonBody(req);
        return await handlePost(session, body, res);
    } catch (e) {
        if (e instanceof Reject) return bad(res, e.status, e.code, { detail: e.detail });
        console.error('cashbook api error', e);
        return bad(res, 500, 'internal_error');
    }
};
