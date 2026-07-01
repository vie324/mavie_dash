/* =====================================================================
 * vie Dashboard — 月次面談・振り返り機能 (Reviews)
 *
 *  - スタッフ: 月ごとに「自己評価(A〜E)」「定性面の振り返り」「次回アクション」を
 *    施術/物販/総売上の3区分で入力。目標・実績・達成率・乖離は自動算出。
 *  - 管理者: 各スタッフの振り返りを一覧し、目標×実績×自己評価を踏まえた
 *    AI評価（経営者視点）を生成・閲覧できる。
 *
 *  データは localStorage キャッシュ + バックエンド(apiFetch: GAS/Supabase)同期。
 *  レコード形状: reviews[ym][store][staff] = {
 *      meetingDate, interviewer, status,
 *      service:{rating,reflection,action}, retail:{...}, total:{...},
 *      metrics:{service,retail,total: {target,actual,rate,gap}},
 *      ai:{ text, ratedAt } , submittedAt, updatedAt
 *  }
 * ===================================================================== */
(function () {
    'use strict';

    const LS_KEY = 'mavie_reviews_v1';
    const RATINGS = ['S', 'A', 'B', 'C', 'D', 'E'];
    const STORE_NAMES = { chiba: '千葉店', honatsugi: '本厚木店', yamato: '大和店' };
    const SECTIONS = [
        { key: 'service', label: '施術売上', accent: '#5b8aa6', bg: 'rgba(91,138,166,0.10)' },
        { key: 'retail',  label: '物販売上', accent: '#5d9a72', bg: 'rgba(93,154,114,0.10)' },
        { key: 'total',   label: '総売上',   accent: '#6e819c', bg: 'rgba(110,129,156,0.10)' },
    ];

    const $ = id => document.getElementById(id);
    const fmt = n => Math.round(n || 0).toLocaleString();
    const yen = n => '¥' + fmt(n);
    const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    // 二重引用符HTML属性内の単一引用符JS文字列リテラル用エスケープ（onclick="fn('...')"）
    const jsq = s => String(s ?? '').replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '&quot;').replace(/</g, '\\x3c');
    const storeName = id => STORE_NAMES[id] || id || '';

    function lsGet() { try { return JSON.parse(localStorage.getItem(LS_KEY) || '{}'); } catch (e) { return {}; } }
    function lsSet(v) { try { localStorage.setItem(LS_KEY, JSON.stringify(v)); } catch (e) {} }

    let reviewsCache = lsGet();

    function currentStore() { return $('store-selector')?.value || 'all'; }
    function currentStaff() { return $('staff-selector')?.value || 'all'; }
    function isStaffScope() {
        return !!(lockedStaff) || (currentStaff() !== 'all' && !!currentStore() && currentStore() !== 'all');
    }
    function isAdminScope() { return !lockedStore && !lockedStaff; }

    // 面談タブで選択中の年月（YYYY/M）
    function selectedYM() {
        const sel = $('review-month-selector');
        if (sel && sel.value) {
            const p = sel.value.split('/');
            return p.length === 2 ? `${p[0]}/${parseInt(p[1])}` : sel.value;
        }
        const d = new Date();
        return `${d.getFullYear()}/${d.getMonth() + 1}`;
    }

    function getRecord(ym, store, staff) {
        return reviewsCache?.[ym]?.[store]?.[staff] || null;
    }
    function setRecord(ym, store, staff, rec) {
        reviewsCache[ym] = reviewsCache[ym] || {};
        reviewsCache[ym][store] = reviewsCache[ym][store] || {};
        reviewsCache[ym][store][staff] = rec;
        lsSet(reviewsCache);
    }

    /* ---------------- 目標×実績の自動算出 ---------------- */
    function computeMetrics(store, staff, ym) {
        const [y, m] = ym.split('/').map(Number);
        const start = new Date(y, m - 1, 1);
        const end = new Date(y, m, 0, 23, 59, 59);
        const data = (Array.isArray(rawData) ? rawData : []).filter(d => {
            if (!d || !d.date) return false;
            if (d.store !== store) return false;
            if (d.staff?.toLowerCase() !== staff.toLowerCase()) return false;
            const rd = parseDate(d.date);
            return rd >= start && rd <= end;
        });
        const met = calculateMetrics(data);
        const inc = calculateIncentive(data, store, staff);
        const goal = getStaffGoal(store, staff, ym);

        const serviceTarget = (goal.weekdays || 0) * (goal.weekdayTarget || 0) + (goal.weekends || 0) * (goal.weekendTarget || 0);
        const serviceActual = met.salesTotal;
        const retailTarget = goal.retail || 0;
        const retailActual = inc.retailSalesRaw || 0;
        const totalTarget = serviceTarget + retailTarget;
        const totalActual = serviceActual + retailActual;

        const mk = (target, actual) => ({
            target, actual,
            rate: target > 0 ? Math.round(actual / target * 100) : 0,
            gap: actual - target,
        });
        return {
            service: mk(serviceTarget, serviceActual),
            retail: mk(retailTarget, retailActual),
            total: mk(totalTarget, totalActual),
            customersTotal: met.customersTotal,
            customersNew: met.customersNew,
        };
    }

    /* ---------------- 月セレクタ初期化 ---------------- */
    function initMonthSelector() {
        const sel = $('review-month-selector');
        if (!sel || sel.dataset.ready) return;
        const now = new Date();
        const opts = [];
        for (let i = 0; i < 18; i++) {
            const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
            const val = `${d.getFullYear()}/${d.getMonth() + 1}`;
            opts.push(`<option value="${val}">${d.getFullYear()}年${d.getMonth() + 1}月</option>`);
        }
        sel.innerHTML = opts.join('');
        sel.dataset.ready = '1';
    }

    /* ---------------- メインレンダー ---------------- */
    function render() {
        initMonthSelector();
        const ym = selectedYM();
        const badge = $('review-mode-badge');
        if (isStaffScope()) {
            const store = lockedStore || currentStore();
            const staff = lockedStaff || currentStaff();
            if (badge) { badge.textContent = `${esc(staff)} さんの振り返り`; badge.className = 'review-badge staff'; }
            renderStaffForm(store, staff, ym);
        } else {
            if (badge) { badge.textContent = '管理者ビュー（全スタッフ）'; badge.className = 'review-badge admin'; }
            renderAdminList(ym);
        }
        if (window.lucide) lucide.createIcons();
    }

    /* ---------------- スタッフ入力フォーム ---------------- */
    function metricsPanel(metrics, sectionKey) {
        const mm = metrics[sectionKey];
        const rateCls = mm.rate >= 100 ? 'good' : mm.rate >= 80 ? 'mid' : 'low';
        return `<div class="rv-metrics">
            <div><span class="rv-mlabel">目標</span><span class="rv-mval">${yen(mm.target)}</span></div>
            <div><span class="rv-mlabel">実績</span><span class="rv-mval">${yen(mm.actual)}</span></div>
            <div><span class="rv-mlabel">達成率</span><span class="rv-mrate ${rateCls}">${mm.rate}%</span></div>
            <div><span class="rv-mlabel">目標乖離</span><span class="rv-mval ${mm.gap >= 0 ? 'pos' : 'neg'}">${mm.gap >= 0 ? '+' : ''}${yen(mm.gap)}</span></div>
        </div>`;
    }

    function ratingSelect(id, value) {
        return `<select id="${id}" class="rv-rating-select">
            <option value="">—</option>
            ${RATINGS.map(r => `<option value="${r}" ${value === r ? 'selected' : ''}>${r}</option>`).join('')}
        </select>`;
    }

    function renderStaffForm(store, staff, ym) {
        const host = $('review-body');
        if (!host) return;
        const metrics = computeMetrics(store, staff, ym);
        const rec = getRecord(ym, store, staff) || {};
        const readonly = false; // 自分の振り返りは編集可

        const sectionsHtml = SECTIONS.map(s => {
            const sr = rec[s.key] || {};
            return `
            <div class="rv-section" style="--rv-accent:${s.accent};--rv-bg:${s.bg}">
                <div class="rv-section-head"><span class="rv-dot"></span>${s.label}</div>
                ${metricsPanel(metrics, s.key)}
                <div class="rv-inputs">
                    <label class="rv-field rv-field-rating">
                        <span>自己評価</span>
                        ${ratingSelect(`rv-${s.key}-rating`, sr.rating || '')}
                    </label>
                    <label class="rv-field">
                        <span>定性面の振り返り</span>
                        <textarea id="rv-${s.key}-reflection" rows="3" class="rv-textarea" placeholder="良かった点・課題・気づきなど">${esc(sr.reflection || '')}</textarea>
                    </label>
                    <label class="rv-field">
                        <span>${s.key === 'total' ? '次回までのアクションまとめ' : '次回までの具体的なアクション'}</span>
                        <textarea id="rv-${s.key}-action" rows="2" class="rv-textarea" placeholder="次の1ヶ月で取り組むこと">${esc(sr.action || '')}</textarea>
                    </label>
                </div>
            </div>`;
        }).join('');

        const aiHtml = rec.ai && rec.ai.text
            ? `<div class="rv-ai-result">
                   <div class="rv-ai-head"><span>🤖</span> 店長・経営者からのAI評価</div>
                   <div class="rv-ai-body">${(window.Enhance?.mdLite ? Enhance.mdLite(rec.ai.text) : esc(rec.ai.text).replace(/\n/g, '<br>'))}</div>
                   <p class="rv-ai-time">${rec.ai.ratedAt ? new Date(rec.ai.ratedAt).toLocaleString('ja-JP') : ''}</p>
               </div>`
            : `<div class="rv-ai-pending"><i data-lucide="sparkles" class="w-4 h-4"></i> 振り返りを提出すると、管理者による確認とAI評価が受けられます。</div>`;

        host.innerHTML = `
            <div class="rv-form-head">
                <div class="rv-grid-2">
                    <label class="rv-field"><span>面談日時</span>
                        <input type="date" id="rv-meeting-date" class="rv-input" value="${rec.meetingDate || ''}"></label>
                    <label class="rv-field"><span>面談担当者</span>
                        <input type="text" id="rv-interviewer" class="rv-input" placeholder="担当者名" value="${esc(rec.interviewer || '')}"></label>
                </div>
            </div>
            <div class="rv-sections">${sectionsHtml}</div>
            <div class="rv-actions">
                <button onclick="Reviews.save('${jsq(store)}','${jsq(staff)}')" class="btn-primary py-3 px-6 rounded-xl font-bold flex items-center gap-2">
                    <i data-lucide="send" class="w-4 h-4"></i>振り返りを保存・提出
                </button>
                ${rec.submittedAt ? `<span class="rv-saved">最終保存: ${new Date(rec.updatedAt || rec.submittedAt).toLocaleString('ja-JP')}</span>` : ''}
            </div>
            ${aiHtml}`;
    }

    /* ---------------- 管理者：一覧 + 詳細 ---------------- */
    function rosterForScope() {
        const store = currentStore();
        const list = [];
        if (store === 'all') {
            Object.entries(STAFF_ROSTER || {}).forEach(([s, names]) => names.forEach(n => list.push({ store: s, name: n })));
        } else {
            (STAFF_ROSTER?.[store] || []).forEach(n => list.push({ store, name: n }));
        }
        return list;
    }

    function renderAdminList(ym) {
        const host = $('review-body');
        if (!host) return;
        const roster = rosterForScope();
        const rows = roster.map(({ store, name }) => {
            const rec = getRecord(ym, store, name);
            const metrics = computeMetrics(store, name, ym);
            const submitted = !!(rec && rec.submittedAt);
            const aiDone = !!(rec && rec.ai && rec.ai.text);
            const selfTotal = rec?.total?.rating || '—';
            return { store, name, rec, metrics, submitted, aiDone, selfTotal };
        });

        const listHtml = rows.map(r => `
            <tr class="rv-row" onclick="Reviews.openDetail('${jsq(r.store)}','${jsq(r.name)}')">
                <td class="rv-td-name">${esc(r.name)}<span class="rv-td-store">${storeName(r.store)}</span></td>
                <td class="text-right">${yen(r.metrics.total.actual)}</td>
                <td class="text-right"><span class="rv-mrate ${r.metrics.total.rate >= 100 ? 'good' : r.metrics.total.rate >= 80 ? 'mid' : 'low'}">${r.metrics.total.rate}%</span></td>
                <td class="text-center"><span class="rv-self-badge">${r.selfTotal}</span></td>
                <td class="text-center">${r.submitted ? '<span class="rv-status done">提出済</span>' : '<span class="rv-status pending">未提出</span>'}</td>
                <td class="text-center">${r.aiDone ? '<span class="rv-status ai">AI評価済</span>' : (r.submitted ? '<span class="rv-status wait">未評価</span>' : '<span class="rv-status">—</span>')}</td>
            </tr>`).join('');

        const submittedCount = rows.filter(r => r.submitted).length;
        host.innerHTML = `
            <div class="rv-admin-summary">
                <span><b>${ym}</b> の提出状況: ${submittedCount} / ${rows.length} 名</span>
            </div>
            <div class="overflow-x-auto">
                <table class="rv-table">
                    <thead><tr>
                        <th class="text-left">スタッフ</th>
                        <th class="text-right">総売上実績</th>
                        <th class="text-right">達成率</th>
                        <th class="text-center">自己評価</th>
                        <th class="text-center">提出</th>
                        <th class="text-center">AI評価</th>
                    </tr></thead>
                    <tbody>${listHtml || '<tr><td colspan="6" class="text-center py-6 text-surface-400">スタッフが登録されていません</td></tr>'}</tbody>
                </table>
            </div>
            <div id="review-detail" class="rv-detail hidden"></div>`;
    }

    function openDetail(store, staff) {
        const ym = selectedYM();
        const host = $('review-detail');
        if (!host) return;
        const metrics = computeMetrics(store, staff, ym);
        const rec = getRecord(ym, store, staff) || {};

        const sectionsHtml = SECTIONS.map(s => {
            const sr = rec[s.key] || {};
            return `
            <div class="rv-section" style="--rv-accent:${s.accent};--rv-bg:${s.bg}">
                <div class="rv-section-head"><span class="rv-dot"></span>${s.label}
                    <span class="rv-self-inline">自己評価 <b>${sr.rating || '—'}</b></span>
                </div>
                ${metricsPanel(metrics, s.key)}
                <div class="rv-readonly">
                    <p class="rv-ro-label">振り返り</p>
                    <p class="rv-ro-text">${sr.reflection ? esc(sr.reflection) : '<span class="rv-empty">未記入</span>'}</p>
                    <p class="rv-ro-label">次アクション</p>
                    <p class="rv-ro-text">${sr.action ? esc(sr.action) : '<span class="rv-empty">未記入</span>'}</p>
                </div>
            </div>`;
        }).join('');

        const aiHtml = rec.ai && rec.ai.text
            ? `<div class="rv-ai-result">
                   <div class="rv-ai-head"><span>🤖</span> AI評価
                       <button onclick="Reviews.runAi('${jsq(store)}','${jsq(staff)}')" class="rv-ai-regen">再生成</button>
                   </div>
                   <div class="rv-ai-body">${(window.Enhance?.mdLite ? Enhance.mdLite(rec.ai.text) : esc(rec.ai.text).replace(/\n/g, '<br>'))}</div>
                   <p class="rv-ai-time">${rec.ai.ratedAt ? new Date(rec.ai.ratedAt).toLocaleString('ja-JP') : ''}</p>
               </div>`
            : `<div class="rv-ai-cta"><span id="review-ai-status" class="text-sm text-surface-500">${rec.submittedAt ? '目標・実績・振り返りを踏まえてAIが評価します。' : 'このスタッフはまだ振り返りを提出していません（実績のみで評価も可能）。'}</span>
                   <button onclick="Reviews.runAi('${jsq(store)}','${jsq(staff)}')" id="review-ai-btn" class="btn-gold py-2 px-5 text-sm font-bold rounded-lg">AIで評価する</button></div>`;

        host.classList.remove('hidden');
        host.innerHTML = `
            <div class="rv-detail-head">
                <h4><i data-lucide="user" class="w-4 h-4"></i> ${esc(staff)} <span class="rv-td-store">${storeName(store)} / ${ym}</span></h4>
                <div class="rv-detail-meta">
                    ${rec.meetingDate ? `面談日: ${rec.meetingDate}` : ''} ${rec.interviewer ? `・担当: ${esc(rec.interviewer)}` : ''}
                </div>
            </div>
            <div class="rv-sections">${sectionsHtml}</div>
            ${aiHtml}`;
        host.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        if (window.lucide) lucide.createIcons();
    }

    /* ---------------- 保存（スタッフ入力） ---------------- */
    async function save(store, staff) {
        const ym = selectedYM();
        const readVal = id => ($(id)?.value || '').trim();
        const metrics = computeMetrics(store, staff, ym);
        const existing = getRecord(ym, store, staff) || {};
        const rec = {
            ...existing,
            meetingDate: readVal('rv-meeting-date'),
            interviewer: readVal('rv-interviewer'),
            service: { rating: readVal('rv-service-rating'), reflection: readVal('rv-service-reflection'), action: readVal('rv-service-action') },
            retail:  { rating: readVal('rv-retail-rating'),  reflection: readVal('rv-retail-reflection'),  action: readVal('rv-retail-action') },
            total:   { rating: readVal('rv-total-rating'),   reflection: readVal('rv-total-reflection'),   action: readVal('rv-total-action') },
            metrics,
            status: 'submitted',
            submittedAt: existing.submittedAt || Date.now(),
            updatedAt: Date.now(),
        };
        setRecord(ym, store, staff, rec);

        // バックエンド同期
        try {
            await syncSave({ yearMonth: ym, store, staff, ...rec });
            if (typeof showSettingsToast === 'function') showSettingsToast('振り返りを保存しました。おつかれさまでした！');
            if (window.Enhance?.fireConfetti) Enhance.fireConfetti({ count: 50, origin: 0.3, spread: 0.6 });
        } catch (e) {
            if (typeof showSettingsToast === 'function') showSettingsToast('ローカルに保存しました（サーバー同期は後で再試行されます）', 'warning');
        }
        render();
    }

    /* ---------------- AI評価（管理者） ---------------- */
    function buildPrompt(store, staff, ym, rec, metrics) {
        const line = (label, s, mm) => `【${label}】目標 ${yen(mm.target)} / 実績 ${yen(mm.actual)}（達成率 ${mm.rate}%・乖離 ${mm.gap >= 0 ? '+' : ''}${yen(mm.gap)}）
  自己評価: ${s?.rating || '未評価'} / 振り返り: 「${(s?.reflection || '記入なし').replace(/\n/g, ' ')}」 / 次アクション: 「${(s?.action || '記入なし').replace(/\n/g, ' ')}」`;
        return `あなたはアイラッシュサロン「vie」の経営者兼店長です。スタッフ「${staff}」さん（${storeName(store)}）の ${ym} 月次面談における自己評価と実績を確認し、成長を後押しする建設的な評価を日本語で行ってください。

${line('施術売上', rec.service, metrics.service)}
${line('物販売上', rec.retail, metrics.retail)}
${line('総売上', rec.total, metrics.total)}

参考: 当月来店 ${metrics.customersTotal}名（うち新規 ${metrics.customersNew}名）

以下の見出し・順序でMarkdownで出力してください（各項目は簡潔に）:
**■ 総合評価** … S/A/B/C/D の5段階で判定し、一言添える
**■ 施術・物販の講評** … 実績と本人の自己評価のギャップに触れ、事実ベースで評価
**■ 良かった点** … 具体的に2つ、箇条書き
**■ 次月の重点アクション** … 優先度順に3つ、すぐ実行できる具体策を箇条書き
**■ ひとことメッセージ** … 前向きな応援を1〜2文

辛口すぎず甘すぎず、経営者として公平に。本人の振り返りが実績と乖離している場合はやさしく指摘してください。`;
    }

    async function runAi(store, staff) {
        const ym = selectedYM();
        const rec = getRecord(ym, store, staff) || {};
        const metrics = computeMetrics(store, staff, ym);
        const btn = $('review-ai-btn');
        const statusEl = $('review-ai-status');
        if (btn) { btn.disabled = true; btn.textContent = 'AIが評価中…'; }
        if (statusEl) statusEl.innerHTML = '<span class="ai-thinking"><span></span><span></span><span></span> 目標と実績を分析しています…</span>';
        try {
            if (!window.Enhance || !Enhance.callGemini) throw new Error('AI機能が初期化されていません');
            const text = await Enhance.callGemini(buildPrompt(store, staff, ym, rec, metrics));
            const updated = { ...(getRecord(ym, store, staff) || {}), ai: { text, ratedAt: Date.now() } };
            // 実績スナップショットも保存
            updated.metrics = metrics;
            setRecord(ym, store, staff, updated);
            try { await syncAi({ yearMonth: ym, store, staff, ai: updated.ai, metrics }); } catch (e) { /* ローカルは保持 */ }
            openDetail(store, staff);
            if (typeof showSettingsToast === 'function') showSettingsToast('AI評価を生成しました');
        } catch (e) {
            if (statusEl) statusEl.innerHTML = `<span class="text-rose-500 text-sm">${esc(e.message)}</span>`;
            if (btn) { btn.disabled = false; btn.textContent = 'AIで評価する'; }
        }
    }

    /* ---------------- バックエンド同期 ---------------- */
    const send = (url, opts) => (window.apiFetch || fetch)(url, opts);
    function apiBase() {
        return localStorage.getItem('mavie_spreadsheet_api_url')
            || (typeof API_URL !== 'undefined' && API_URL)
            || (typeof DEFAULT_API_URL !== 'undefined' && DEFAULT_API_URL)
            || (window.Backend && Backend.mode() === 'supabase' ? 'supabase://rpc' : '');
    }

    async function syncSave(review) {
        const url = apiBase();
        if (!url) return;
        const res = await send(url, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ action: 'save_review', review }) });
        const r = await res.json();
        if (r.status !== 'success') throw new Error(r.message || '保存に失敗しました');
    }
    async function syncAi(payload) {
        const url = apiBase();
        if (!url) return;
        const res = await send(url, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ action: 'save_review_ai', ...payload }) });
        const r = await res.json();
        if (r.status !== 'success') throw new Error(r.message || '保存に失敗しました');
    }

    async function loadReviews() {
        const url = apiBase();
        if (!url) return;
        try {
            const res = await send(`${url}${url.includes('?') ? '&' : '?'}action=get_reviews`);
            const r = await res.json();
            if (r.status === 'success' && r.reviews) {
                reviewsCache = mergeDeep(reviewsCache, r.reviews);
                lsSet(reviewsCache);
            }
        } catch (e) { console.warn('レビュー読み込みスキップ:', e.message); }
    }

    // サーバー優先で浅くマージ（同一レコードはサーバー値を採用）
    function mergeDeep(local, remote) {
        const out = { ...local };
        Object.keys(remote || {}).forEach(ym => {
            out[ym] = out[ym] || {};
            Object.keys(remote[ym]).forEach(store => {
                out[ym][store] = { ...(out[ym][store] || {}), ...remote[ym][store] };
            });
        });
        return out;
    }

    /* ---------------- 初期化・フック ---------------- */
    document.addEventListener('DOMContentLoaded', () => {
        // データ読み込み後にサーバーからレビューを取得
        setTimeout(() => { loadReviews().then(() => { if (isReviewTabActive()) render(); }); }, 1500);
    });

    function isReviewTabActive() {
        const el = $('content-review');
        return el && !el.classList.contains('hidden');
    }

    window.Reviews = {
        render,
        save,
        runAi,
        openDetail,
        onMonthChange: render,
        loadReviews,
        isReviewTabActive,
    };
})();
