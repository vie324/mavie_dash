/* =====================================================================
 * vie Dashboard Enhancements (Phase 1-3)
 *  - 目標リング / あと◯◯で達成 / パーソナル挨拶 / 達成セレブレーション
 *  - LIVE自動更新 / 今週のハイライト(MVP・自己ベスト・異常検知・ストリーク)
 *  - 店舗対抗レース&ベンチマーク / 曜日別分析 / 前年比 / 着地予測v2
 *  - スタッフスキルレーダー / AIコーチ / フォロー候補 / 媒体ROI
 *  - アプリ内日報入力 / インセンティブ月締め確定 / PWA登録
 * dashboard.js のグローバル(rawData, calculateMetrics 等)に依存し、
 * updateDashboard() 末尾の Enhance.onDashboardUpdated() から駆動される。
 * ===================================================================== */
(function () {
    'use strict';

    window.GEMINI_MODEL = 'gemini-2.0-flash';

    const LIVE_INTERVAL_MS = 5 * 60 * 1000;
    const STORE_NAMES = { chiba: '千葉店', honatsugi: '本厚木店', yamato: '大和店' };
    const LS = {
        celebrated: 'mavie_celebrated_v1',
        rankHistory: 'mavie_rank_history_v1',
        bestCelebrated: 'mavie_best_celebrated_v1',
        adCosts: 'mavie_ad_costs',
        monthlyClose: 'mavie_monthly_close',
        aiCache: 'mavie_ai_advice_cache_v1',
    };

    const $ = id => document.getElementById(id);
    const fmt = n => Math.round(n).toLocaleString();
    const yen = n => '¥' + fmt(n);
    const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    const storeName = id => STORE_NAMES[id] || id || '';
    const todayKey = () => { const d = new Date(); return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`; };
    const monthKey = (d = new Date()) => `${d.getFullYear()}/${d.getMonth() + 1}`;
    const reducedMotion = () => window.matchMedia('(prefers-reduced-motion: reduce)').matches;

    function lsGet(key, fallback) {
        try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch (e) { return fallback; }
    }
    function lsSet(key, value) {
        try { localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* quota */ }
    }

    function currentStore() { return $('store-selector')?.value || 'all'; }
    function currentStaff() { return $('staff-selector')?.value || 'all'; }
    function scopeKey() { return `${currentStore()}|${currentStaff()}`; }
    function selectedYM() { return $('date-selector')?.value || monthKey(); }
    function isStaffPage() { return !!(lockedStaff || lockedStore && lockedStaff); }
    function isAdminView() { return !lockedStore && !lockedStaff; }

    // 現在の店舗/スタッフフィルタで期間を指定してレコード抽出
    function filterScope(from, to, { store, staff } = {}) {
        const st = store !== undefined ? store : currentStore();
        const sf = staff !== undefined ? staff : currentStaff();
        return (Array.isArray(rawData) ? rawData : []).filter(d => {
            if (!d || !d.date) return false;
            if (st !== 'all' && d.store !== st) return false;
            if (sf !== 'all' && d.staff?.toLowerCase() !== sf.toLowerCase()) return false;
            const rd = parseDate(d.date);
            return rd >= from && rd <= to;
        });
    }

    // 今月1日〜今日のメトリクス（現在スコープ）
    function monthToDateMetrics() {
        const now = new Date();
        const start = new Date(now.getFullYear(), now.getMonth(), 1);
        const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
        return calculateMetrics(filterScope(start, end));
    }

    // 現在スコープの目標データ（リング・残額・達成判定は常に「実際の今月」基準）
    function scopedGoalData() {
        const ym = monthKey();
        const ctx = getCurrentGoalContext();
        if (ctx.type === 'staff') return getStaffGoal(ctx.store, ctx.staff, ym);
        if (ctx.type === 'store') return getStoreAggregateGoal(ctx.store, ym);
        return getAllStoresAggregateGoal(ym);
    }
    const goalMonthlyTotal = g => (g.weekdays || 0) * (g.weekdayTarget || 0) + (g.weekends || 0) * (g.weekendTarget || 0);

    /* ================= Confetti (自前・軽量) ================= */
    let confettiCanvas = null;
    function fireConfetti({ count = 90, origin = 0.4, spread = 1 } = {}) {
        if (reducedMotion()) return;
        if (!confettiCanvas) {
            confettiCanvas = document.createElement('canvas');
            confettiCanvas.id = 'confetti-canvas';
            confettiCanvas.style.cssText = 'position:fixed;inset:0;pointer-events:none;z-index:9998;';
            document.body.appendChild(confettiCanvas);
        }
        const cv = confettiCanvas, ctx = cv.getContext('2d');
        cv.width = window.innerWidth; cv.height = window.innerHeight;
        const colors = ['#b8956a', '#c9a96e', '#739977', '#b08f8a', '#566882', '#e8d5b5'];
        const parts = Array.from({ length: count }, () => ({
            x: cv.width * (0.5 + (Math.random() - 0.5) * spread * 0.8),
            y: cv.height * origin,
            vx: (Math.random() - 0.5) * 11,
            vy: -(Math.random() * 11 + 5),
            w: 6 + Math.random() * 5,
            h: 8 + Math.random() * 7,
            rot: Math.random() * Math.PI,
            vr: (Math.random() - 0.5) * 0.3,
            color: colors[(Math.random() * colors.length) | 0],
            life: 1,
        }));
        const t0 = performance.now();
        (function frame(now) {
            const dt = Math.min((now - t0) / 1600, 1);
            ctx.clearRect(0, 0, cv.width, cv.height);
            parts.forEach(p => {
                p.vy += 0.28; p.x += p.vx; p.y += p.vy; p.rot += p.vr; p.life = 1 - dt;
                ctx.save();
                ctx.globalAlpha = Math.max(p.life, 0);
                ctx.translate(p.x, p.y); ctx.rotate(p.rot);
                ctx.fillStyle = p.color;
                ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
                ctx.restore();
            });
            if (dt < 1) requestAnimationFrame(frame);
            else ctx.clearRect(0, 0, cv.width, cv.height);
        })(t0);
    }

    /* ================= 挨拶 + あと◯◯で達成 ================= */
    function renderGreeting(goal) {
        const el = $('snapshot-greeting');
        if (!el) return;
        const h = new Date().getHours();
        const hello = h < 4 ? 'おつかれさまです' : h < 11 ? 'おはようございます' : h < 18 ? 'こんにちは' : 'おつかれさまです';
        let name = '';
        const staff = lockedStaff || (currentStaff() !== 'all' ? currentStaff() : '');
        if (staff) name = `、${esc(staff)} さん`;
        let cheer = '';
        if (goal > 0) {
            const m = monthToDateMetrics();
            const now = new Date();
            const expected = now.getDate() / new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate() * 100;
            const actual = m.salesTotal / goal * 100;
            if (actual >= 100) cheer = '月間目標を達成しています。素晴らしい！🎉';
            else if (actual >= expected + 3) cheer = 'いいペースです。この調子で！✨';
            else if (actual >= expected - 5) cheer = '予定どおりのペースです。';
            else cheer = 'ここから巻き返しのチャンスです💪';
        }
        el.innerHTML = `${hello}${name}。<span class="text-primary-600 dark:text-primary-400">${cheer}</span>`;
    }

    function renderRemaining(goal) {
        const el = $('pace-remaining');
        if (!el) return;
        if (!goal || goal <= 0) { el.innerHTML = ''; return; }
        const m = monthToDateMetrics();
        const now = new Date();
        const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
        const remainDays = Math.max(daysInMonth - now.getDate() + 1, 1);
        const remain = goal - m.salesTotal;
        if (remain <= 0) {
            el.innerHTML = `<span class="remaining-chip done"><i data-lucide="party-popper" class="w-3.5 h-3.5"></i>月間目標達成！ +${yen(-remain)} 超過</span>`;
        } else {
            const perDay = Math.ceil(remain / remainDays);
            el.innerHTML = `<span class="remaining-chip"><i data-lucide="flag" class="w-3.5 h-3.5"></i>達成まで あと <b>${yen(remain)}</b></span>
                <span class="remaining-sub">残り${remainDays}日 → 1日 <b>${yen(perDay)}</b> ペース</span>`;
        }
    }

    /* ================= 着地予測 v2（曜日係数モデル + レンジ） ================= */
    function renderForecastV2(goal) {
        const el = $('pace-forecast');
        const rangeEl = $('pace-forecast-range');
        if (!el) return;
        const now = new Date();
        const start = new Date(now.getFullYear(), now.getMonth(), 1);
        const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
        const m = calculateMetrics(filterScope(start, end));
        const dayKeys = Object.keys(m.daily || {});
        const mtd = m.salesTotal;
        const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();

        if (dayKeys.length < 5) { // データ僅少時は単純日割り（既存表示のまま）
            if (rangeEl) rangeEl.textContent = '';
            return;
        }

        // 曜日別平均（営業実績のある日のみ）
        const byDow = Array.from({ length: 7 }, () => []);
        const dailyVals = [];
        dayKeys.forEach(k => {
            const v = m.daily[k].sales || 0;
            const dow = parseDate(k).getDay();
            byDow[dow].push(v);
            dailyVals.push(v);
        });
        const overallAvg = dailyVals.reduce((a, b) => a + b, 0) / dailyVals.length;
        const dowAvg = byDow.map(arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : overallAvg);

        // 残り日数を曜日係数で積み上げ
        let projected = mtd;
        for (let day = now.getDate() + 1; day <= daysInMonth; day++) {
            projected += dowAvg[new Date(now.getFullYear(), now.getMonth(), day).getDay()];
        }
        // 予測レンジ: 日次標準偏差 × √残日数
        const variance = dailyVals.reduce((a, v) => a + Math.pow(v - overallAvg, 2), 0) / dailyVals.length;
        const sd = Math.sqrt(variance);
        const remainDays = daysInMonth - now.getDate();
        const band = sd * Math.sqrt(Math.max(remainDays, 0));
        const low = Math.max(mtd, projected - band);
        const high = projected + band;

        el.textContent = yen(projected);
        el.title = '曜日別の平均売上で残り日数を積み上げた予測値';
        if (rangeEl) {
            rangeEl.textContent = remainDays > 0 ? `${yen(low)} 〜 ${yen(high)}` : '今月の実績確定';
            if (goal > 0 && remainDays > 0) {
                rangeEl.textContent += projected >= goal ? ' ・目標到達見込み◎' : '';
            }
        }
    }

    /* ================= 目標リング ================= */
    function ringSvg(pct, color, over) {
        const r = 30, c = 2 * Math.PI * r;
        const clamped = Math.max(0, Math.min(pct, 100));
        const off = c * (1 - clamped / 100);
        return `<svg viewBox="0 0 72 72" class="goal-ring-svg${over ? ' over' : ''}">
            <circle cx="36" cy="36" r="${r}" class="goal-ring-track"/>
            <circle cx="36" cy="36" r="${r}" class="goal-ring-fill" stroke="${color}"
                stroke-dasharray="${c.toFixed(1)}" stroke-dashoffset="${off.toFixed(1)}" pathLength="${c.toFixed(1)}"/>
        </svg>`;
    }

    function renderGoalRings(metrics, goal) {
        const wrap = $('goal-rings');
        if (!wrap) return;
        const g = scopedGoalData();
        const m = monthToDateMetrics();

        const salesPct = goal > 0 ? m.salesTotal / goal * 100 : 0;
        const newPct = g.newCustomers > 0 ? m.customersNew / g.newCustomers * 100 : 0;
        const resRate = m.customersTotal > 0 ? m.nextRes.total / m.customersTotal * 100 : 0;
        const resPct = g.reservationRate > 0 ? resRate / g.reservationRate * 100 : 0;

        const rings = [
            { label: '売上', pct: salesPct, color: '#b8956a', value: yen(m.salesTotal), sub: goal > 0 ? `目標 ${yen(goal)}` : '目標未設定' },
            { label: '新規獲得', pct: newPct, color: '#739977', value: `${m.customersNew}名`, sub: g.newCustomers > 0 ? `目標 ${g.newCustomers}名` : '目標未設定' },
            { label: '次回予約率', pct: resPct, color: '#c9a96e', value: `${resRate.toFixed(0)}%`, sub: g.reservationRate > 0 ? `目標 ${g.reservationRate}%` : '目標未設定' },
        ];
        wrap.innerHTML = rings.map(r => `
            <div class="goal-ring">
                <div class="goal-ring-chart">
                    ${ringSvg(r.pct, r.color, r.pct >= 100)}
                    <div class="goal-ring-center">
                        <span class="goal-ring-pct" style="color:${r.color}">${Math.round(r.pct)}<small>%</small></span>
                    </div>
                    ${r.pct >= 100 ? '<span class="goal-ring-star">✨</span>' : ''}
                </div>
                <p class="goal-ring-label">${r.label}</p>
                <p class="goal-ring-value">${r.value}</p>
                <p class="goal-ring-sub">${r.sub}</p>
            </div>`).join('');
    }

    /* ================= セレブレーション ================= */
    function maybeCelebrate(goal) {
        try {
            const flags = lsGet(LS.celebrated, {});
            const scope = scopeKey();
            // 月間目標達成（スコープごとに月1回）
            if (goal > 0) {
                const key = `m:${monthKey()}:${scope}`;
                const m = monthToDateMetrics();
                if (m.salesTotal >= goal && !flags[key]) {
                    flags[key] = true; lsSet(LS.celebrated, flags);
                    fireConfetti({ count: 160, origin: 0.35, spread: 1.4 });
                    showCelebrationBanner('🎉 月間目標を達成しました！', `${yen(m.salesTotal)} / ${yen(goal)}`);
                }
            }
            // 日次目標達成（1日1回・小さめ）
            const g = scopedGoalData();
            const isWeekend = [0, 6].includes(new Date().getDay());
            const dailyTarget = isWeekend ? (g.weekendTarget || 0) : (g.weekdayTarget || 0);
            if (dailyTarget > 0) {
                const dKey = `d:${todayKey()}:${scope}`;
                const today = new Date();
                const t0 = new Date(today.getFullYear(), today.getMonth(), today.getDate());
                const t1 = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
                const tm = calculateMetrics(filterScope(t0, t1));
                if (tm.salesTotal >= dailyTarget && !flags[dKey]) {
                    flags[dKey] = true; lsSet(LS.celebrated, flags);
                    fireConfetti({ count: 60, origin: 0.25, spread: 0.7 });
                    if (typeof showSettingsToast === 'function') showSettingsToast('本日の売上目標を達成しました！🎉');
                }
            }
        } catch (e) { console.warn('celebrate error', e); }
    }

    function showCelebrationBanner(title, sub) {
        const old = $('celebration-banner');
        if (old) old.remove();
        const div = document.createElement('div');
        div.id = 'celebration-banner';
        div.className = 'celebration-banner';
        div.innerHTML = `<div class="celebration-inner">
            <span class="text-2xl">🏆</span>
            <div><p class="font-bold">${title}</p><p class="text-xs opacity-90">${sub}</p></div>
            <button onclick="this.closest('.celebration-banner').remove()" aria-label="閉じる" class="ml-2 opacity-70 hover:opacity-100">✕</button>
        </div>`;
        document.body.appendChild(div);
        setTimeout(() => div.classList.add('show'), 30);
        setTimeout(() => { div.classList.remove('show'); setTimeout(() => div.remove(), 500); }, 8000);
    }

    /* ================= LIVE 自動更新 ================= */
    let lastFetchAt = null;
    let liveTimer = null;

    function setLiveIndicator(state, time) {
        const el = $('live-indicator');
        if (!el) return;
        const t = time || lastFetchAt;
        const timeStr = t ? new Date(t).toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }) : '--:--';
        el.classList.remove('hidden');
        el.classList.toggle('syncing', state === 'syncing');
        el.classList.toggle('error', state === 'error');
        el.querySelector('.live-time').textContent = state === 'syncing' ? '更新中…' : `${timeStr} 更新`;
    }

    async function silentRefresh(reason) {
        if (document.hidden) return;
        const activeTab = document.querySelector('.tab-content:not(.hidden)')?.id || '';
        if (activeTab === 'content-edit' || (typeof changedRows !== "undefined" && changedRows.size > 0)) return; // 編集中は触らない
        setLiveIndicator('syncing');
        window.__suppressLoadErrorBanner = true;
        try {
            const prevTotal = (Array.isArray(rawData) ? rawData : []).length;
            const prevSig = JSON.stringify((Array.isArray(rawData) ? rawData.slice(-3) : []));
            const ok = await loadDataFromSpreadsheet();
            if (ok) {
                lastFetchAt = Date.now();
                updateDashboard();
                const newSig = JSON.stringify(rawData.slice(-3));
                if (rawData.length !== prevTotal || newSig !== prevSig) {
                    flashKpis();
                }
                setLiveIndicator('ok');
            } else {
                setLiveIndicator('error');
            }
        } catch (e) {
            setLiveIndicator('error');
        } finally {
            window.__suppressLoadErrorBanner = false;
        }
    }

    function flashKpis() {
        document.querySelectorAll('#kpi-grid .kpi-card').forEach((c, i) => {
            setTimeout(() => {
                c.classList.remove('kpi-flash');
                void c.offsetWidth;
                c.classList.add('kpi-flash');
            }, i * 80);
        });
    }

    let realtimeDebounce = null;
    function startLive() {
        if (liveTimer) clearInterval(liveTimer);
        liveTimer = setInterval(() => silentRefresh('interval'), LIVE_INTERVAL_MS);
        document.addEventListener('visibilitychange', () => {
            if (!document.hidden && lastFetchAt && Date.now() - lastFetchAt > LIVE_INTERVAL_MS) {
                silentRefresh('visible');
            }
        });
        // Supabaseモード: Realtimeで日報の追加・修正を数秒で反映
        if (window.Backend && Backend.mode() === 'supabase') {
            Backend.initRealtime(() => {
                clearTimeout(realtimeDebounce);
                realtimeDebounce = setTimeout(() => silentRefresh('realtime'), 1200);
            });
        }
    }

    /* ================= 前年比 (YoY) ================= */
    function renderYoY(metrics) {
        const ids = { sales: 'yoy-sales', customers: 'yoy-customers' };
        const ym = selectedYM();
        const [y, m] = ym.split('/').map(Number);
        if (!y || !m || currentPeriodFilter !== 'month') {
            Object.values(ids).forEach(id => { const el = $(id); if (el) el.textContent = ''; });
            return;
        }
        const start = new Date(y - 1, m - 1, 1);
        const end = new Date(y - 1, m, 0, 23, 59, 59);
        const prev = calculateMetrics(filterScope(start, end));
        const set = (id, cur, pre) => {
            const el = $(id);
            if (!el) return;
            if (!pre) { el.textContent = ''; return; }
            const diff = (cur - pre) / pre * 100;
            const cls = diff >= 0 ? 'text-sage-600 dark:text-sage-400' : 'text-rose-500';
            el.innerHTML = `<span class="${cls}">前年比 ${diff >= 0 ? '+' : ''}${Math.abs(diff) >= 10 ? Math.round(diff) : diff.toFixed(1)}%</span>`;
        };
        set(ids.sales, metrics.salesTotal, prev.salesTotal);
        set(ids.customers, metrics.customersTotal, prev.customersTotal);
    }

    /* ================= 今週のハイライト ================= */
    function weekWindow(offsetWeeks = 0) {
        const now = new Date();
        const end = new Date(now.getFullYear(), now.getMonth(), now.getDate() - offsetWeeks * 7, 23, 59, 59);
        const start = new Date(end.getFullYear(), end.getMonth(), end.getDate() - 6);
        return [start, end];
    }

    function staffListInScope() {
        const st = currentStore();
        if (st === 'all') {
            const list = [];
            Object.entries(STAFF_ROSTER).forEach(([store, names]) => names.forEach(n => list.push({ store, name: n })));
            return list;
        }
        return (STAFF_ROSTER[st] || []).map(n => ({ store: st, name: n }));
    }

    function computeHighlights(goal) {
        const cards = [];
        const [curStart, curEnd] = weekWindow(0);
        const [prevStart, prevEnd] = weekWindow(1);
        const staffs = staffListInScope();

        // --- 週間MVP（伸び率） & 売上王 ---
        let mvp = null, king = null;
        staffs.forEach(({ store, name }) => {
            const cur = calculateMetrics(filterScope(curStart, curEnd, { store, staff: name })).salesTotal;
            const prev = calculateMetrics(filterScope(prevStart, prevEnd, { store, staff: name })).salesTotal;
            if (!king || cur > king.cur) king = { store, name, cur };
            if (prev >= 10000 && cur > prev) {
                const growth = (cur - prev) / prev * 100;
                if (!mvp || growth > mvp.growth) mvp = { store, name, cur, prev, growth };
            }
        });
        if (mvp && mvp.cur >= 30000) {
            cards.push({ type: 'gold', icon: '👑', title: `今週のMVP: ${esc(mvp.name)}`, body: `直近7日 ${yen(mvp.cur)}（前週比 <b>+${Math.round(mvp.growth)}%</b>）${storeName(mvp.store)}` });
        }
        if (king && king.cur > 0 && (!mvp || king.name !== mvp.name)) {
            cards.push({ type: 'primary', icon: '🏅', title: `週間売上トップ: ${esc(king.name)}`, body: `直近7日 ${yen(king.cur)} ${storeName(king.store)}` });
        }

        // --- 自己ベスト更新（日次売上） ---
        try {
            const bestFlags = lsGet(LS.bestCelebrated, {});
            staffs.forEach(({ store, name }) => {
                const recs = (rawData || []).filter(d => d && d.date && d.store === store && d.staff?.toLowerCase() === name.toLowerCase());
                if (recs.length < 8) return;
                const dailyMap = {};
                recs.forEach(d => {
                    const s = d.sales || {}, dc = d.discounts || {};
                    const v = (s.cash || 0) + (s.credit || 0) + (s.qr || 0) + (dc.hpbPoints || 0) + (dc.hpbGift || 0);
                    dailyMap[d.date] = (dailyMap[d.date] || 0) + v;
                });
                const entries = Object.entries(dailyMap).map(([date, v]) => ({ date, v, t: parseDate(date).getTime() }));
                const cutoff = curStart.getTime();
                const recent = entries.filter(e => e.t >= cutoff);
                const past = entries.filter(e => e.t < cutoff);
                if (!recent.length || past.length < 5) return;
                const recentMax = recent.reduce((a, b) => (b.v > a.v ? b : a));
                const pastMax = Math.max(...past.map(e => e.v));
                if (recentMax.v > pastMax && recentMax.v > 0) {
                    cards.push({ type: 'sage', icon: '🚀', title: `自己ベスト更新: ${esc(name)}`, body: `${recentMax.date} に日次売上 <b>${yen(recentMax.v)}</b>（これまでの最高 ${yen(pastMax)}）` });
                    const fKey = `${store}:${name}:${recentMax.date}`;
                    if (!bestFlags[fKey] && recentMax.date === todayKey()) {
                        bestFlags[fKey] = true; lsSet(LS.bestCelebrated, bestFlags);
                        fireConfetti({ count: 70, origin: 0.3, spread: 0.8 });
                    }
                }
            });
        } catch (e) { console.warn('best record error', e); }

        // --- 異常検知（現在スコープ） ---
        try {
            const last7 = calculateMetrics(filterScope(...weekWindow(0)));
            const prior28Start = new Date(); prior28Start.setDate(prior28Start.getDate() - 35);
            const prior28End = new Date(); prior28End.setDate(prior28End.getDate() - 7);
            const prior = calculateMetrics(filterScope(prior28Start, prior28End));
            const d7 = Object.keys(last7.daily || {}).length;
            const dP = Object.keys(prior.daily || {}).length;
            if (d7 >= 3 && dP >= 6) {
                const avg7 = last7.salesTotal / d7;
                const avgP = prior.salesTotal / dP;
                if (avgP > 0) {
                    const diff = (avg7 - avgP) / avgP * 100;
                    if (diff <= -25) cards.push({ type: 'rose', icon: '⚠️', title: '売上ペースが低下しています', body: `直近7日の日平均 ${yen(avg7)}（過去4週平均比 <b>${Math.round(diff)}%</b>）。要因を確認しましょう。` });
                    else if (diff >= 25) cards.push({ type: 'sage', icon: '📈', title: '売上ペースが好調です', body: `直近7日の日平均 ${yen(avg7)}（過去4週平均比 <b>+${Math.round(diff)}%</b>）` });
                }
            }
            // 新規次回予約率: 今月 vs 先月
            const now = new Date();
            const cm = calculateMetrics(filterScope(new Date(now.getFullYear(), now.getMonth(), 1), now));
            const pm = calculateMetrics(filterScope(new Date(now.getFullYear(), now.getMonth() - 1, 1), new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59)));
            if (cm.customersNew >= 5 && pm.customersNew >= 5) {
                const cr = (cm.nextRes.hpbNew + cm.nextRes.mininaiNew) / cm.customersNew * 100;
                const pr = (pm.nextRes.hpbNew + pm.nextRes.mininaiNew) / pm.customersNew * 100;
                if (cr - pr <= -10) cards.push({ type: 'rose', icon: '📉', title: '新規の次回予約率が低下', body: `今月 ${cr.toFixed(0)}%（先月 ${pr.toFixed(0)}%）。クロージングトークを見直すタイミングです。` });
            }
            // ロス率
            if (cm.salesTotal > 0) {
                const lossRate = cm.lossTotal / cm.salesTotal * 100;
                if (lossRate >= 8) cards.push({ type: 'rose', icon: '🧾', title: '割引・返金ロスが高水準', body: `今月のロス率 <b>${lossRate.toFixed(1)}%</b>（${yen(cm.lossTotal)}）` });
            }
        } catch (e) { console.warn('anomaly error', e); }

        // --- 月間ストリーク（連続目標達成） ---
        try {
            let bestStreak = null;
            staffs.forEach(({ store, name }) => {
                const streak = calcMonthlyGoalStreak(store, name);
                if (streak >= 2 && (!bestStreak || streak > bestStreak.streak)) bestStreak = { store, name, streak };
            });
            if (bestStreak) {
                cards.push({ type: 'gold', icon: '🔥', title: `${esc(bestStreak.name)}: 目標達成 ${bestStreak.streak}ヶ月連続中`, body: `${storeName(bestStreak.store)} のストリーク記録です。継続に注目！` });
            }
        } catch (e) { console.warn('streak error', e); }

        return cards.slice(0, 6);
    }

    // 月間売上目標の連続達成数（当月は達成済みのときのみカウント）
    function calcMonthlyGoalStreak(store, staff) {
        let streak = 0;
        const now = new Date();
        for (let i = 0; i < 24; i++) {
            const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
            const ym = `${d.getFullYear()}/${d.getMonth() + 1}`;
            const g = getStaffGoal(store, staff, ym);
            const target = goalMonthlyTotal(g);
            if (target <= 0) { if (i === 0) continue; break; }
            const end = new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59);
            const m = calculateMetrics(filterScope(d, end, { store, staff }));
            if (m.salesTotal >= target) streak++;
            else { if (i === 0) continue; break; } // 当月未達はスキップ（進行中のため）
        }
        return streak;
    }

    // ブログ更新目標（10件/月）の連続達成数
    function calcBlogStreak(store, staff) {
        let streak = 0;
        const now = new Date();
        for (let i = 0; i < 24; i++) {
            const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
            const end = new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59);
            const m = calculateMetrics(filterScope(d, end, { store, staff }));
            if (m.blogUpdatesTotal >= 10) streak++;
            else { if (i === 0) continue; break; }
        }
        return streak;
    }

    function renderHighlights(goal) {
        const section = $('highlights-section');
        const grid = $('highlights-grid');
        if (!section || !grid) return;
        const cards = computeHighlights(goal);
        if (!cards.length) { section.classList.add('hidden'); return; }
        section.classList.remove('hidden');
        grid.innerHTML = cards.map((c, i) => `
            <div class="highlight-card hl-${c.type}" style="animation-delay:${i * 70}ms">
                <span class="hl-icon">${c.icon}</span>
                <div class="min-w-0">
                    <p class="hl-title">${c.title}</p>
                    <p class="hl-body">${c.body}</p>
                </div>
            </div>`).join('');
    }

    /* ================= 店舗対抗レース & ベンチマーク ================= */
    function renderStoreRace() {
        const section = $('store-race-section');
        if (!section) return;
        if (currentStore() !== 'all' || !isAdminView()) { section.classList.add('hidden'); return; }
        const stores = Object.keys(STAFF_ROSTER);
        if (stores.length < 2) { section.classList.add('hidden'); return; }
        section.classList.remove('hidden');

        const now = new Date();
        const mStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const colors = { chiba: '#b8956a', honatsugi: '#739977', yamato: '#b08f8a' };
        const rows = stores.map(st => {
            const m = calculateMetrics(filterScope(mStart, now, { store: st, staff: 'all' }));
            const g = getStoreAggregateGoal(st, monthKey());
            const target = goalMonthlyTotal(g);
            const pct = target > 0 ? m.salesTotal / target * 100 : 0;
            const unit = m.customersTotal > 0 ? m.salesTotal / m.customersTotal : 0;
            const newRes = m.customersNew > 0 ? (m.nextRes.hpbNew + m.nextRes.mininaiNew) / m.customersNew * 100 : 0;
            const totRes = m.customersTotal > 0 ? m.nextRes.total / m.customersTotal * 100 : 0;
            return { st, m, target, pct, unit, newRes, totRes, color: colors[st] || '#566882' };
        }).sort((a, b) => b.pct - a.pct);

        const raceEl = $('store-race-bars');
        if (raceEl) {
            raceEl.innerHTML = rows.map((r, i) => `
                <div class="race-row">
                    <div class="race-head">
                        <span class="race-rank">${i === 0 ? '🥇' : i === 1 ? '🥈' : '🥉'}</span>
                        <span class="race-name">${storeName(r.st)}</span>
                        <span class="race-pct" style="color:${r.color}">${Math.round(r.pct)}%</span>
                        <span class="race-detail">${yen(r.m.salesTotal)}${r.target > 0 ? ` / ${yen(r.target)}` : ''}</span>
                    </div>
                    <div class="race-track">
                        <div class="race-fill" style="width:${Math.min(r.pct, 100)}%;background:linear-gradient(90deg, ${r.color}cc, ${r.color});">
                            <span class="race-runner">🏃‍♀️</span>
                        </div>
                        <div class="race-goalline"></div>
                    </div>
                </div>`).join('');
        }

        const tbody = $('store-benchmark-body');
        if (tbody) {
            tbody.innerHTML = rows.map(r => `
                <tr class="border-b border-surface-100 dark:border-accent-700">
                    <td class="py-2.5 px-3 font-semibold text-accent-800 dark:text-surface-100"><span class="inline-block w-2.5 h-2.5 rounded-full mr-2" style="background:${r.color}"></span>${storeName(r.st)}</td>
                    <td class="py-2.5 px-3 text-right">${yen(r.m.salesTotal)}</td>
                    <td class="py-2.5 px-3 text-right font-semibold" style="color:${r.color}">${Math.round(r.pct)}%</td>
                    <td class="py-2.5 px-3 text-right">${yen(r.unit)}</td>
                    <td class="py-2.5 px-3 text-right">${r.m.customersNew}名</td>
                    <td class="py-2.5 px-3 text-right">${r.newRes.toFixed(0)}%</td>
                    <td class="py-2.5 px-3 text-right">${r.totRes.toFixed(0)}%</td>
                </tr>`).join('');
        }

        // レーダー（各指標を店舗間最大値で正規化）
        const canvas = $('storeRadarChart');
        if (canvas && window.Chart) {
            const axes = ['売上', '達成率', '客単価', '新規数', '新規予約率', '総予約率'];
            const maxes = {
                sales: Math.max(...rows.map(r => r.m.salesTotal), 1),
                pct: Math.max(...rows.map(r => r.pct), 1),
                unit: Math.max(...rows.map(r => r.unit), 1),
                nu: Math.max(...rows.map(r => r.m.customersNew), 1),
                nr: Math.max(...rows.map(r => r.newRes), 1),
                tr: Math.max(...rows.map(r => r.totRes), 1),
            };
            const datasets = rows.map(r => ({
                label: storeName(r.st),
                data: [r.m.salesTotal / maxes.sales, r.pct / maxes.pct, r.unit / maxes.unit, r.m.customersNew / maxes.nu, r.newRes / maxes.nr, r.totRes / maxes.tr].map(v => Math.round(v * 100)),
                borderColor: r.color,
                backgroundColor: r.color + '22',
                pointBackgroundColor: r.color,
                borderWidth: 2,
            }));
            if (!charts.storeRadar) {
                const t = chartTheme();
                charts.storeRadar = new Chart(canvas, {
                    type: 'radar',
                    data: { labels: axes, datasets },
                    options: {
                        ...chartCommonOptions(),
                        scales: { r: { min: 0, max: 100, ticks: { display: false }, grid: { color: t.grid }, angleLines: { color: t.grid }, pointLabels: { color: t.text, font: { size: 11, family: 'Inter, sans-serif' } } } },
                    }
                });
            } else {
                applyChartData(charts.storeRadar, { labels: axes, datasets });
            }
        }
    }

    /* ================= 曜日別分析（KPIタブ） ================= */
    function renderWeekdayChart(metrics) {
        const canvas = $('weekdayChart');
        if (!canvas || !window.Chart) return;
        const daily = metrics.daily || {};
        const byDow = Array.from({ length: 7 }, () => ({ sales: [], cust: [] }));
        Object.keys(daily).forEach(k => {
            const dow = parseDate(k).getDay();
            byDow[dow].sales.push(daily[k].sales || 0);
            byDow[dow].cust.push(daily[k].customers || 0);
        });
        const order = [1, 2, 3, 4, 5, 6, 0]; // 月→日
        const labels = ['月', '火', '水', '木', '金', '土', '日'];
        const avg = arr => arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0;
        const salesAvg = order.map(d => Math.round(avg(byDow[d].sales)));
        const custAvg = order.map(d => +avg(byDow[d].cust).toFixed(1));
        const maxIdx = salesAvg.indexOf(Math.max(...salesAvg));

        const data = {
            labels,
            datasets: [
                {
                    type: 'bar', label: '平均売上', data: salesAvg, yAxisID: 'y',
                    backgroundColor: ctx => makeVGradient(ctx, ctx.dataIndex === maxIdx ? '#c9a96e' : '#dcc9b3', ctx.dataIndex === maxIdx ? '#b8956a' : '#cdb594'),
                    borderRadius: 6, borderSkipped: false, maxBarThickness: 34,
                },
                { type: 'line', label: '平均来店数', data: custAvg, yAxisID: 'y1', borderColor: '#566882', backgroundColor: '#566882', tension: 0.35, pointRadius: 3 },
            ]
        };
        if (!charts.weekday) {
            const t = chartTheme();
            charts.weekday = new Chart(canvas, {
                type: 'bar',
                data,
                options: {
                    ...chartCommonOptions(),
                    scales: {
                        x: { grid: { display: false }, ticks: { color: t.textMuted } },
                        y: { position: 'left', grid: { color: t.grid }, ticks: { color: t.textMuted, callback: v => '¥' + Number(v).toLocaleString() } },
                        y1: { position: 'right', grid: { display: false }, ticks: { color: t.textMuted, precision: 0 } },
                    }
                }
            });
        } else {
            applyChartData(charts.weekday, data);
        }
        const best = $('weekday-best');
        if (best) {
            const has = salesAvg.some(v => v > 0);
            best.innerHTML = has
                ? `ベスト曜日: <b class="text-primary-600">${labels[maxIdx]}曜日</b>（平均 ${yen(salesAvg[maxIdx])}）／ 最も静かな曜日: ${labels[salesAvg.indexOf(Math.min(...salesAvg.filter(v => v > 0)))] ?? '-'}曜日`
                : 'データが集まると曜日傾向が表示されます';
        }
    }

    /* ================= スタッフスキルレーダー & ストリーク ================= */
    function renderStaffRadar() {
        const canvas = $('staffSkillRadar');
        if (!canvas || !window.Chart) return;
        const staff = lockedStaff || (currentStaff() !== 'all' ? currentStaff() : null);
        const store = lockedStore || (currentStore() !== 'all' ? currentStore() : null);
        if (!staff || !store) return;

        const ym = selectedYM();
        const [y, m] = ym.split('/').map(Number);
        const start = new Date(y, m - 1, 1);
        const end = new Date(y, m, 0, 23, 59, 59);

        const roster = STAFF_ROSTER[store] || [];
        const all = roster.map(name => {
            const mm = calculateMetrics(filterScope(start, end, { store, staff: name }));
            return {
                name,
                sales: mm.salesTotal,
                nu: mm.customersNew,
                unit: mm.customersTotal > 0 ? mm.salesTotal / mm.customersTotal : 0,
                res: mm.customersTotal > 0 ? mm.nextRes.total / mm.customersTotal * 100 : 0,
                rev: mm.reviews5StarTotal,
            };
        });
        const mine = all.find(a => a.name.toLowerCase() === staff.toLowerCase());
        if (!mine) return;
        const maxOf = k => Math.max(...all.map(a => a[k]), 0.001);
        const norm = a => ['sales', 'nu', 'unit', 'res', 'rev'].map(k => Math.round(a[k] / maxOf(k) * 100));
        const avg = key => all.reduce((s, a) => s + a[key], 0) / (all.length || 1);
        const storeAvg = { sales: avg('sales'), nu: avg('nu'), unit: avg('unit'), res: avg('res'), rev: avg('rev') };

        const labels = ['売上', '新規獲得', '客単価', '次回予約率', '★5口コミ'];
        const data = {
            labels,
            datasets: [
                { label: 'あなた', data: norm(mine), borderColor: '#b8956a', backgroundColor: 'rgba(184,149,106,0.18)', pointBackgroundColor: '#b8956a', borderWidth: 2.5 },
                { label: '店舗平均', data: norm(storeAvg), borderColor: '#8f9fb5', backgroundColor: 'rgba(143,159,181,0.10)', pointBackgroundColor: '#8f9fb5', borderWidth: 1.5, borderDash: [5, 4] },
            ]
        };
        if (!charts.staffRadar) {
            const t = chartTheme();
            charts.staffRadar = new Chart(canvas, {
                type: 'radar',
                data,
                options: {
                    ...chartCommonOptions(),
                    scales: { r: { min: 0, max: 100, ticks: { display: false }, grid: { color: t.grid }, angleLines: { color: t.grid }, pointLabels: { color: t.text, font: { size: 11, family: 'Inter, sans-serif' } } } },
                }
            });
        } else {
            applyChartData(charts.staffRadar, data);
        }
    }

    function renderStaffStreaks() {
        const el = $('staff-streaks');
        if (!el) return;
        const staff = lockedStaff || (currentStaff() !== 'all' ? currentStaff() : null);
        const store = lockedStore || (currentStore() !== 'all' ? currentStore() : null);
        if (!staff || !store) { el.innerHTML = ''; return; }
        const goalStreak = calcMonthlyGoalStreak(store, staff);
        const blogStreak = calcBlogStreak(store, staff);
        const chips = [];
        if (goalStreak >= 1) chips.push(`<span class="streak-chip">🔥 売上目標 <b>${goalStreak}ヶ月</b>連続達成</span>`);
        if (blogStreak >= 1) chips.push(`<span class="streak-chip blue">📝 ブログ10件 <b>${blogStreak}ヶ月</b>連続</span>`);
        // 今日の日報入力済みチェック
        const today = todayKey();
        const hasToday = (rawData || []).some(d => d && d.date === today && d.store === store && d.staff?.toLowerCase() === staff.toLowerCase());
        chips.push(hasToday
            ? '<span class="streak-chip green">✓ 本日の日報 入力済み</span>'
            : '<span class="streak-chip gray">⏳ 本日の日報 未入力</span>');
        el.innerHTML = chips.join('');
    }

    /* ================= AIコーチ（個人向け・週次自動） ================= */
    function aiCache() { return lsGet(LS.aiCache, {}); }

    async function callGemini(prompt) {
        // Supabaseモード: Edge Function 中継（APIキーをブラウザに置かない）
        if (window.Backend && Backend.mode() === 'supabase') {
            try {
                return await Backend.aiProxy(prompt);
            } catch (e) {
                console.warn('Edge Function中継に失敗。ローカルキーでフォールバック:', e.message);
            }
        }
        const apiKey = typeof loadGeminiApiKey === 'function' ? loadGeminiApiKey() : null;
        if (!apiKey) throw new Error('Gemini APIキーが未設定です（設定タブで登録するか、Edge Functionをデプロイしてください）');
        const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${window.GEMINI_MODEL}:generateContent?key=${apiKey}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
        });
        const data = await res.json();
        const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
        if (!text) throw new Error(data?.error?.message || 'AIからの応答が不正です');
        return text;
    }

    function mdLite(text) {
        return esc(text)
            .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
            .replace(/^[\-・\*] (.+)$/gm, '<span class="ai-bullet">$1</span>')
            .replace(/\n/g, '<br>');
    }

    function personalPrompt(staff, store) {
        const m = calculateMetrics(getFilteredData());
        const g = getStaffGoal(store, staff);
        const goal = goalMonthlyTotal(g);
        const unit = m.customersTotal > 0 ? Math.round(m.salesTotal / m.customersTotal) : 0;
        const newRes = m.customersNew > 0 ? ((m.nextRes.hpbNew + m.nextRes.mininaiNew) / m.customersNew * 100).toFixed(1) : 0;
        const totRes = m.customersTotal > 0 ? (m.nextRes.total / m.customersTotal * 100).toFixed(1) : 0;
        return `あなたはアイラッシュサロンの優秀な店長兼コーチです。スタッフ「${staff}」さん個人の今月の実績を見て、明日からすぐ実行できる具体的なアドバイスを日本語で簡潔に出してください。

【${staff}さんの今月実績】
- 売上: ¥${m.salesTotal.toLocaleString()} / 目標 ¥${goal.toLocaleString()}（達成率 ${goal > 0 ? Math.round(m.salesTotal / goal * 100) : 0}%）
- 来店: ${m.customersTotal}名（新規${m.customersNew} / 既存${m.customersExisting}）目標 新規${g.newCustomers}・既存${g.existingCustomers}
- 客単価: ¥${unit.toLocaleString()} / 目標 ¥${(g.unitPrice || 0).toLocaleString()}
- 新規次回予約率: ${newRes}% / 目標 ${g.newReservationRate}%
- 総次回予約率: ${totRes}% / 目標 ${g.reservationRate}%
- ★5口コミ: ${m.reviews5StarTotal}件 / 目標 ${g.reviews5Star || 0}件
- ブログ更新: ${m.blogUpdatesTotal}件 / 目標 10件

出力条件:
- 「今週がんばること」3つ（番号付き・各2行以内・接客トーク例も入れる）
- 最後に一言、前向きな応援メッセージ
- 全体で300字程度`;
    }

    async function getPersonalAIAdvice(manual = false) {
        const container = $('staff-ai-content');
        const btn = $('btn-staff-ai');
        if (!container) return;
        const staff = lockedStaff || (currentStaff() !== 'all' ? currentStaff() : null);
        const store = lockedStore || (currentStore() !== 'all' ? currentStore() : null);
        if (!staff || !store) return;

        if (btn) { btn.disabled = true; btn.textContent = 'コーチが考え中…'; }
        container.innerHTML = '<div class="ai-thinking"><span></span><span></span><span></span> AIコーチが実績を分析しています…</div>';
        try {
            const text = await callGemini(personalPrompt(staff, store));
            const cache = aiCache();
            cache[`personal:${store}:${staff}`] = { at: Date.now(), text };
            lsSet(LS.aiCache, cache);
            container.innerHTML = `<div class="ai-advice-body">${mdLite(text)}</div><p class="ai-advice-time">生成: ${new Date().toLocaleString('ja-JP')}</p>`;
        } catch (e) {
            if (manual) container.innerHTML = `<p class="text-sm text-rose-500">${esc(e.message)}</p>`;
            else container.innerHTML = '<p class="text-xs text-surface-500">「アドバイスをもらう」を押すと、AIコーチが今月の実績から具体的なアクションを提案します。</p>';
        } finally {
            if (btn) { btn.disabled = false; btn.textContent = 'アドバイスをもらう'; }
        }
    }
    window.getPersonalAIAdvice = () => getPersonalAIAdvice(true);

    let aiAutoTried = false;
    function maybeAutoAI() {
        const container = $('staff-ai-content');
        if (!container) return;
        const staff = lockedStaff || (currentStaff() !== 'all' ? currentStaff() : null);
        const store = lockedStore || (currentStore() !== 'all' ? currentStore() : null);
        if (!staff || !store) return;
        const cached = aiCache()[`personal:${store}:${staff}`];
        if (cached) {
            container.innerHTML = `<div class="ai-advice-body">${mdLite(cached.text)}</div><p class="ai-advice-time">生成: ${new Date(cached.at).toLocaleString('ja-JP')}（週1回自動更新）</p>`;
        }
        const apiKey = typeof loadGeminiApiKey === 'function' ? loadGeminiApiKey() : null;
        const stale = !cached || Date.now() - cached.at > 7 * 24 * 3600 * 1000;
        if (apiKey && stale && navigator.onLine && !aiAutoTried) {
            aiAutoTried = true;
            getPersonalAIAdvice(false);
        }
    }

    // 管理者向け: 前回のAIアドバイスをキャッシュ表示
    function setupAdminAiCache() {
        if (typeof window.displayAIAdvice !== 'function') return;
        const orig = window.displayAIAdvice;
        window.displayAIAdvice = function (advice, rate) {
            orig(advice, rate);
            const cache = aiCache();
            cache['admin'] = { at: Date.now(), text: advice, rate };
            lsSet(LS.aiCache, cache);
        };
        const cached = aiCache()['admin'];
        if (cached && $('ai-advice-content')) {
            orig(cached.text, cached.rate || 0);
            const div = document.createElement('p');
            div.className = 'text-[10px] text-surface-400 text-right mt-1';
            div.textContent = `前回の分析結果（${new Date(cached.at).toLocaleString('ja-JP')}）を表示中`;
            $('ai-advice-content').appendChild(div);
        }
    }

    /* ================= 順位変動バッジ ================= */
    function rankOrdersNow() {
        try {
            const sm = getStaffMetrics();
            return {
                sales: [...sm].sort((a, b) => b.sales - a.sales).map(s => s.name),
                new: [...sm].sort((a, b) => b.newCustomers - a.newCustomers).map(s => s.name),
                unitPrice: [...sm].filter(s => s.unitPrice > 0).sort((a, b) => b.unitPrice - a.unitPrice).map(s => s.name),
            };
        } catch (e) { return null; }
    }

    function updateRankHistory() {
        const orders = rankOrdersNow();
        if (!orders) return;
        const hist = lsGet(LS.rankHistory, {});
        const today = todayKey();
        ['sales', 'new', 'unitPrice'].forEach(type => {
            const key = `${type}:${currentStore()}:${currentPeriodFilter}`;
            let h = hist[key];
            if (!h) h = { baselineDate: today, baseline: orders[type], last: orders[type] };
            else if (h.baselineDate !== today) { h.baseline = h.last || orders[type]; h.baselineDate = today; }
            h.last = orders[type];
            hist[key] = h;
        });
        lsSet(LS.rankHistory, hist);
    }

    function rankMoveBadge(type, name, curIdx) {
        try {
            const hist = lsGet(LS.rankHistory, {});
            const h = hist[`${type}:${currentStore()}:${currentPeriodFilter}`];
            if (!h || !Array.isArray(h.baseline)) return '';
            const baseIdx = h.baseline.indexOf(name);
            if (baseIdx < 0) return '<span class="rank-move new">NEW</span>';
            const diff = baseIdx - curIdx;
            if (diff > 0) return `<span class="rank-move up">▲${diff}</span>`;
            if (diff < 0) return `<span class="rank-move down">▼${-diff}</span>`;
            return '';
        } catch (e) { return ''; }
    }

    /* ================= フォロー候補（マーケタブ） ================= */
    function renderFollowList() {
        const tbody = $('follow-list-body');
        const countEl = $('follow-list-count');
        if (!tbody) return;
        const list = (typeof customerData !== "undefined" && Array.isArray(customerData)) ? customerData : [];
        const storeFilter = $('marketing-store-filter')?.value || 'all';
        const now = Date.now();
        const DAY = 24 * 3600 * 1000;
        const targets = list.filter(c => {
            if (!c || !c.timestamp) return false;
            if (storeFilter !== 'all' && c.store !== storeFilter) return false;
            const days = (now - parseDate(c.timestamp).getTime()) / DAY;
            return days >= 21 && days <= 60;
        }).map(c => ({ ...c, days: Math.floor((now - parseDate(c.timestamp).getTime()) / DAY) }))
          .sort((a, b) => a.days - b.days);

        if (countEl) countEl.textContent = `${targets.length}名`;
        if (!targets.length) {
            tbody.innerHTML = '<tr><td colspan="4" class="text-center py-6 text-surface-400 text-sm">該当者なし（来店から3〜8週間の顧客が表示されます）</td></tr>';
            return;
        }
        tbody.innerHTML = targets.slice(0, 30).map(c => {
            const urgency = c.days >= 42 ? 'rose' : c.days >= 30 ? 'amber' : 'sage';
            return `<tr class="border-b border-surface-100 dark:border-accent-700">
                <td class="py-2.5 px-3 font-medium text-accent-800 dark:text-surface-100">${esc(c.name || '-')}</td>
                <td class="py-2.5 px-3 text-surface-600 dark:text-surface-300">${storeName(c.store)}</td>
                <td class="py-2.5 px-3 text-surface-600 dark:text-surface-300">${parseDate(c.timestamp).toLocaleDateString('ja-JP')}</td>
                <td class="py-2.5 px-3 text-right"><span class="follow-badge ${urgency}">${c.days}日経過</span></td>
            </tr>`;
        }).join('');
    }

    /* ================= 媒体ROI（広告費→CPA） ================= */
    function adCostsData() { return lsGet(LS.adCosts, {}); }

    function renderRoi() {
        const wrap = $('roi-content');
        if (!wrap) return;
        const ym = selectedYM();
        const all = adCostsData();
        const costs = all[ym] || {};
        const stores = Object.keys(STAFF_ROSTER);
        const [y, m] = ym.split('/').map(Number);
        const start = new Date(y, m - 1, 1), end = new Date(y, m, 0, 23, 59, 59);

        let totalCost = 0, totalHpbCost = 0, totalMiniCost = 0, totalNew = 0, totalHpbNew = 0, totalMiniNew = 0;
        const rows = stores.map(st => {
            const c = costs[st] || { hpb: 0, minimo: 0 };
            const mm = calculateMetrics(filterScope(start, end, { store: st, staff: 'all' }));
            const hpbNew = mm.newByChannel.hpb, miniNew = mm.newByChannel.mininai;
            totalHpbCost += c.hpb || 0; totalMiniCost += c.minimo || 0; totalCost += (c.hpb || 0) + (c.minimo || 0);
            totalHpbNew += hpbNew; totalMiniNew += miniNew; totalNew += hpbNew + miniNew;
            return { st, c, hpbNew, miniNew };
        });

        const cpa = (cost, n) => (cost > 0 && n > 0) ? yen(cost / n) : (cost > 0 ? '獲得0名' : '—');
        wrap.innerHTML = `
            <div class="overflow-x-auto">
            <table class="w-full text-sm">
                <thead><tr class="border-b border-surface-200 dark:border-accent-700 text-surface-500">
                    <th class="text-left py-2 px-3 font-semibold">店舗</th>
                    <th class="text-right py-2 px-3 font-semibold">HPB費用</th>
                    <th class="text-right py-2 px-3 font-semibold">minimo等費用</th>
                    <th class="text-right py-2 px-3 font-semibold">HPB新規</th>
                    <th class="text-right py-2 px-3 font-semibold">minimo新規</th>
                    <th class="text-right py-2 px-3 font-semibold text-primary-600">HPB CPA</th>
                    <th class="text-right py-2 px-3 font-semibold text-primary-600">minimo CPA</th>
                </tr></thead>
                <tbody>
                ${rows.map(r => `<tr class="border-b border-surface-100 dark:border-accent-700">
                    <td class="py-2 px-3 font-medium text-accent-800 dark:text-surface-100">${storeName(r.st)}</td>
                    <td class="py-2 px-3 text-right"><input type="text" inputmode="numeric" class="roi-input" data-store="${r.st}" data-channel="hpb" value="${(r.c.hpb || 0).toLocaleString()}"></td>
                    <td class="py-2 px-3 text-right"><input type="text" inputmode="numeric" class="roi-input" data-store="${r.st}" data-channel="minimo" value="${(r.c.minimo || 0).toLocaleString()}"></td>
                    <td class="py-2 px-3 text-right">${r.hpbNew}名</td>
                    <td class="py-2 px-3 text-right">${r.miniNew}名</td>
                    <td class="py-2 px-3 text-right font-semibold">${cpa(r.c.hpb || 0, r.hpbNew)}</td>
                    <td class="py-2 px-3 text-right font-semibold">${cpa(r.c.minimo || 0, r.miniNew)}</td>
                </tr>`).join('')}
                </tbody>
                <tfoot><tr class="bg-surface-50 dark:bg-accent-800/40 font-semibold">
                    <td class="py-2 px-3">合計</td>
                    <td class="py-2 px-3 text-right">${yen(totalHpbCost)}</td>
                    <td class="py-2 px-3 text-right">${yen(totalMiniCost)}</td>
                    <td class="py-2 px-3 text-right">${totalHpbNew}名</td>
                    <td class="py-2 px-3 text-right">${totalMiniNew}名</td>
                    <td class="py-2 px-3 text-right text-primary-600">${cpa(totalHpbCost, totalHpbNew)}</td>
                    <td class="py-2 px-3 text-right text-primary-600">${cpa(totalMiniCost, totalMiniNew)}</td>
                </tr></tfoot>
            </table>
            </div>
            <div class="flex items-center justify-between mt-3 flex-wrap gap-2">
                <p class="text-xs text-surface-500">${ym} の広告費を入力すると新規1名あたりの獲得単価（CPA）を自動計算します${totalCost > 0 && totalNew > 0 ? ` ／ 全体CPA: <b class="text-primary-600">${yen(totalCost / totalNew)}</b>` : ''}</p>
                <button onclick="Enhance.saveRoi()" class="btn-primary py-1.5 px-4 text-sm">広告費を保存</button>
            </div>`;
        wrap.querySelectorAll('.roi-input').forEach(inp => {
            inp.addEventListener('blur', () => { inp.value = (parseInt(inp.value.replace(/[^0-9]/g, '')) || 0).toLocaleString(); });
        });
    }

    function saveRoi() {
        const ym = selectedYM();
        const all = adCostsData();
        const cur = all[ym] || {};
        document.querySelectorAll('#roi-content .roi-input').forEach(inp => {
            const st = inp.dataset.store, ch = inp.dataset.channel;
            cur[st] = cur[st] || {};
            cur[st][ch] = parseInt(inp.value.replace(/[^0-9]/g, '')) || 0;
        });
        all[ym] = cur;
        lsSet(LS.adCosts, all);
        if (typeof saveSettingsToSpreadsheet === 'function') saveSettingsToSpreadsheet(false);
        if (typeof showSettingsToast === 'function') showSettingsToast('広告費を保存しました');
        renderRoi();
    }

    /* ================= 月締め確定（インセンティブ） ================= */
    function monthlyCloseData() { return lsGet(LS.monthlyClose, {}); }

    function collectIncentiveRows(ym) {
        const stores = Object.keys(STAFF_ROSTER);
        const rows = [];
        stores.forEach(store => {
            (STAFF_ROSTER[store] || []).forEach(name => {
                const staffData = (rawData || []).filter(d => {
                    if (!d || !d.date) return false;
                    const p = d.date.split('/');
                    if (p.length < 2) return false;
                    if (`${p[0]}/${parseInt(p[1])}` !== ym) return false;
                    return d.store === store && d.staff === name;
                });
                const inc = calculateIncentive(staffData, store, name);
                rows.push({
                    store, name,
                    baseSalary: Math.round(inc.baseSalary),
                    serviceSales: Math.round(inc.serviceSalesTaxExcl),
                    serviceIncentive: Math.round(inc.serviceIncentive),
                    retailSales: Math.round(inc.retailSalesTaxExcl),
                    retailIncentive: Math.round(inc.retailIncentive),
                    total: Math.round(inc.totalIncentive),
                });
            });
        });
        return rows;
    }

    function renderMonthlyClose() {
        const bar = $('monthly-close-bar');
        if (!bar) return;
        if (!isAdminView()) { bar.classList.add('hidden'); return; }
        bar.classList.remove('hidden');
        const ym = selectedYM();
        const closed = monthlyCloseData()[ym];
        if (closed) {
            const total = closed.rows.reduce((s, r) => s + r.total, 0);
            bar.innerHTML = `
                <div class="close-banner closed">
                    <div class="flex items-center gap-3 flex-wrap">
                        <span class="close-badge">✓ ${ym} 締め済み</span>
                        <span class="text-sm">確定日時: ${new Date(closed.closedAt).toLocaleString('ja-JP')} ／ 確定報酬総額 <b>${yen(total)}</b></span>
                    </div>
                    <div class="flex gap-2">
                        <button onclick="Enhance.toggleCloseDetail()" class="btn-secondary py-1.5 px-3 text-xs">確定明細を表示</button>
                        <button onclick="Enhance.reopenMonth()" class="text-xs text-rose-500 underline px-2">再オープン</button>
                    </div>
                </div>
                <div id="close-detail" class="hidden mt-3 overflow-x-auto premium-card p-4">
                    <p class="text-xs text-surface-500 mb-2">※確定時点のスナップショットです。以後データを修正しても、この明細は変わりません。</p>
                    <table class="w-full text-sm">
                        <thead><tr class="border-b-2 border-[#b8956a] text-surface-500">
                            <th class="text-left py-2 px-3">店舗</th><th class="text-left py-2 px-3">スタッフ</th>
                            <th class="text-right py-2 px-3">基本給</th><th class="text-right py-2 px-3">施術手当</th>
                            <th class="text-right py-2 px-3">物販手当</th><th class="text-right py-2 px-3">確定報酬</th>
                        </tr></thead>
                        <tbody>${closed.rows.map(r => `<tr class="border-b border-surface-100 dark:border-accent-700">
                            <td class="py-2 px-3">${storeName(r.store)}</td>
                            <td class="py-2 px-3 font-medium">${esc(r.name)}</td>
                            <td class="py-2 px-3 text-right">${yen(r.baseSalary)}</td>
                            <td class="py-2 px-3 text-right">${yen(r.serviceIncentive)}</td>
                            <td class="py-2 px-3 text-right">${yen(r.retailIncentive)}</td>
                            <td class="py-2 px-3 text-right font-bold text-[#b8956a]">${yen(r.total)}</td>
                        </tr>`).join('')}</tbody>
                    </table>
                </div>`;
        } else {
            bar.innerHTML = `
                <div class="close-banner open">
                    <p class="text-sm flex items-center gap-2"><i data-lucide="lock-open" class="w-4 h-4"></i>${ym} は未確定です。月末締め後に確定すると、その時点の明細がスナップショット保存されます。</p>
                    <button onclick="Enhance.closeMonth()" class="btn-gold py-1.5 px-4 text-sm font-bold">この月を締める（確定）</button>
                </div>`;
        }
        if (window.lucide) lucide.createIcons();
    }

    function closeMonth() {
        const ym = selectedYM();
        if (!confirm(`${ym} のインセンティブを確定します。\n現在の計算結果がスナップショット保存され、給与明細の根拠として固定されます。よろしいですか？`)) return;
        const rows = collectIncentiveRows(ym);
        const all = monthlyCloseData();
        all[ym] = { closedAt: Date.now(), rows };
        lsSet(LS.monthlyClose, all);
        if (typeof saveSettingsToSpreadsheet === 'function') saveSettingsToSpreadsheet(false);
        if (typeof showSettingsToast === 'function') showSettingsToast(`${ym} を確定しました`);
        fireConfetti({ count: 50, origin: 0.3, spread: 0.6 });
        renderMonthlyClose();
    }

    function reopenMonth() {
        const ym = selectedYM();
        if (!confirm(`${ym} の確定を解除します。確定明細は削除されます。よろしいですか？`)) return;
        const all = monthlyCloseData();
        delete all[ym];
        lsSet(LS.monthlyClose, all);
        if (typeof saveSettingsToSpreadsheet === 'function') saveSettingsToSpreadsheet(false);
        if (typeof showSettingsToast === 'function') showSettingsToast(`${ym} の確定を解除しました`, 'warning');
        renderMonthlyClose();
    }

    function toggleCloseDetail() {
        $('close-detail')?.classList.toggle('hidden');
    }

    /* ================= アプリ内 日報入力 ================= */
    function buildDailyReportModal() {
        if ($('daily-report-modal')) return;
        const stores = Object.keys(STAFF_ROSTER);
        const today = new Date();
        const dateVal = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
        const numInput = (id, label, opts = {}) => `
            <label class="dr-field${opts.wide ? ' dr-wide' : ''}">
                <span>${label}</span>
                <input type="text" id="dr-${id}" inputmode="numeric" pattern="[0-9,]*" placeholder="0" value="" class="dr-input${opts.money ? ' money' : ''}">
            </label>`;

        const modal = document.createElement('div');
        modal.id = 'daily-report-modal';
        modal.className = 'dr-modal hidden';
        modal.innerHTML = `
            <div class="dr-backdrop" onclick="Enhance.closeDailyReport()"></div>
            <div class="dr-sheet" role="dialog" aria-label="日報入力">
                <div class="dr-handle"></div>
                <div class="flex items-center justify-between mb-4">
                    <h3 class="text-lg font-display font-bold text-accent-800 dark:text-surface-100 flex items-center gap-2">
                        <i data-lucide="notebook-pen" class="w-5 h-5 text-primary-500"></i>日報入力
                    </h3>
                    <button onclick="Enhance.closeDailyReport()" class="p-2 text-surface-500" aria-label="閉じる"><i data-lucide="x" class="w-5 h-5"></i></button>
                </div>
                <div class="dr-grid-3">
                    <label class="dr-field"><span>日付</span><input type="date" id="dr-date" value="${dateVal}" class="dr-input"></label>
                    <label class="dr-field"><span>店舗</span>
                        <select id="dr-store" class="dr-input" onchange="Enhance.onDrStoreChange()">
                            ${stores.map(s => `<option value="${s}">${storeName(s)}</option>`).join('')}
                        </select></label>
                    <label class="dr-field"><span>スタッフ</span><select id="dr-staff" class="dr-input"></select></label>
                </div>
                <p class="dr-section">💰 売上</p>
                <div class="dr-grid-4">
                    ${numInput('cash', '現金', { money: 1 })}${numInput('credit', 'クレジット', { money: 1 })}${numInput('qr', 'QR決済', { money: 1 })}${numInput('product', '物販', { money: 1 })}
                </div>
                <details class="dr-details">
                    <summary>値引き・返金（HPBポイント/ギフト等）</summary>
                    <div class="dr-grid-4">
                        ${numInput('hpbPoints', 'HPBポイント', { money: 1 })}${numInput('hpbGift', 'HPBギフト', { money: 1 })}${numInput('discOther', 'その他値引', { money: 1 })}${numInput('refund', '返金', { money: 1 })}
                    </div>
                </details>
                <p class="dr-section">👥 来店</p>
                <div class="dr-grid-4">
                    ${numInput('newHPB', '新規HPB')}${numInput('newMini', '新規minimo等')}${numInput('existing', '既存')}${numInput('acquaintance', '知り合い')}
                </div>
                <p class="dr-section">📅 次回予約</p>
                <div class="dr-grid-3">
                    ${numInput('resNewHPB', '新規HPB')}${numInput('resNewMini', '新規minimo等')}${numInput('resExisting', '既存')}
                </div>
                <p class="dr-section">⭐ その他</p>
                <div class="dr-grid-3">
                    ${numInput('reviews', '★5口コミ')}${numInput('blog', 'ブログ更新')}${numInput('sns', 'SNS更新')}
                </div>
                <div class="dr-summary" id="dr-summary">売上合計: ¥0</div>
                <button id="dr-submit" onclick="Enhance.submitDailyReport()" class="btn-primary w-full py-3.5 mt-3 text-base font-bold rounded-xl">
                    送信する
                </button>
            </div>`;
        document.body.appendChild(modal);

        // 数値整形 & 合計プレビュー
        modal.querySelectorAll('.dr-input[inputmode="numeric"]').forEach(inp => {
            inp.addEventListener('input', updateDrSummary);
            inp.addEventListener('blur', () => {
                const v = parseInt(inp.value.replace(/[^0-9]/g, '')) || 0;
                inp.value = v ? v.toLocaleString() : '';
                updateDrSummary();
            });
        });
        onDrStoreChange();
        if (window.lucide) lucide.createIcons();
    }

    const drVal = id => parseInt(($(`dr-${id}`)?.value || '').replace(/[^0-9]/g, '')) || 0;

    function updateDrSummary() {
        const total = drVal('cash') + drVal('credit') + drVal('qr') + drVal('hpbPoints') + drVal('hpbGift');
        const cust = drVal('newHPB') + drVal('newMini') + drVal('existing') + drVal('acquaintance');
        const el = $('dr-summary');
        if (el) el.innerHTML = `売上合計: <b>${yen(total)}</b>　来店: <b>${cust}名</b>${cust > 0 ? `　客単価: ${yen(total / cust)}` : ''}`;
    }

    function onDrStoreChange() {
        const storeSel = $('dr-store');
        const staffSel = $('dr-staff');
        if (!storeSel || !staffSel) return;
        const names = STAFF_ROSTER[storeSel.value] || [];
        staffSel.innerHTML = names.map(n => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
    }

    function openDailyReport() {
        buildDailyReportModal();
        const modal = $('daily-report-modal');
        // 店舗リストを最新の名簿で更新（設定読み込み後に変わる可能性があるため）
        const storeSel = $('dr-store');
        if (storeSel) {
            const prev = storeSel.value;
            storeSel.innerHTML = Object.keys(STAFF_ROSTER).map(s => `<option value="${s}">${storeName(s)}</option>`).join('');
            if ([...storeSel.options].some(o => o.value === prev)) storeSel.value = prev;
            onDrStoreChange();
        }
        // スタッフ専用ページならロック
        if (lockedStore) {
            $('dr-store').value = lockedStore;
            $('dr-store').disabled = true;
            onDrStoreChange();
        }
        if (lockedStaff) {
            $('dr-staff').value = lockedStaff;
            $('dr-staff').disabled = true;
        }
        updateDrSummary();
        modal.classList.remove('hidden');
        requestAnimationFrame(() => modal.classList.add('open'));
        document.body.style.overflow = 'hidden';
    }

    function closeDailyReport() {
        const modal = $('daily-report-modal');
        if (!modal) return;
        modal.classList.remove('open');
        setTimeout(() => modal.classList.add('hidden'), 250);
        document.body.style.overflow = '';
    }

    async function submitDailyReport() {
        const btn = $('dr-submit');
        const dateStr = $('dr-date')?.value;
        const store = $('dr-store')?.value;
        const staff = $('dr-staff')?.value;
        if (!dateStr || !store || !staff) { alert('日付・店舗・スタッフを入力してください'); return; }
        const [y, m, d] = dateStr.split('-').map(Number);
        const dateKey = `${y}/${m}/${d}`;

        // 重複チェック
        const dup = (rawData || []).some(r => r && r.date === dateKey && r.store === store && r.staff?.toLowerCase() === staff.toLowerCase());
        if (dup && !confirm(`${dateKey} の ${staff} さんの日報は既に存在します。\n追加でもう1件登録しますか？（合算されます）`)) return;

        const record = {
            date: dateKey, store, staff,
            sales: { cash: drVal('cash'), credit: drVal('credit'), qr: drVal('qr'), product: drVal('product') },
            discounts: { hpbPoints: drVal('hpbPoints'), hpbGift: drVal('hpbGift'), other: drVal('discOther'), refund: drVal('refund') },
            customers: { newHPB: drVal('newHPB'), newMiniNai: drVal('newMini'), existing: drVal('existing'), acquaintance: drVal('acquaintance') },
            nextRes: { newHPB: drVal('resNewHPB'), newMiniNai: drVal('resNewMini'), existing: drVal('resExisting') },
            reviews5Star: drVal('reviews'), blogUpdates: drVal('blog'), snsUpdates: drVal('sns'),
        };

        btn.disabled = true; btn.textContent = '送信中…';
        try {
            const apiUrl = localStorage.getItem('mavie_spreadsheet_api_url')
                || (typeof API_URL !== 'undefined' && API_URL)
                || (typeof DEFAULT_API_URL !== 'undefined' && DEFAULT_API_URL)
                || (window.Backend && Backend.mode() === 'supabase' ? 'supabase://rpc' : '');
            if (!apiUrl) throw new Error('バックエンドが未設定です（設定タブで登録してください）');
            const send = window.apiFetch || fetch;
            const res = await send(apiUrl, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain' },
                body: JSON.stringify({ action: 'add_record', record }),
            });
            const result = await res.json();
            if (result.status !== 'success') throw new Error(result.message || '保存に失敗しました');
            closeDailyReport();
            if (typeof showSettingsToast === 'function') showSettingsToast('日報を送信しました！おつかれさまでした 🎉');
            fireConfetti({ count: 50, origin: 0.3, spread: 0.6 });
            // 入力値をリセットして最新データを取得
            document.querySelectorAll('#daily-report-modal .dr-input[inputmode="numeric"]').forEach(i => i.value = '');
            await silentRefresh('report');
        } catch (e) {
            alert(`送信エラー: ${e.message}`);
        } finally {
            btn.disabled = false; btn.textContent = '送信する';
        }
    }

    /* ================= PWA 登録 ================= */
    function registerPwa() {
        if (!('serviceWorker' in navigator)) return;
        if (location.protocol !== 'https:' && location.hostname !== 'localhost' && location.hostname !== '127.0.0.1') return;
        navigator.serviceWorker.register('./sw.js').catch(e => console.warn('SW登録スキップ:', e.message));
    }

    /* ================= メインフック ================= */
    function onDashboardUpdated({ filtered, metrics, currentGoal }) {
        if (lastFetchAt === null) { lastFetchAt = Date.now(); setLiveIndicator('ok'); }
        renderGreeting(currentGoal);
        renderGoalRings(metrics, currentGoal);
        renderRemaining(currentGoal);
        renderForecastV2(currentGoal);
        renderYoY(metrics);
        updateRankHistory();
        renderHighlights(currentGoal);
        renderStoreRace();
        renderWeekdayChart(metrics);
        renderStaffRadar();
        renderStaffStreaks();
        renderFollowList();
        renderRoi();
        renderMonthlyClose();
        maybeCelebrate(currentGoal);
        maybeAutoAI();
        if (window.lucide) lucide.createIcons();
    }

    /* ================= 初期化 ================= */
    document.addEventListener('DOMContentLoaded', () => {
        // マーケタブの更新にフォローリスト/ROIを連動
        if (typeof window.updateMarketingDashboard === 'function') {
            const orig = window.updateMarketingDashboard;
            window.updateMarketingDashboard = function (...args) {
                const r = orig.apply(this, args);
                try { renderFollowList(); renderRoi(); } catch (e) {}
                return r;
            };
        }
        setupAdminAiCache();
        startLive();
        registerPwa();
        buildDailyReportModal();
    });

    window.Enhance = {
        onDashboardUpdated,
        rankMoveBadge,
        openDailyReport,
        closeDailyReport,
        submitDailyReport,
        onDrStoreChange,
        saveRoi,
        closeMonth,
        reopenMonth,
        toggleCloseDetail,
        silentRefresh,
        fireConfetti,
    };
})();
