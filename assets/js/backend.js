/* =====================================================================
 * vie Dashboard Backend Adapter
 *
 * GAS(スプレッドシート) と Supabase(Postgres) を設定で切り替える抽象化層。
 *
 *   window.apiFetch(url, options)
 *     - GASモード      : fetch() への素通し（挙動は完全に従来どおり）
 *     - Supabaseモード : URL/bodyから action を解釈し、Postgres RPC
 *                        (api_* 関数) を呼んで GAS と同じ形のレスポンスを
 *                        Response オブジェクトとして返す
 *
 * dashboard.js / enhancements.js は apiFetch を呼ぶだけで、
 * どちらのバックエンドかを意識しない。移行は設定タブで切替・即ロールバック可。
 * スキーマとRPC定義: supabase/migrations/0001_init.sql
 * ===================================================================== */
(function () {
    'use strict';

    const MODE_KEY = 'mavie_backend_mode';          // 'gas' | 'supabase'
    const SB_URL_KEY = 'mavie_supabase_url';
    const SB_ANON_KEY = 'mavie_supabase_anon_key';
    const SB_SESSION_KEY = 'mavie_sb_sessions';     // {token: {pageType, expiresAt}}

    function mode() {
        const m = localStorage.getItem(MODE_KEY);
        return (m === 'supabase' && sbConfigured()) ? 'supabase' : 'gas';
    }
    function sbUrl() { return (localStorage.getItem(SB_URL_KEY) || '').replace(/\/+$/, ''); }
    function sbKey() { return localStorage.getItem(SB_ANON_KEY) || ''; }
    function sbConfigured() { return !!(localStorage.getItem(SB_URL_KEY) && localStorage.getItem(SB_ANON_KEY)); }

    /* ---------------- Supabase REST/RPC ---------------- */
    async function rpc(fn, args = {}) {
        const res = await fetch(`${sbUrl()}/rest/v1/rpc/${fn}`, {
            method: 'POST',
            headers: {
                'apikey': sbKey(),
                'Authorization': `Bearer ${sbKey()}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(args),
        });
        const text = await res.text();
        if (!res.ok) {
            let msg = text;
            try { msg = JSON.parse(text).message || text; } catch (e) {}
            throw new Error(`Supabase ${fn}: ${msg} (HTTP ${res.status})`);
        }
        return text ? JSON.parse(text) : null;
    }

    /* ---------------- セッション（クライアント保持・GAS同等レベル） ---------------- */
    function sessions() {
        try { return JSON.parse(localStorage.getItem(SB_SESSION_KEY) || '{}'); } catch (e) { return {}; }
    }
    function createSession(pageType) {
        const token = (crypto.randomUUID && crypto.randomUUID()) || String(Math.random()).slice(2) + Date.now();
        const expiresAt = Date.now() + 24 * 3600 * 1000;
        const all = sessions();
        all[token] = { pageType, expiresAt };
        // 期限切れを掃除
        Object.keys(all).forEach(t => { if (all[t].expiresAt < Date.now()) delete all[t]; });
        localStorage.setItem(SB_SESSION_KEY, JSON.stringify(all));
        return { token, expiresAt };
    }
    function verifySession(token, pageType) {
        const s = sessions()[token];
        return !!(s && s.pageType === pageType && s.expiresAt > Date.now());
    }

    /* ---------------- action → RPC ルーティング ---------------- */
    async function handleAction(action, params, body) {
        switch (action) {
            case 'get_data':
            case 'health': {
                if (action === 'health') return rpc('api_health');
                return rpc('api_get_sales'); // GASと同じく素の配列を返す
            }
            case 'get_customers':
                return { status: 'success', data: await rpc('api_get_customers') };
            case 'get_customers_today':
                return { status: 'success', data: await rpc('api_get_customers', { p_today: true }) };
            case 'get_customers_by_store':
                return { status: 'success', data: await rpc('api_get_customers', { p_store: params.store || null }) };
            case 'load_goals': {
                const r = await rpc('api_load_goals');
                return { status: 'success', goals: r.goals || {}, salaries: r.salaries || {} };
            }
            case 'save_goals':
                await rpc('api_save_goals', { p_goals: body.goals || {}, p_salaries: body.salaries || {} });
                return { status: 'success', message: '目標を保存しました' };
            case 'load_settings': {
                const r = await rpc('api_load_settings');
                return { status: 'success', settings: r };
            }
            case 'save_settings':
                await rpc('api_save_settings', { p: body.settings || {} });
                return { status: 'success', message: '設定を保存しました' };
            case 'load_passwords': {
                const r = await rpc('api_load_passwords');
                return { status: 'success', passwords: r || {} };
            }
            case 'save_passwords':
                await rpc('api_save_passwords', { p: body.passwords || {} });
                return { status: 'success', message: 'パスワードを保存しました' };
            case 'add_record':
                await rpc('api_add_record', { p: body.record });
                return { status: 'success', message: 'レコードを追加しました' };
            case 'update':
                await rpc('api_update_rows', { p: body.rows || [] });
                return { status: 'success', message: `${(body.rows || []).length}件を更新しました` };
            case 'verify_password': {
                const r = await rpc('api_verify_password', {
                    p_page_type: params.page_type || 'admin',
                    p_store: params.store || '',
                    p_staff: params.staff || '',
                    p_password: params.password || '',
                });
                if (r && r.ok) {
                    const s = createSession(params.page_type || 'admin');
                    return { status: 'success', sessionToken: s.token, expiresAt: s.expiresAt };
                }
                return { status: 'error', message: 'パスワードが正しくありません' };
            }
            case 'verify_session':
                return verifySession(params.session_token, params.page_type)
                    ? { status: 'success' }
                    : { status: 'error', message: 'セッションが無効です' };
            case 'clear_cache':
                return { status: 'success', message: 'Supabaseモードではキャッシュ操作は不要です' };
            default:
                return { status: 'error', message: `未対応のアクション: ${action}` };
        }
    }

    /* ---------------- apiFetch ゲートウェイ ---------------- */
    async function apiFetch(url, options = {}) {
        if (mode() !== 'supabase') {
            return fetch(url, options); // GASモード: 完全素通し
        }
        // action の解釈（GET: クエリ / POST: body JSON）
        let action = 'get_data';
        const params = {};
        let body = null;
        try {
            const u = new URL(String(url), location.href);
            u.searchParams.forEach((v, k) => { params[k] = v; });
            if (params.action) action = params.action;
        } catch (e) { /* 相対/不正URLでも続行 */ }
        if (options && options.method === 'POST' && options.body) {
            try {
                body = JSON.parse(options.body);
                action = body.action || 'update';
            } catch (e) {
                return new Response(JSON.stringify({ status: 'error', message: 'リクエストボディの解析に失敗しました' }), { status: 200, headers: { 'Content-Type': 'application/json' } });
            }
        }
        try {
            const result = await handleAction(action, params, body);
            return new Response(JSON.stringify(result), { status: 200, headers: { 'Content-Type': 'application/json' } });
        } catch (e) {
            console.error('Supabaseバックエンドエラー:', e);
            return new Response(JSON.stringify({ status: 'error', message: e.message }), { status: 200, headers: { 'Content-Type': 'application/json' } });
        }
    }

    /* ---------------- スタッフパスワード照合（Supabaseはハッシュのため都度RPC） ---------------- */
    async function verifyStaffPassword(store, staff, password) {
        const r = await rpc('api_verify_password', { p_page_type: 'staff', p_store: store, p_staff: staff, p_password: password || '' });
        return !!(r && r.ok);
    }

    /* ---------------- Gemini 中継 (Edge Function / キー秘匿) ---------------- */
    async function aiProxy(prompt) {
        const res = await fetch(`${sbUrl()}/functions/v1/gemini-advice`, {
            method: 'POST',
            headers: {
                'apikey': sbKey(),
                'Authorization': `Bearer ${sbKey()}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ prompt }),
        });
        const data = await res.json();
        if (!res.ok || data.error) throw new Error(data.error || `Edge Function エラー (HTTP ${res.status})`);
        return data.text;
    }

    /* ---------------- Realtime（日報の即時反映・任意） ---------------- */
    let realtimeChannel = null;
    async function initRealtime(onChange) {
        if (mode() !== 'supabase' || realtimeChannel) return false;
        try {
            if (!window.supabase) {
                await new Promise((resolve, reject) => {
                    const s = document.createElement('script');
                    s.src = 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/dist/umd/supabase.min.js';
                    s.onload = resolve;
                    s.onerror = () => reject(new Error('supabase-js の読み込みに失敗'));
                    document.head.appendChild(s);
                });
            }
            const client = window.supabase.createClient(sbUrl(), sbKey());
            realtimeChannel = client
                .channel('daily_reports_live')
                .on('postgres_changes', { event: '*', schema: 'public', table: 'daily_reports' }, () => {
                    if (typeof onChange === 'function') onChange();
                })
                .subscribe();
            console.log('✓ Supabase Realtime: 日報の即時反映が有効になりました');
            return true;
        } catch (e) {
            console.warn('Realtime初期化スキップ:', e.message);
            return false;
        }
    }

    /* ---------------- 設定タブ UI ---------------- */
    function initSettingsUI() {
        const modeRadios = document.querySelectorAll('input[name="backend-mode"]');
        if (!modeRadios.length) return;
        const saved = localStorage.getItem(MODE_KEY) || 'gas';
        modeRadios.forEach(r => { r.checked = (r.value === saved); });
        const urlInput = document.getElementById('supabase-url');
        const keyInput = document.getElementById('supabase-anon-key');
        if (urlInput) urlInput.value = localStorage.getItem(SB_URL_KEY) || '';
        if (keyInput) keyInput.value = localStorage.getItem(SB_ANON_KEY) || '';
        updateModeBadge();
    }

    function updateModeBadge() {
        const badge = document.getElementById('backend-mode-badge');
        if (!badge) return;
        if (mode() === 'supabase') {
            badge.textContent = 'Supabase 接続中';
            badge.className = 'text-xs font-bold px-3 py-1 rounded-full bg-emerald-100 text-emerald-700';
        } else {
            badge.textContent = 'GAS (スプレッドシート) 接続中';
            badge.className = 'text-xs font-bold px-3 py-1 rounded-full bg-surface-200 text-surface-700';
        }
    }

    function saveBackendConfig() {
        const selected = document.querySelector('input[name="backend-mode"]:checked')?.value || 'gas';
        const url = document.getElementById('supabase-url')?.value.trim() || '';
        const key = document.getElementById('supabase-anon-key')?.value.trim() || '';
        if (url) localStorage.setItem(SB_URL_KEY, url); else localStorage.removeItem(SB_URL_KEY);
        if (key) localStorage.setItem(SB_ANON_KEY, key); else localStorage.removeItem(SB_ANON_KEY);
        if (selected === 'supabase' && !(url && key)) {
            if (typeof showSettingsToast === 'function') showSettingsToast('Supabase URL と anon key を入力してください', 'error');
            return;
        }
        localStorage.setItem(MODE_KEY, selected);
        updateModeBadge();
        if (typeof showSettingsToast === 'function') showSettingsToast(`バックエンドを ${selected === 'supabase' ? 'Supabase' : 'GAS'} に切り替えました。再読み込みします…`);
        setTimeout(() => location.reload(), 900);
    }

    async function testSupabaseConnection() {
        const status = document.getElementById('supabase-connection-status');
        const url = document.getElementById('supabase-url')?.value.trim();
        const key = document.getElementById('supabase-anon-key')?.value.trim();
        if (!url || !key) {
            if (status) { status.classList.remove('hidden'); status.innerHTML = '<div class="bg-yellow-50 border border-yellow-400 text-yellow-700 px-4 py-2 rounded text-sm">URL と anon key を入力してください</div>'; }
            return;
        }
        if (status) { status.classList.remove('hidden'); status.innerHTML = '<div class="bg-blue-50 border border-blue-400 text-blue-700 px-4 py-2 rounded text-sm">接続テスト中…</div>'; }
        try {
            const res = await fetch(`${url.replace(/\/+$/, '')}/rest/v1/rpc/api_health`, {
                method: 'POST',
                headers: { 'apikey': key, 'Authorization': `Bearer ${key}`, 'Content-Type': 'application/json' },
                body: '{}',
            });
            const data = await res.json();
            if (!res.ok) throw new Error(data.message || `HTTP ${res.status}（migrations未適用の可能性）`);
            if (status) status.innerHTML = `<div class="bg-green-50 border border-green-400 text-green-700 px-4 py-2 rounded text-sm">✓ 接続成功！ 日報 ${data.reports ?? '?'}件 / 顧客 ${data.customers ?? '?'}件 / スキーマ ${data.version ?? '?'}</div>`;
        } catch (e) {
            if (status) status.innerHTML = `<div class="bg-red-50 border border-red-400 text-red-700 px-4 py-2 rounded text-sm">✗ 接続エラー: ${e.message}</div>`;
        }
    }

    document.addEventListener('DOMContentLoaded', initSettingsUI);

    window.apiFetch = apiFetch;
    window.Backend = {
        mode,
        sbConfigured,
        rpc,
        verifyStaffPassword,
        aiProxy,
        initRealtime,
        saveBackendConfig,
        testSupabaseConnection,
        updateModeBadge,
    };
})();
