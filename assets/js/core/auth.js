// 認証フロー
// URLパラメータ（?store=&staff= 従来互換 / ?shop_id=&staff_id= 新形式）で閲覧コンテキストを決め、
// サーバーにセッションを問い合わせる。パスワードが必要ならログインモーダルを表示。

import { authGet, authLogin } from './api.js';
import { state } from './state.js';

export function urlContext() {
    const p = new URLSearchParams(location.search);
    const store = p.get('store') || p.get('shop_id') || '';
    const staff = p.get('staff') || p.get('staff_id') || '';
    return { store, staff };
}

function show(el) { el?.classList.remove('hidden'); }
function hide(el) { el?.classList.add('hidden'); }

// 認証完了までresolveしないPromiseを返す
export async function ensureAuthenticated() {
    const ctx = urlContext();
    let res;
    try {
        res = await authGet(ctx);
    } catch (e) {
        // サーバー未達: 認証不能。エラー画面を出して停止
        showFatal('サーバーに接続できません', 'ネットワーク接続を確認して再読み込みしてください。');
        throw e;
    }

    state.demo = !!res.demo;

    if (res.context?.role === 'invalid') {
        showFatal('URLが正しくありません', '店舗またはスタッフの指定が見つかりませんでした。管理者に正しいURLを確認してください。');
        throw new Error('invalid context');
    }

    if (res.authenticated) {
        state.session = res.session;
        return res.session;
    }

    // パスワード入力が必要
    return new Promise(resolve => {
        const modal = document.getElementById('login-modal');
        const nameEl = document.getElementById('login-target-name');
        const form = document.getElementById('login-form');
        const input = document.getElementById('login-password');
        const errEl = document.getElementById('login-error');
        const isStaff = res.context?.role === 'staff';
        if (nameEl) {
            nameEl.textContent = isStaff
                ? `${res.context.shopName || ''} ${res.context.staffName || ''}`.trim()
                : '管理者ログイン';
        }
        show(modal);
        input?.focus();

        form.onsubmit = async (ev) => {
            ev.preventDefault();
            hide(errEl);
            const btn = form.querySelector('button[type=submit]');
            btn.disabled = true;
            try {
                const out = await authLogin({ ...ctx, password: input.value });
                state.session = out.session;
                hide(modal);
                resolve(out.session);
            } catch (e) {
                show(errEl);
                errEl.textContent = e.status === 401 ? 'パスワードが正しくありません' : 'ログインに失敗しました。時間をおいて再度お試しください';
                input.select();
            } finally {
                btn.disabled = false;
            }
        };
    });
}

function showFatal(title, message) {
    const overlay = document.getElementById('loading-overlay');
    if (overlay) {
        overlay.innerHTML = `
            <div class="text-center px-6">
                <div class="w-16 h-16 mx-auto mb-4 rounded-full bg-red-100 flex items-center justify-center text-3xl">⚠️</div>
                <p class="text-lg font-display font-semibold text-accent-800 mb-2"></p>
                <p class="text-sm text-surface-600"></p>
            </div>`;
        overlay.querySelector('p.text-lg').textContent = title;
        overlay.querySelector('p.text-sm').textContent = message;
    }
}
