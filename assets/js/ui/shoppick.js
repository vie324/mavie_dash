// 「店舗を選択してください」ガード内の店舗クイック選択ボタン
// ヘッダーの店舗セレクタと同期させる（選ぶとフィルタ変更として扱われる）

import { state, isLocked } from '../core/state.js';
import { esc } from '../core/format.js';

export function renderShopPick(containerId) {
    const el = document.getElementById(containerId);
    if (!el) return;
    if (isLocked() || state.masters.shops.length === 0) { el.innerHTML = ''; return; }
    el.innerHTML = state.masters.shops.map(s =>
        `<button type="button" class="shop-pick" data-pick-shop="${s.id}"><i data-lucide="store" class="w-4 h-4"></i>${esc(s.name)}</button>`
    ).join('');
    if (window.lucide) lucide.createIcons({ nodes: [...el.querySelectorAll('[data-lucide]')] });
    el.querySelectorAll('[data-pick-shop]').forEach(btn => btn.addEventListener('click', () => {
        const sel = document.getElementById('store-selector');
        if (!sel) return;
        sel.value = btn.dataset.pickShop;
        sel.dispatchEvent(new Event('change'));
    }));
}
