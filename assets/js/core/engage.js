// エンゲージメント演出（旧enhancements.jsから移植）
// 目標リング・紙吹雪・達成バナー・LIVEインジケータ・挨拶

import { yen } from './format.js';

function reducedMotion() {
    return matchMedia('(prefers-reduced-motion: reduce)').matches;
}

// ---- 目標プログレスリング ----
export function ringSvg(pctValue, color, over) {
    const r = 30, c = 2 * Math.PI * r;
    const clamped = Math.max(0, Math.min(pctValue, 100));
    const off = c * (1 - clamped / 100);
    return `<svg viewBox="0 0 72 72" class="goal-ring-svg${over ? ' over' : ''}">
        <circle cx="36" cy="36" r="${r}" class="goal-ring-track"/>
        <circle cx="36" cy="36" r="${r}" class="goal-ring-fill" stroke="${color}"
            stroke-dasharray="${c.toFixed(1)}" stroke-dashoffset="${off.toFixed(1)}" pathLength="${c.toFixed(1)}"/>
    </svg>`;
}

// rings: [{label, pct, color, value, sub}]
export function renderRings(containerId, rings) {
    const wrap = document.getElementById(containerId);
    if (!wrap) return;
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

// ---- 紙吹雪 ----
let confettiCanvas = null;

export function fireConfetti({ count = 90, origin = 0.4, spread = 1 } = {}) {
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
            ctx.translate(p.x, p.y);
            ctx.rotate(p.rot);
            ctx.fillStyle = p.color;
            ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
            ctx.restore();
        });
        if (dt < 1) requestAnimationFrame(frame);
        else ctx.clearRect(0, 0, cv.width, cv.height);
    })(t0);
}

export function showCelebrationBanner(title, sub) {
    const old = document.getElementById('celebration-banner');
    if (old) old.remove();
    const div = document.createElement('div');
    div.id = 'celebration-banner';
    div.className = 'celebration-banner';
    div.innerHTML = `<div class="celebration-inner">
        <span class="text-2xl">🏆</span>
        <div><p class="font-bold"></p><p class="text-xs opacity-90"></p></div>
        <button aria-label="閉じる" class="ml-2 opacity-70 hover:opacity-100">✕</button>
    </div>`;
    div.querySelector('p.font-bold').textContent = title;
    div.querySelector('p.text-xs').textContent = sub;
    div.querySelector('button').addEventListener('click', () => div.remove());
    document.body.appendChild(div);
    setTimeout(() => div.classList.add('show'), 30);
    setTimeout(() => { div.classList.remove('show'); setTimeout(() => div.remove(), 500); }, 8000);
}

// 月間目標達成のセレブレーション（スコープ×月ごとに1回だけ）
const CELEBRATE_KEY = 'vie_celebrated_v3';

export function maybeCelebrate(scopeKey, monthKey, salesTotal, goal) {
    if (!goal || goal <= 0 || salesTotal < goal) return;
    try {
        const flags = JSON.parse(localStorage.getItem(CELEBRATE_KEY) || '{}');
        const key = `${monthKey}:${scopeKey}`;
        if (flags[key]) return;
        flags[key] = true;
        localStorage.setItem(CELEBRATE_KEY, JSON.stringify(flags));
        fireConfetti({ count: 160, origin: 0.35, spread: 1.4 });
        showCelebrationBanner('🎉 月間目標を達成しました！', `${yen(salesTotal)} / ${yen(goal)}`);
    } catch (_) { /* localStorage不可でも動作継続 */ }
}

// ---- LIVEインジケータ ----
export function setLiveIndicator(stateName, time) {
    const el = document.getElementById('live-indicator');
    if (!el) return;
    const timeStr = time ? new Date(time).toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' }) : '--:--';
    el.classList.remove('hidden');
    el.classList.toggle('syncing', stateName === 'syncing');
    el.classList.toggle('error', stateName === 'error');
    const label = el.querySelector('.live-time');
    if (label) label.textContent = stateName === 'syncing' ? '更新中…' : stateName === 'error' ? '接続エラー' : `${timeStr} 更新`;
}

// ---- 挨拶メッセージ ----
export function greeting(name, progressRatio) {
    const h = new Date().getHours();
    const timeGreet = h < 5 ? 'お疲れさまです' : h < 11 ? 'おはようございます' : h < 18 ? 'こんにちは' : 'お疲れさまです';
    const who = name ? `、${name}さん` : '';
    let cheer = '';
    if (progressRatio !== null && progressRatio !== undefined && isFinite(progressRatio)) {
        if (progressRatio >= 1) cheer = ' 目標達成おめでとうございます！🎉';
        else if (progressRatio >= 0.85) cheer = ' 目標まであと少しです！';
        else if (progressRatio >= 0.5) cheer = ' 好調です、この調子！';
        else cheer = ' 追い上げチャンスです💪';
    }
    return `${timeGreet}${who}。${cheer}`;
}

// ---- トースト ----
export function toast(message, type = 'info') {
    const old = document.getElementById('app-toast');
    if (old) old.remove();
    const div = document.createElement('div');
    div.id = 'app-toast';
    div.className = `app-toast ${type}`;
    div.setAttribute('role', 'status');
    div.textContent = message;
    document.body.appendChild(div);
    setTimeout(() => div.classList.add('show'), 20);
    setTimeout(() => { div.classList.remove('show'); setTimeout(() => div.remove(), 400); }, 4000);
}
