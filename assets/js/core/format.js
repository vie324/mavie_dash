// 表示フォーマットと日付ヘルパー（日付はすべて日本時間の暦日で扱う）

export function yen(n) {
    if (n === null || n === undefined || isNaN(n)) return '—';
    return '¥' + Math.round(n).toLocaleString('ja-JP');
}

export function yenShort(n) {
    if (n === null || n === undefined || isNaN(n)) return '—';
    const abs = Math.abs(n);
    if (abs >= 100000000) return '¥' + (n / 100000000).toFixed(1) + '億';
    if (abs >= 10000) return '¥' + Math.round(n / 10000).toLocaleString('ja-JP') + '万';
    return yen(n);
}

export function num(n) {
    if (n === null || n === undefined || isNaN(n)) return '—';
    return Math.round(n).toLocaleString('ja-JP');
}

export function pct(n, digits = 1) {
    if (n === null || n === undefined || isNaN(n)) return '—';
    return Number(n).toFixed(digits) + '%';
}

export function esc(s) {
    return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

const pad2 = n => String(n).padStart(2, '0');

// 日本時間の「今日」を {y, m, d} で返す
export function todayJst() {
    const t = new Date(Date.now() + 9 * 3600 * 1000);
    return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
}

export function ymd(y, m, d) {
    return `${y}-${pad2(m)}-${pad2(d)}`;
}

export function todayStr() {
    const t = todayJst();
    return ymd(t.y, t.m, t.d);
}

export function daysInMonth(y, m) {
    return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

// 期間フィルタ → {from, to}（fromは開始月の1日、toは基準月の末日）
export function periodRange(kind, anchor) {
    const months = { month: 1, '3months': 3, '6months': 6, year: 12 }[kind] || 1;
    let sy = anchor.y, sm = anchor.m - (months - 1);
    while (sm < 1) { sm += 12; sy -= 1; }
    return {
        from: ymd(sy, sm, 1),
        to: ymd(anchor.y, anchor.m, daysInMonth(anchor.y, anchor.m)),
        months,
    };
}

// 直前の同じ長さの期間
export function prevRange(range) {
    const months = range.months;
    const [fy, fm] = range.from.split('-').map(Number);
    let sy = fy, sm = fm - months;
    while (sm < 1) { sm += 12; sy -= 1; }
    let ey = fy, em = fm - 1;
    while (em < 1) { em += 12; ey -= 1; }
    return { from: ymd(sy, sm, 1), to: ymd(ey, em, daysInMonth(ey, em)), months };
}

// 前年同期間
export function yoyRange(range) {
    const [fy, fm] = range.from.split('-').map(Number);
    const [ty, tm] = range.to.split('-').map(Number);
    return { from: ymd(fy - 1, fm, 1), to: ymd(ty - 1, tm, daysInMonth(ty - 1, tm)), months: range.months };
}

export function monthLabel(anchor) {
    return `${anchor.y}年${anchor.m}月`;
}

// 'YYYY-MM-DD' → 表示用 'M/D' or 'M月'
export function shortDate(dateStr) {
    const [, m, d] = dateStr.split('-').map(Number);
    return `${m}/${d}`;
}

// 曜日（閲覧環境のタイムゾーンに依存しないようUTC正午で判定）
export function dowIndex(dateStr) {
    return new Date(dateStr + 'T12:00:00Z').getUTCDay();
}

export function dowJa(dateStr) {
    return ['日', '月', '火', '水', '木', '金', '土'][dowIndex(dateStr)];
}

// 差分バッジ用: {text, dir: 'up'|'down'|'flat'}
export function delta(cur, prev) {
    if (!prev || prev === 0 || cur === null || cur === undefined) return { text: '—', dir: 'flat' };
    const diff = (cur - prev) / prev * 100;
    if (Math.abs(diff) < 0.05) return { text: '±0%', dir: 'flat' };
    return {
        text: `${diff > 0 ? '+' : ''}${diff.toFixed(1)}%`,
        dir: diff > 0 ? 'up' : 'down',
    };
}

export function applyDeltaBadge(el, d) {
    if (!el) return;
    el.textContent = d.text;
    el.classList.remove('up', 'down');
    if (d.dir === 'up') el.classList.add('up');
    if (d.dir === 'down') el.classList.add('down');
}

// 数値カウントアップ演出
export function countUp(el, value, { prefix = '', suffix = '', duration = 600 } = {}) {
    if (!el) return;
    const start = Number(el.dataset.value || 0);
    const end = Math.round(value || 0);
    el.dataset.value = end;
    if (matchMedia('(prefers-reduced-motion: reduce)').matches || Math.abs(end - start) < 2) {
        el.textContent = prefix + end.toLocaleString('ja-JP') + suffix;
        return;
    }
    const t0 = performance.now();
    (function frame(now) {
        const p = Math.min((now - t0) / duration, 1);
        const eased = 1 - Math.pow(1 - p, 3);
        const v = Math.round(start + (end - start) * eased);
        el.textContent = prefix + v.toLocaleString('ja-JP') + suffix;
        if (p < 1) requestAnimationFrame(frame);
    })(t0);
}
