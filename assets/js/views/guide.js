// 使い方・業務分担: SalonOneでやること / このツールでやること / 役割別ルーティン / 見る数字の定義
// 内容は docs/OPERATIONS.md と同じ（現場でスマホからすぐ確認できるようアプリ内にも置く）。

import { state, on } from '../core/state.js';
import { switchTab } from '../ui/nav.js';

// SalonOneでやること（入力の正）と、それがこのツールのどこに効くか
const SALONONE_TASKS = [
    { what: '予約の登録・変更（担当スタッフを必ず設定）', why: 'スタッフ別の新規数・入会率、次回予約の自動推定' },
    { what: '会計（支払い方法を正しく選ぶ・物販は物販として登録）', why: '売上・客単価・入金突合・物販インセンティブ' },
    { what: '次回予約は会計時にSalonOneで予約登録', why: '次回予約率の自動推定（β）。日報の数字と突き合わせ' },
    { what: 'キャンセル・無断キャンセルの処理（当日中）', why: 'キャンセル率。処理漏れはホームに「会計未処理」として出ます' },
    { what: '初回来店の顧客に流入元・生年を登録', why: '媒体別の集客効果（CPA・入会率）・年代分析' },
    { what: '回数券・サブスクの契約登録', why: '入会数・継続率' },
    { what: 'スタッフの入退社・メニュー・店舗の登録', why: 'スタッフ一覧・各集計（このツール側での登録は不要）' },
    { what: '予約台帳・カルテ・顧客対応', why: 'このツールでは扱いません（個人情報は表示しない設計）' },
];

const TOOL_TASKS = [
    { head: '見る', items: ['売上・来店・客単価と目標の進捗、着地予測', 'スタッフ別・店舗別の比較', '新規 → 入会 → 継続の流れ、媒体別の集客効果'] },
    { head: '入れる（SalonOneにない数字だけ）', items: ['日報: 次回予約（新規/既存）・ブログ・SNS・★5口コミ', '入金突合: 現金の実査額・カード/QRの端末集計', '目標・基本給、HPBなどAPIにない広告費'] },
    { head: '決める・回す', items: ['シフト: 希望休の申請 → 自動分配 → 承認', 'インセンティブの確認（月初）', 'ホームの「やること」で抜け漏れをゼロに'] },
];

const ROUTINES = {
    staff: {
        label: 'スタッフ',
        daily: ['退勤前に日報を入力（次回予約の新規/既存・ブログ・SNS・★5口コミ）', 'ホームで今日の実績と今月の進捗を確認'],
        weekly: ['マイ成績で次回予約率・客単価を振り返る', 'SalonOneで次回予約を会計時に入れているか確認'],
        monthly: ['翌月の希望休を締切日までに申請', '月初に今月の目標（店長が設定）を確認'],
        numbers: ['今月の売上と目標までの残り', '次回予約率（新規・既存）', '新規の入会（契約）数', '店内順位・客単価'],
    },
    store: {
        label: '店長',
        daily: ['閉店後に入金突合（現金を数えて入力・カード/QRは端末の日計）', 'ホームで日報の未入力者を確認して声かけ', 'SalonOneの会計未処理・担当者未設定をなくす'],
        weekly: ['スタッフ別の次回予約率・客単価・キャンセル率を確認', '着地予測と目標の差を見て打ち手を決める'],
        monthly: ['月末までに翌月の目標を設定', '申請締切後にシフトを自動分配 → 調整 → 承認', '月末に今月の未突合をゼロにする'],
        numbers: ['店舗の目標進捗・着地予測', '日報の入力率', '次回予約率（新規・既存）', '新規の入会率・キャンセル率'],
    },
    manager: {
        label: 'マネージャー',
        daily: ['ホームで全店舗の未突合・日報未入力を確認'],
        weekly: ['店舗比較（目標達成率・客単価・新規・キャンセル率）', '媒体別の新規獲得と入会率（マーケ）'],
        monthly: ['店舗目標の設定', '継続率・離反の確認（顧客分析）'],
        numbers: ['店舗別の目標達成率', '媒体別の新規・入会率・CPA', '継続率'],
    },
    admin: {
        label: 'オーナー',
        daily: ['ホームの「やること」（未突合・設定の不足）を確認'],
        weekly: ['全店舗の着地予測と目標の差', 'SalonOneで確認すること（入力漏れ）が残っていないか'],
        monthly: ['月初にインセンティブを確認（対象月を先月に）', '媒体別のCPA・ROASと広告予算', 'スタッフアカウントの発行・パスワード管理（設定）'],
        numbers: ['ブランド全体の売上・着地予測', '店舗別の達成率', '媒体別CPA・ROAS・入会率', '継続率'],
    },
};

const METRICS = [
    { name: '売上（会計済み）', def: '会計が完了した売上。日別・スタッフ別・支払い内訳と一致する基準', guide: '目標に対する進捗と着地予測で見る' },
    { name: '客単価', def: '売上 ÷ 来店数', guide: '前期間・前年と比べる' },
    { name: '次回予約率', def: '次回予約を取れた人数 ÷ 来店数（新規・既存それぞれ）', guide: '70%以上=緑、50%以上=黄、50%未満=赤で表示' },
    { name: '新規の入会率', def: '期間内に回数券・サブスクを契約した新規客 ÷ 新規予約数', guide: '媒体別・スタッフ別に比べる' },
    { name: 'キャンセル率', def: '（キャンセル＋無断）÷（来店＋キャンセル＋無断）', guide: '15%以上で注意表示' },
    { name: '継続率', def: '期間内に入会した顧客のうち契約が有効な人の割合', guide: '媒体・担当者別に見る' },
    { name: '日報の入力率', def: '日報を入力した人日 ÷（スタッフ数 × 経過日数）', guide: '100%を目指す（ホームで未入力者を表示）' },
    { name: '着地予測', def: '今月の実績 + 残り日数を曜日別の平均売上で見込んだ値', guide: '目標との差を見て打ち手を決める' },
];

export function init() {
    on('tab:shown', id => { if (id === 'guide') render(); });
    document.getElementById('guide-body')?.addEventListener('click', ev => {
        const btn = ev.target.closest('button[data-goto]');
        if (btn) switchTab(btn.dataset.goto);
    });
}

function list(items) {
    return `<ul class="guide-list">${items.map(i => `<li>${i}</li>`).join('')}</ul>`;
}

function routineBody(key) {
    const r = ROUTINES[key];
    return `
        <div class="guide-routine">
            <div><p class="guide-sub">毎日</p>${list(r.daily)}</div>
            <div><p class="guide-sub">毎週</p>${list(r.weekly)}</div>
            <div><p class="guide-sub">毎月</p>${list(r.monthly)}</div>
            <div><p class="guide-sub">見る数字</p>${list(r.numbers)}</div>
        </div>`;
}

function routineCard(key) {
    return `
    <details class="guide-details">
        <summary><span class="guide-role-chip">${ROUTINES[key].label}</span>のルーティン</summary>
        ${routineBody(key)}
    </details>`;
}

function render() {
    const el = document.getElementById('guide-body');
    if (!el) return;
    const role = state.session?.role || 'admin';
    const others = ['staff', 'store', 'manager', 'admin'].filter(k => k !== role);
    el.innerHTML = `
    <div class="premium-card home-card">
        <div class="home-card-head">
            <div class="flex items-center gap-2.5 min-w-0">
                <span class="home-card-icon gradient-primary"><i data-lucide="book-open" class="w-4 h-4 text-white"></i></span>
                <h3 class="home-card-title">使い方・業務分担</h3>
            </div>
        </div>
        <div class="guide-split">
            <div class="guide-split-col">
                <p class="guide-split-head">SalonOne</p>
                <p class="guide-split-lead">お客様の予約・会計を<b>記録する</b>場所（数字の出どころ）</p>
            </div>
            <div class="guide-split-arrow" aria-hidden="true">→</div>
            <div class="guide-split-col tool">
                <p class="guide-split-head">このダッシュボード</p>
                <p class="guide-split-lead">記録から<b>見る・振り返る・決める</b>場所。SalonOneにない数字だけ入力</p>
            </div>
        </div>
    </div>

    <div class="premium-card home-card">
        <p class="guide-title">あなた（${ROUTINES[role].label}）のルーティン</p>
        ${routineBody(role)}
    </div>

    <div class="premium-card home-card">
        <p class="guide-title">SalonOneでやること</p>
        <p class="guide-note">ここが正しく入力されていれば、このツールの数字は自動で揃います。</p>
        <ul class="guide-task-list">
            ${SALONONE_TASKS.map(t => `<li><p class="guide-task-what">${t.what}</p><p class="guide-task-why">→ ${t.why}</p></li>`).join('')}
        </ul>
    </div>

    <div class="premium-card home-card">
        <p class="guide-title">このツールでやること</p>
        <div class="guide-tool-grid">
            ${TOOL_TASKS.map(g => `<div><p class="guide-sub">${g.head}</p>${list(g.items)}</div>`).join('')}
        </div>
    </div>

    <div class="premium-card home-card">
        <p class="guide-title">数字の定義と目安</p>
        <dl class="guide-metrics">
            ${METRICS.map(m => `<div><dt>${m.name}</dt><dd>${m.def}<span>${m.guide}</span></dd></div>`).join('')}
        </dl>
    </div>

    <div class="premium-card home-card">
        <p class="guide-title">次回予約の自動推定（β）</p>
        <p class="guide-note">SalonOneの予約データから「来店した日の終わりまでに次の予約が入っていたか」を数えて、次回予約率を推定します。会計時にSalonOneで次回予約を登録していれば、日報の手入力とほぼ同じ値になります。日報入力の画面に推定値が出るので、確認してから保存してください（数字が大きく違う場合は、SalonOneでの予約登録の仕方を見直すきっかけになります）。</p>
    </div>

    <div class="premium-card home-card">
        <p class="guide-title">ほかの役割のルーティン</p>
        ${others.map(k => routineCard(k)).join('')}
    </div>`;
    if (window.lucide) lucide.createIcons();
}
