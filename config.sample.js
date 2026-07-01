/* ---------------------------------------------------------------------
 * vie Dashboard フロント設定サンプル
 *
 * このファイルを config.js としてコピーすると、バックエンド接続先を
 * コードに埋め込めます（各端末で設定タブから入力する手間が省けます）。
 *
 * Vercel デプロイでは scripts/gen-config.mjs が環境変数から config.js を
 * 自動生成するため、通常このファイルを手で編集する必要はありません。
 *
 * localStorage（設定タブでの入力）が優先されます。ここは「デプロイ既定値」。
 * anon key は公開して問題ないキーですが、RLSが有効であることが前提です。
 * ------------------------------------------------------------------- */
window.__VIE_CONFIG__ = {
    backendMode: 'supabase',                 // 'supabase' | 'gas'
    supabaseUrl: 'https://xxxx.supabase.co',
    supabaseAnonKey: 'eyJhbGciOi...(anon public key)',
};
