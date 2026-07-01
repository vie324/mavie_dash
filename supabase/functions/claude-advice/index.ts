// Claude (Anthropic) API 中継 Edge Function
// 目的: Anthropic APIキーをブラウザ(localStorage)に置かず、Supabase Secret で秘匿する。
//
// デプロイ:
//   supabase functions deploy claude-advice
//   supabase secrets set ANTHROPIC_API_KEY=sk-ant-....
//   （任意）supabase secrets set ANTHROPIC_MODEL=claude-opus-4-8
//
// 呼び出し(フロント): POST {url}/functions/v1/claude-advice  body: { prompt }
//   headers: apikey / Authorization: Bearer <anon key>
// レスポンス: { text } | { error }

const ANTHROPIC_MODEL = Deno.env.get('ANTHROPIC_MODEL') || 'claude-opus-4-8';
const ANTHROPIC_MAX_TOKENS = Number(Deno.env.get('ANTHROPIC_MAX_TOKENS') || '2048');

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders });
  }
  try {
    const apiKey = Deno.env.get('ANTHROPIC_API_KEY');
    if (!apiKey) {
      return json({ error: 'ANTHROPIC_API_KEY が未設定です（supabase secrets set ANTHROPIC_API_KEY=...）' }, 500);
    }
    const { prompt } = await req.json();
    if (!prompt || typeof prompt !== 'string' || prompt.length > 20000) {
      return json({ error: 'prompt が不正です' }, 400);
    }

    // サーバー側呼び出しのため anthropic-dangerous-direct-browser-access は不要
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: ANTHROPIC_MODEL,
        max_tokens: ANTHROPIC_MAX_TOKENS,
        messages: [{ role: 'user', content: prompt }],
      }),
    });
    const data = await res.json();
    if (!res.ok || data?.type === 'error') {
      return json({ error: data?.error?.message || `Anthropic API エラー (HTTP ${res.status})` }, 502);
    }
    if (data?.stop_reason === 'refusal') {
      return json({ error: 'この内容には回答できませんでした' }, 502);
    }
    const text = (data?.content || []).find((b: { type?: string }) => b.type === 'text')?.text;
    if (!text) {
      return json({ error: 'Claude からの応答が不正です' }, 502);
    }
    return json({ text });
  } catch (e) {
    return json({ error: String((e as Error)?.message || e) }, 500);
  }
});

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
  });
}
