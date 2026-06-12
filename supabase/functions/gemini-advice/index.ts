// Gemini API 中継 Edge Function
// 目的: Gemini APIキーをブラウザ(localStorage)に置かず、Supabase Secret で秘匿する。
//
// デプロイ:
//   supabase functions deploy gemini-advice
//   supabase secrets set GEMINI_API_KEY=AIza....
//
// 呼び出し(フロント): POST {url}/functions/v1/gemini-advice  body: { prompt }
//   headers: apikey / Authorization: Bearer <anon key>
// レスポンス: { text } | { error }

const GEMINI_MODEL = Deno.env.get('GEMINI_MODEL') || 'gemini-2.0-flash';

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
    const apiKey = Deno.env.get('GEMINI_API_KEY');
    if (!apiKey) {
      return json({ error: 'GEMINI_API_KEY が未設定です（supabase secrets set GEMINI_API_KEY=...）' }, 500);
    }
    const { prompt } = await req.json();
    if (!prompt || typeof prompt !== 'string' || prompt.length > 20000) {
      return json({ error: 'prompt が不正です' }, 400);
    }

    const res = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MODEL}:generateContent?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
      },
    );
    const data = await res.json();
    const text = data?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) {
      return json({ error: data?.error?.message || 'Gemini からの応答が不正です' }, 502);
    }
    return json({ text });
  } catch (e) {
    return json({ error: String(e?.message || e) }, 500);
  }
});

function json(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
  });
}
