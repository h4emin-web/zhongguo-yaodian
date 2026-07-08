// 업무일지 "내용" 텍스트를 Gemini(무료 API)로 짧게 요약해서 돌려준다.
// 프론트에서 { items: [{ id, text }] } 를 보내면 { ok, summaries: { [id]: summary } } 를 돌려준다.

import "@supabase/functions-js/edge-runtime.d.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const BATCH_SIZE = 40;

function buildPrompt(items: { id: string; text: string }[]): string {
  return `아래는 업무일지 항목들입니다. 각 항목의 "text"를 다음 규칙에 따라 처리하세요.

목표는 "요약(축약)"이 아니라 "정리(가독성 개선)"입니다. 내용을 줄이지 마세요.

규칙:
- 원문에 있는 모든 정보, 숫자, 단가, 회사명, 결론을 하나도 빠짐없이 그대로 유지하세요. 표현도 원문 그대로 최대한 유지하고, 절대 다른 말로 바꾸거나 생략하지 마세요.
- 문맥 없이 여러 내용이 한 문장으로 길게 이어져 있으면, 자연스러운 문장 단위로 줄바꿈해서 나누기만 하세요. 문장을 합치거나 재구성하지 말고, 순서와 표현을 원문 그대로 유지한 채 줄만 나누세요.
- 원문이 이미 짧거나(1~2문장), 이미 줄바꿈으로 잘 정리되어 있으면 절대 손대지 말고 원문 그대로 돌려주세요.
- 원문에 없는 문장, 해석, 결론, 예측을 추가하지 마세요.
- 회사명, 원료명, 단가, 숫자는 원문 표기 그대로 정확히 옮기세요 (오타 수정 금지, 다른 단어로 바꾸기 금지).
- 각 항목의 id는 입력과 동일하게 그대로 응답하세요.
- 오직 JSON 배열만 응답하세요. 다른 설명 텍스트는 포함하지 마세요.
- 형식: [{"id": "...", "summary": "..."}]

입력:
${JSON.stringify(items)}`;
}

function extractJsonArray(text: string): unknown[] {
  try {
    return JSON.parse(text);
  } catch {
    const match = text.match(/\[[\s\S]*\]/);
    return match ? JSON.parse(match[0]) : [];
  }
}

async function summarizeBatch(
  items: { id: string; text: string }[],
  apiKey: string,
  model: string
): Promise<{ id: string; summary: string }[]> {
  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-goog-api-key": apiKey
      },
      body: JSON.stringify({
        contents: [{ parts: [{ text: buildPrompt(items) }] }],
        generationConfig: {
          responseMimeType: "application/json",
          temperature: 0
        }
      })
    }
  );

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Gemini API error ${response.status}: ${errText}`);
  }

  const data = await response.json();
  const text = data?.candidates?.[0]?.content?.parts?.[0]?.text ?? "[]";
  const parsed = extractJsonArray(text);

  return Array.isArray(parsed) ? (parsed as { id: string; summary: string }[]) : [];
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: CORS_HEADERS });
  }

  try {
    const { items } = await req.json();

    if (!Array.isArray(items) || items.length === 0) {
      return new Response(JSON.stringify({ ok: false, error: "items required" }), {
        status: 400,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" }
      });
    }

    const apiKey = Deno.env.get("GEMINI_API_KEY");

    if (!apiKey) {
      return new Response(JSON.stringify({ ok: false, error: "GEMINI_API_KEY not configured" }), {
        status: 500,
        headers: { ...CORS_HEADERS, "Content-Type": "application/json" }
      });
    }

    const model = Deno.env.get("GEMINI_MODEL") || "gemini-2.5-flash-lite";
    const summaries: Record<string, string> = {};

    for (let i = 0; i < items.length; i += BATCH_SIZE) {
      const batch = items.slice(i, i + BATCH_SIZE);

      try {
        const result = await summarizeBatch(batch, apiKey, model);

        for (const entry of result) {
          if (entry && entry.id) {
            summaries[entry.id] = entry.summary ?? "";
          }
        }
      } catch (batchError) {
        console.error("batch failed", batchError);
      }
    }

    return new Response(JSON.stringify({ ok: true, summaries }), {
      headers: { ...CORS_HEADERS, "Content-Type": "application/json" }
    });
  } catch (error) {
    return new Response(JSON.stringify({ ok: false, error: String(error) }), {
      status: 500,
      headers: { ...CORS_HEADERS, "Content-Type": "application/json" }
    });
  }
});
