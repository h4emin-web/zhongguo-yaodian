// 업무일지(파싱된 데이터 + AI 요약 캐시)를 기기 간에 공유하기 위한 동기화 함수.
// 비공개 Storage 버킷(haemin-worklog)에 service_role 권한으로만 접근한다.
// 프론트는 이 함수를 통해서만 데이터를 읽고 쓰며, 버킷 자체는 공개 URL이 없다.

import "@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "@supabase/supabase-js";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const BUCKET = "haemin-worklog";
const FILE_PATH = "current.json";

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" }
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: CORS_HEADERS });
  }

  const supabase = createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
  );

  try {
    const { action, people, aiSummaries, expiresAt } = await req.json();

    if (action === "save") {
      const payload = JSON.stringify({
        people: people ?? [],
        aiSummaries: aiSummaries ?? {},
        expiresAt,
        savedAt: Date.now()
      });

      const { error } = await supabase.storage
        .from(BUCKET)
        .upload(FILE_PATH, new Blob([payload], { type: "application/json" }), {
          upsert: true,
          contentType: "application/json"
        });

      if (error) {
        throw error;
      }

      return json({ ok: true });
    }

    if (action === "load") {
      const { data, error } = await supabase.storage.from(BUCKET).download(FILE_PATH);

      if (error) {
        return json({ ok: true, data: null });
      }

      const text = await data.text();
      const parsed = JSON.parse(text);

      if (!parsed.expiresAt || Date.now() > parsed.expiresAt) {
        return json({ ok: true, data: null });
      }

      return json({ ok: true, data: parsed });
    }

    if (action === "clear") {
      await supabase.storage.from(BUCKET).remove([FILE_PATH]);
      return json({ ok: true });
    }

    return json({ ok: false, error: "invalid action" }, 400);
  } catch (error) {
    return json({ ok: false, error: String(error) }, 500);
  }
});
