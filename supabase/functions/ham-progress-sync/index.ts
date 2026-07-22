import "@supabase/functions-js/edge-runtime.d.ts";
import { createClient } from "@supabase/supabase-js";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};
const JSON_HEADERS = {
  ...CORS_HEADERS,
  "Content-Type": "application/json",
  "Cache-Control": "no-store"
};
const BUCKET = "haemin-ham-progress";
const FILE_PATH = "hidden.json";
const MAX_HIDDEN_IDS = 1500;

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: JSON_HEADERS
  });
}

function getSupabaseAdmin() {
  return createClient(
    Deno.env.get("SUPABASE_URL")!,
    Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
  );
}

async function ensureBucket() {
  const supabase = getSupabaseAdmin();
  const { data, error } = await supabase.storage.listBuckets();

  if (error) {
    throw error;
  }

  if (!data?.some((bucket) => bucket.name === BUCKET)) {
    const { error: createError } = await supabase.storage.createBucket(BUCKET, {
      public: false
    });

    if (createError) {
      throw createError;
    }
  }

  return supabase;
}

function normalizeHiddenIds(value: unknown) {
  if (!Array.isArray(value)) {
    return [];
  }

  return Array.from(new Set(
    value
      .filter((item): item is string => typeof item === "string")
      .map((item) => item.trim())
      .filter(Boolean)
  )).slice(0, MAX_HIDDEN_IDS);
}

async function loadHiddenIds() {
  const supabase = await ensureBucket();
  const { data, error } = await supabase.storage.from(BUCKET).download(FILE_PATH);

  if (error) {
    return json({ ok: true, hiddenIds: [], savedAt: null });
  }

  const parsed = JSON.parse(await data.text());
  return json({
    ok: true,
    hiddenIds: normalizeHiddenIds(parsed.hiddenIds),
    savedAt: parsed.savedAt ?? null
  });
}

async function saveHiddenIds(value: unknown) {
  const hiddenIds = normalizeHiddenIds(value);
  const supabase = await ensureBucket();
  const payload = JSON.stringify({
    hiddenIds,
    savedAt: new Date().toISOString()
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

  return json({ ok: true, hiddenIds });
}

async function clearHiddenIds() {
  const supabase = await ensureBucket();
  await supabase.storage.from(BUCKET).remove([FILE_PATH]);
  return json({ ok: true, hiddenIds: [] });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: CORS_HEADERS });
  }

  if (req.method !== "POST") {
    return json({ ok: false, error: "POST requests only." }, 405);
  }

  try {
    const body = await req.json();

    if (body.action === "load") {
      return await loadHiddenIds();
    }

    if (body.action === "save") {
      return await saveHiddenIds(body.hiddenIds);
    }

    if (body.action === "clear") {
      return await clearHiddenIds();
    }

    return json({ ok: false, error: "invalid action" }, 400);
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 500);
  }
});
