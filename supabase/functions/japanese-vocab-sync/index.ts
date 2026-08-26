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
const BUCKET = "haemin-japanese-vocab";
const FILE_PATH = "words.json";
const MAX_ITEMS = 3000;
const MAX_TEXT_LENGTH = 300;
const MAX_MEMO_LENGTH = 2000;

type JapaneseVocabItem = {
  id: string;
  japanese: string;
  meaning: string;
  reading: string;
  memo: string;
  createdAt: string;
  updatedAt: string;
};

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

function normalizeText(value: unknown, max = MAX_TEXT_LENGTH) {
  return typeof value === "string" ? value.trim().slice(0, max) : "";
}

function normalizeItems(value: unknown) {
  if (!Array.isArray(value)) {
    return [];
  }

  return value.slice(0, MAX_ITEMS).reduce<JapaneseVocabItem[]>((items, rawItem) => {
    if (!rawItem || typeof rawItem !== "object") {
      return items;
    }

    const record = rawItem as Record<string, unknown>;
    const japanese = normalizeText(record.japanese);
    const meaning = normalizeText(record.meaning);
    const reading = normalizeText(record.reading);
    const memo = normalizeText(record.memo, MAX_MEMO_LENGTH);

    if (!japanese && !meaning) {
      return items;
    }

    const now = new Date().toISOString();

    items.push({
      id: normalizeText(record.id, 80) || crypto.randomUUID(),
      japanese,
      meaning,
      reading,
      memo,
      createdAt: normalizeText(record.createdAt, 40) || now,
      updatedAt: normalizeText(record.updatedAt, 40) || now
    });

    return items;
  }, []);
}

async function loadVocabulary() {
  const supabase = await ensureBucket();
  const { data, error } = await supabase.storage.from(BUCKET).download(FILE_PATH);

  if (error) {
    return json({ ok: true, items: [], savedAt: null });
  }

  const parsed = JSON.parse(await data.text());
  return json({
    ok: true,
    items: normalizeItems(parsed.items),
    savedAt: parsed.savedAt ?? null
  });
}

async function saveVocabulary(itemsValue: unknown) {
  const supabase = await ensureBucket();
  const items = normalizeItems(itemsValue);
  const payload = JSON.stringify({
    items,
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

  return json({ ok: true, items });
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
      return await loadVocabulary();
    }

    if (body.action === "save") {
      return await saveVocabulary(body.items);
    }

    return json({ ok: false, error: "invalid action" }, 400);
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 500);
  }
});
