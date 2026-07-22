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
const FILE_PATH = "progress.json";
const LEGACY_HIDDEN_PATH = "hidden.json";
const MAX_STARRED_IDS = 1500;
const MAX_SEEN_IDS = 5000;
const MAX_HIGHLIGHT_ITEMS = 1500;
const MAX_HIGHLIGHTS_PER_FIELD = 40;

type HamHighlights = Record<string, { question: string[]; note: string[] }>;

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

function normalizeStringArray(value: unknown, max = MAX_STARRED_IDS) {
  if (!Array.isArray(value)) {
    return [];
  }

  return Array.from(new Set(
    value
      .filter((item): item is string => typeof item === "string")
      .map((item) => item.trim())
      .filter(Boolean)
  )).slice(0, max);
}

function normalizeHighlights(value: unknown) {
  if (!value || typeof value !== "object") {
    return {};
  }

  return Object.entries(value as Record<string, unknown>)
    .slice(0, MAX_HIGHLIGHT_ITEMS)
    .reduce<HamHighlights>((acc, [itemId, fields]) => {
      if (!fields || typeof fields !== "object") {
        return acc;
      }

      const record = fields as Record<string, unknown>;
      const question = normalizeStringArray(record.question, MAX_HIGHLIGHTS_PER_FIELD);
      const note = normalizeStringArray(record.note, MAX_HIGHLIGHTS_PER_FIELD);

      if (question.length > 0 || note.length > 0) {
        acc[itemId] = { question, note };
      }

      return acc;
    }, {});
}

async function clearLegacyHiddenFile(supabase: ReturnType<typeof getSupabaseAdmin>) {
  try {
    await supabase.storage.from(BUCKET).remove([LEGACY_HIDDEN_PATH]);
  } catch {
    // Legacy cleanup is best-effort only.
  }
}

async function loadProgress() {
  const supabase = await ensureBucket();
  await clearLegacyHiddenFile(supabase);

  const { data, error } = await supabase.storage.from(BUCKET).download(FILE_PATH);

  if (error) {
    return json({ ok: true, starredIds: [], seenIds: [], highlights: {}, savedAt: null });
  }

  const parsed = JSON.parse(await data.text());
  return json({
    ok: true,
    starredIds: normalizeStringArray(parsed.starredIds),
    seenIds: normalizeStringArray(parsed.seenIds, MAX_SEEN_IDS),
    highlights: normalizeHighlights(parsed.highlights),
    savedAt: parsed.savedAt ?? null
  });
}

async function saveProgress(starredIdsValue: unknown, highlightsValue: unknown, seenIdsValue: unknown) {
  const supabase = await ensureBucket();
  await clearLegacyHiddenFile(supabase);

  const starredIds = normalizeStringArray(starredIdsValue);
  const seenIds = normalizeStringArray(seenIdsValue, MAX_SEEN_IDS);
  const highlights = normalizeHighlights(highlightsValue);
  const payload = JSON.stringify({
    starredIds,
    seenIds,
    highlights,
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

  return json({ ok: true, starredIds, seenIds, highlights });
}

async function clearProgress() {
  const supabase = await ensureBucket();
  await supabase.storage.from(BUCKET).remove([FILE_PATH, LEGACY_HIDDEN_PATH]);
  return json({ ok: true, starredIds: [], seenIds: [], highlights: {} });
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
      return await loadProgress();
    }

    if (body.action === "save") {
      return await saveProgress(body.starredIds, body.highlights, body.seenIds);
    }

    if (body.action === "clear") {
      return await clearProgress();
    }

    return json({ ok: false, error: "invalid action" }, 400);
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 500);
  }
});
