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
const BUCKET = "haemin-password-vault";
const FILE_PATH = "vault.json";

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

function timingSafeEqual(a: string, b: string) {
  const encoder = new TextEncoder();
  const left = encoder.encode(a);
  const right = encoder.encode(b);
  const length = Math.max(left.length, right.length);
  let diff = left.length ^ right.length;

  for (let index = 0; index < length; index += 1) {
    diff |= (left[index] ?? 0) ^ (right[index] ?? 0);
  }

  return diff === 0;
}

function requireAccessCode(value: unknown) {
  const expected = Deno.env.get("PASSWORD_VAULT_ACCESS_CODE");

  if (!expected || typeof value !== "string" || !timingSafeEqual(value, expected)) {
    throw new Response(JSON.stringify({ ok: false, error: "비밀번호가 맞지 않습니다." }), {
      status: 403,
      headers: JSON_HEADERS
    });
  }
}

function requireAdminKey(value: unknown) {
  const expected = Deno.env.get("PASSWORD_VAULT_ADMIN_KEY");

  if (!expected || typeof value !== "string" || !timingSafeEqual(value, expected)) {
    throw new Response(JSON.stringify({ ok: false, error: "관리자 키가 맞지 않습니다." }), {
      status: 403,
      headers: JSON_HEADERS
    });
  }
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

async function loadVault(accessCode: unknown) {
  requireAccessCode(accessCode);
  const supabase = getSupabaseAdmin();
  const { data, error } = await supabase.storage.from(BUCKET).download(FILE_PATH);

  if (error) {
    return json({ ok: true, data: null });
  }

  return json({ ok: true, data: JSON.parse(await data.text()) });
}

async function saveVault(adminKey: unknown, vault: unknown) {
  requireAdminKey(adminKey);

  if (!vault || typeof vault !== "object") {
    return json({ ok: false, error: "저장할 데이터가 없습니다." }, 400);
  }

  const supabase = await ensureBucket();
  const payload = JSON.stringify({
    ...(vault as Record<string, unknown>),
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

  return json({ ok: true });
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
      return await loadVault(body.accessCode);
    }

    if (body.action === "save") {
      return await saveVault(body.adminKey, body.vault);
    }

    return json({ ok: false, error: "invalid action" }, 400);
  } catch (error) {
    if (error instanceof Response) {
      return error;
    }

    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 500);
  }
});
