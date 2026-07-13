import "@supabase/functions-js/edge-runtime.d.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const COM_CODE = Deno.env.get("ECOUNT_COM_CODE") ?? "";
const USER_ID = Deno.env.get("ECOUNT_USER_ID") ?? "";
const API_CERT_KEY = Deno.env.get("ECOUNT_API_CERT_KEY") ?? "";
const SEARCH_ENDPOINT = Deno.env.get("ECOUNT_IMPORT_COST_SEARCH_ENDPOINT") ?? "";

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" }
  });
}

function findExchangeRate(value: unknown): string {
  const text = typeof value === "string" ? value : JSON.stringify(value ?? "");
  const explicit = text.match(/@\s*([\d,]+(?:\.\d+)?)/);

  if (explicit) {
    return explicit[1].replace(/,/g, "");
  }

  const labeled = text.match(/환율[^0-9]*([\d,]+(?:\.\d+)?)/);

  if (labeled) {
    return labeled[1].replace(/,/g, "");
  }

  return "";
}

async function ecountLogin() {
  const zoneRes = await fetch(
    `https://sboapi.ecount.com/OAPI/V2/Zone?COM_CODE=${encodeURIComponent(COM_CODE)}`
  );
  const zoneJson = await zoneRes.json();
  const zone = zoneJson?.Data?.ZONE;

  if (!zone) {
    throw new Error(`ECOUNT Zone 조회 실패: ${JSON.stringify(zoneJson)}`);
  }

  const loginRes = await fetch(`https://sboapi${zone}.ecount.com/OAPI/V2/OAPILogin`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      COM_CODE,
      USER_ID,
      API_CERT_KEY,
      LAN_TYPE: "ko-KR",
      ZONE: zone
    })
  });
  const loginJson = await loginRes.json();
  const sessionId = loginJson?.Data?.Datas?.SESSION_ID;

  if (!sessionId) {
    throw new Error(`ECOUNT 로그인 실패: ${JSON.stringify(loginJson)}`);
  }

  return { zone, sessionId };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: CORS_HEADERS });
  }

  if (req.method !== "POST") {
    return json({ ok: false, error: "POST 요청만 지원합니다." }, 405);
  }

  const { poNo } = await req.json().catch(() => ({ poNo: "" }));
  const normalizedPoNo = typeof poNo === "string" ? poNo.trim() : "";

  if (!normalizedPoNo) {
    return json({ ok: false, error: "PO 번호가 필요합니다." }, 400);
  }

  if (!COM_CODE || !USER_ID || !API_CERT_KEY) {
    return json({ ok: false, error: "ECOUNT secrets가 설정되지 않았습니다." }, 500);
  }

  try {
    const { zone, sessionId } = await ecountLogin();

    if (!SEARCH_ENDPOINT) {
      return json({
        ok: false,
        error: "ECOUNT_IMPORT_COST_SEARCH_ENDPOINT가 설정되지 않았습니다. 수입비용전표검색 API endpoint를 Supabase secret으로 지정해야 합니다.",
        ecount: { loggedIn: true, zone }
      }, 501);
    }

    const endpoint = SEARCH_ENDPOINT.startsWith("/") ? SEARCH_ENDPOINT : `/${SEARCH_ENDPOINT}`;
    const res = await fetch(`https://sboapi${zone}.ecount.com${endpoint}?SESSION_ID=${sessionId}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        IO_NO: normalizedPoNo,
        MANAGE_NO: normalizedPoNo,
        MGMT_NO: normalizedPoNo,
        REMARKS: normalizedPoNo
      })
    });
    const payload = await res.json();
    const exchangeRate = findExchangeRate(payload);

    if (!exchangeRate) {
      return json({
        ok: false,
        error: "ERP 응답에서 환율을 찾지 못했습니다.",
        ecount: { loggedIn: true, zone },
        payload
      }, 404);
    }

    return json({
      ok: true,
      poNo: normalizedPoNo,
      exchangeRate,
      ecount: { loggedIn: true, zone }
    });
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 502);
  }
});
