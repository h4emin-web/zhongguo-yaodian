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
const NAVER_INDEX_URL = "https://polling.finance.naver.com/api/realtime?query=SERVICE_INDEX:KOSPI,KOSDAQ";

type NaverIndexData = {
  cd?: string;
  nv?: number;
  cv?: number;
  cr?: number;
  rf?: string;
  ms?: string;
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: JSON_HEADERS
  });
}

function directionFromRf(rf: string | undefined) {
  if (rf === "1" || rf === "2") {
    return { direction: "up", text: "상승", sign: 1 };
  }

  if (rf === "4" || rf === "5") {
    return { direction: "down", text: "하락", sign: -1 };
  }

  return { direction: "flat", text: "보합", sign: 0 };
}

function numberOrZero(value: unknown) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function normalizeIndex(item: NaverIndexData) {
  const direction = directionFromRf(item.rf);
  const changeAbs = numberOrZero(item.cv);
  const rateAbs = numberOrZero(item.cr);

  return {
    code: item.cd || "",
    value: numberOrZero(item.nv) / 100,
    change: direction.sign * changeAbs / 100,
    changeAbs: changeAbs / 100,
    changeRate: direction.sign * rateAbs,
    changeRateAbs: rateAbs,
    direction: direction.direction,
    directionText: direction.text,
    marketStatus: item.ms || ""
  };
}

async function fetchMarketIndexTrend() {
  const response = await fetch(NAVER_INDEX_URL, {
    headers: {
      "Accept": "application/json, text/plain, */*",
      "Referer": "https://finance.naver.com/sise/",
      "User-Agent": "Mozilla/5.0"
    }
  });

  if (!response.ok) {
    throw new Error(`Naver Finance response error: ${response.status}`);
  }

  const payload = await response.json();
  const indices = payload?.result?.areas
    ?.find((area: { name?: string }) => area.name === "SERVICE_INDEX")
    ?.datas as NaverIndexData[] | undefined;

  if (!Array.isArray(indices) || indices.length === 0) {
    throw new Error("KOSPI/KOSDAQ index data was not found.");
  }

  return {
    ok: true,
    source: "Naver Finance",
    indices: indices
      .filter((item) => item.cd === "KOSPI" || item.cd === "KOSDAQ")
      .map(normalizeIndex),
    standardAt: payload?.result?.time ? new Date(payload.result.time).toISOString() : "",
    fetchedAt: new Date().toISOString()
  };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: CORS_HEADERS });
  }

  if (req.method !== "POST") {
    return json({ ok: false, error: "POST requests only." }, 405);
  }

  try {
    return json(await fetchMarketIndexTrend());
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 502);
  }
});
