import "@supabase/functions-js/edge-runtime.d.ts";

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
const STOCK_CODE = "476830";
const STOCK_NAME = "알지노믹스";
const NAVER_REALTIME_URL = `https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:${STOCK_CODE}`;

type NaverStockData = {
  cd?: string;
  nm?: string;
  nv?: number;
  sv?: number;
  cv?: number;
  cr?: number;
  rf?: string;
  ms?: string;
  pcv?: number;
  ov?: number;
  hv?: number;
  lv?: number;
  aq?: number;
  aa?: number;
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

async function fetchStockPrice() {
  const response = await fetch(NAVER_REALTIME_URL, {
    headers: {
      "Accept": "application/json, text/plain, */*",
      "Referer": `https://finance.naver.com/item/main.naver?code=${STOCK_CODE}`,
      "User-Agent": "Mozilla/5.0"
    }
  });

  if (!response.ok) {
    throw new Error(`Naver Finance response error: ${response.status}`);
  }

  const payload = await response.json();
  const stock = payload?.result?.areas
    ?.find((area: { name?: string }) => area.name === "SERVICE_ITEM")
    ?.datas?.find((item: NaverStockData) => item.cd === STOCK_CODE) as NaverStockData | undefined;

  if (!stock) {
    throw new Error(`${STOCK_CODE} stock data was not found.`);
  }

  const direction = directionFromRf(stock.rf);
  const changeAbs = numberOrZero(stock.cv);
  const rateAbs = numberOrZero(stock.cr);

  return {
    ok: true,
    source: "Naver Finance",
    code: STOCK_CODE,
    name: stock.nm || STOCK_NAME,
    market: "KOSDAQ",
    price: numberOrZero(stock.nv),
    previousClose: numberOrZero(stock.sv || stock.pcv),
    change: direction.sign * changeAbs,
    changeAbs,
    changeRate: direction.sign * rateAbs,
    changeRateAbs: rateAbs,
    direction: direction.direction,
    directionText: direction.text,
    marketStatus: stock.ms || "",
    open: numberOrZero(stock.ov),
    high: numberOrZero(stock.hv),
    low: numberOrZero(stock.lv),
    volume: numberOrZero(stock.aq),
    tradedValue: numberOrZero(stock.aa),
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
    return json(await fetchStockPrice());
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 502);
  }
});
