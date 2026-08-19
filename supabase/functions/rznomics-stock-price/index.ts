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
const DEFAULT_STOCK_CODE = "476830";
const STOCKS: Record<string, { name: string; market: string }> = {
  "476830": { name: "알지노믹스", market: "KOSDAQ" },
  "066570": { name: "LG전자", market: "KOSPI" }
};

type StockRequest = {
  code?: unknown;
  name?: unknown;
  market?: unknown;
};

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

function stringOrEmpty(value: unknown) {
  return typeof value === "string" ? value.trim() : "";
}

function sanitizeStockCode(value: unknown) {
  const code = stringOrEmpty(value).replace(/\D/g, "");
  return /^\d{6}$/.test(code) ? code : DEFAULT_STOCK_CODE;
}

function getStockMeta(stockCode: string, input: StockRequest) {
  const knownStock = STOCKS[stockCode];

  return {
    name: knownStock?.name || stringOrEmpty(input.name) || stockCode,
    market: knownStock?.market || stringOrEmpty(input.market)
  };
}

async function fetchStockPrice(input: StockRequest = {}) {
  const stockCode = sanitizeStockCode(input.code);
  const stockMeta = getStockMeta(stockCode, input);
  const naverRealtimeUrl = `https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:${stockCode}`;

  const response = await fetch(naverRealtimeUrl, {
    headers: {
      "Accept": "application/json, text/plain, */*",
      "Referer": `https://finance.naver.com/item/main.naver?code=${stockCode}`,
      "User-Agent": "Mozilla/5.0"
    }
  });

  if (!response.ok) {
    throw new Error(`Naver Finance response error: ${response.status}`);
  }

  const payload = await response.json();
  const stock = payload?.result?.areas
    ?.find((area: { name?: string }) => area.name === "SERVICE_ITEM")
    ?.datas?.find((item: NaverStockData) => item.cd === stockCode) as NaverStockData | undefined;

  if (!stock) {
    throw new Error(`${stockCode} stock data was not found.`);
  }

  const direction = directionFromRf(stock.rf);
  const changeAbs = numberOrZero(stock.cv);
  const rateAbs = numberOrZero(stock.cr);

  return {
    ok: true,
    source: "Naver Finance",
    code: stockCode,
    name: stock.nm || stockMeta.name,
    market: stockMeta.market,
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
    let body: StockRequest = {};

    try {
      body = await req.json();
    } catch {
      body = {};
    }

    return json(await fetchStockPrice(body));
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 502);
  }
});
