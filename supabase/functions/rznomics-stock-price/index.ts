import "@supabase/functions-js/edge-runtime.d.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};
const JSON_HEADERS = {
  ...CORS_HEADERS,
  "Content-Type": "application/json; charset=utf-8",
  "Cache-Control": "no-store"
};
type StockRequest = {
  code?: unknown;
  name?: unknown;
  market?: unknown;
  query?: unknown;
};

type StockMeta = {
  code: string;
  name: string;
  market: string;
};

type NaverSearchItem = {
  code?: string;
  name?: string;
  typeCode?: string;
  typeName?: string;
  nationCode?: string;
  category?: string;
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
  countOfListedStock?: number;
};

type NaverNewsItem = {
  title?: string;
  titleFull?: string;
  mobileNewsUrl?: string;
  newsUrl?: string;
  officeName?: string;
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
  return typeof value === "string" || typeof value === "number" ? String(value).trim() : "";
}

function cleanText(value: unknown) {
  return stringOrEmpty(value)
    .replace(/<[^>]*>/g, " ")
    .replace(/&quot;/gi, "\"")
    .replace(/&#34;/g, "\"")
    .replace(/&#x22;/gi, "\"")
    .replace(/&apos;/gi, "'")
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/gi, "'")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&nbsp;/gi, " ")
    .replace(/&[a-z0-9#]+;/gi, "")
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizeStockCode(value: unknown) {
  const code = stringOrEmpty(value).replace(/\D/g, "");
  return /^\d{6}$/.test(code) ? code : "";
}

function normalizeForCompare(value: unknown) {
  return stringOrEmpty(value).replace(/\s+/g, "").toLowerCase();
}

function requestQuery(input: StockRequest) {
  return stringOrEmpty(input.query) || stringOrEmpty(input.code) || stringOrEmpty(input.name);
}

function isKoreanStockItem(item: NaverSearchItem) {
  return item.category === "stock"
    && (!item.nationCode || item.nationCode === "KOR")
    && Boolean(sanitizeStockCode(item.code));
}

function stockMetaFromSearchItem(item: NaverSearchItem): StockMeta {
  const stockCode = sanitizeStockCode(item.code);

  return {
    code: stockCode,
    name: cleanText(item.name) || stockCode,
    market: cleanText(item.typeCode) || cleanText(item.typeName)
  };
}

async function resolveStockMeta(input: StockRequest): Promise<StockMeta> {
  const query = requestQuery(input);
  const directCode = sanitizeStockCode(query) || sanitizeStockCode(input.code);
  const normalizedQuery = normalizeForCompare(query);

  if (!query) {
    throw new Error("검색어를 입력하세요.");
  }

  try {
    const searchUrl = `https://ac.stock.naver.com/ac?q=${encodeURIComponent(query)}&target=stock,ipo,index,marketindicator`;
    const response = await fetch(searchUrl, {
      headers: {
        "Accept": "application/json, text/plain, */*",
        "Referer": "https://finance.naver.com/",
        "User-Agent": "Mozilla/5.0"
      }
    });

    if (!response.ok) {
      throw new Error(`Naver stock search response error: ${response.status}`);
    }

    const payload = await response.json();
    const items = Array.isArray(payload?.items) ? payload.items as NaverSearchItem[] : [];
    const exactCodeMatch = directCode
      ? items.find((item) => isKoreanStockItem(item) && sanitizeStockCode(item.code) === directCode)
      : undefined;
    const exactNameMatch = items.find((item) => isKoreanStockItem(item) && normalizeForCompare(item.name) === normalizedQuery);
    const firstStockMatch = items.find((item) => isKoreanStockItem(item));
    const selected = exactCodeMatch || exactNameMatch || firstStockMatch;

    if (selected) {
      return stockMetaFromSearchItem(selected);
    }
  } catch (error) {
    if (!directCode) {
      throw error;
    }
  }

  if (directCode) {
    return {
      code: directCode,
      name: cleanText(input.name) || directCode,
      market: cleanText(input.market)
    };
  }

  throw new Error(`${query} 종목을 찾지 못했습니다.`);
}

async function fetchStockPrice(input: StockRequest = {}) {
  const stockMeta = await resolveStockMeta(input);
  const stockCode = stockMeta.code;
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
  const price = numberOrZero(stock.nv);
  const listedShares = numberOrZero(stock.countOfListedStock);
  const news = await fetchLatestStockNews(stockCode).catch(() => null);

  return {
    ok: true,
    source: "Naver Finance",
    code: stockCode,
    name: cleanText(stockMeta.name) || cleanText(stock.nm) || stockCode,
    market: stockMeta.market,
    price,
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
    listedShares,
    marketCap: price * listedShares,
    newsTitle: news?.title || "",
    newsUrl: news?.url || "",
    standardAt: payload?.result?.time ? new Date(payload.result.time).toISOString() : "",
    fetchedAt: new Date().toISOString()
  };
}

async function fetchLatestStockNews(stockCode: string) {
  const newsUrl = `https://m.stock.naver.com/api/news/stock/${stockCode}?pageSize=1&page=1`;
  const response = await fetch(newsUrl, {
    headers: {
      "Accept": "application/json, text/plain, */*",
      "Referer": `https://m.stock.naver.com/domestic/stock/${stockCode}/news`,
      "User-Agent": "Mozilla/5.0"
    }
  });

  if (!response.ok) {
    throw new Error(`Naver stock news response error: ${response.status}`);
  }

  const payload = await response.json();
  const newsGroups = Array.isArray(payload) ? payload : [];
  const firstNews = newsGroups
    .flatMap((group: { items?: NaverNewsItem[] }) => Array.isArray(group.items) ? group.items : [])
    .find((item: NaverNewsItem) => stringOrEmpty(item.titleFull) || stringOrEmpty(item.title)) as NaverNewsItem | undefined;

  if (!firstNews) {
    return null;
  }

  return {
    title: cleanText(firstNews.titleFull) || cleanText(firstNews.title),
    url: stringOrEmpty(firstNews.mobileNewsUrl) || stringOrEmpty(firstNews.newsUrl)
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
