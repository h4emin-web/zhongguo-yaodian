import "@supabase/functions-js/edge-runtime.d.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};
const JSON_HEADERS = {
  ...CORS_HEADERS,
  "Content-Type": "application/json",
  "Cache-Control": "public, max-age=3600, s-maxage=3600"
};
const HANA_RATE_URL = "https://www.kebhana.com/cms/rate/wpfxd651_01i_01.do";

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: JSON_HEADERS
  });
}

function pad2(value: number) {
  return String(value).padStart(2, "0");
}

function kstYmd() {
  const now = new Date(Date.now() + 9 * 60 * 60 * 1000);
  return `${now.getUTCFullYear()}${pad2(now.getUTCMonth() + 1)}${pad2(now.getUTCDate())}`;
}

function decodeHtml(value: string) {
  return value
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;/g, "'");
}

function toPlainText(value: string) {
  return decodeHtml(value)
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function parseCells(rowHtml: string) {
  return [...rowHtml.matchAll(/<td\b[^>]*>([\s\S]*?)<\/td>/gi)].map((match) => toPlainText(match[1]));
}

function parseStandardAt(html: string) {
  const text = toPlainText(html);
  const match = text.match(/고시일시\s*:\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*(\d{1,2})시\s*(\d{1,2})분/);

  if (!match) {
    return { standardAt: "", standardAtText: "" };
  }

  const [, year, month, day, hour, minute] = match;

  return {
    standardAt: `${year}-${pad2(Number(month))}-${pad2(Number(day))}T${pad2(Number(hour))}:${pad2(Number(minute))}:00+09:00`,
    standardAtText: `${year}년 ${pad2(Number(month))}월 ${pad2(Number(day))}일 ${pad2(Number(hour))}시 ${pad2(Number(minute))}분`
  };
}

function extractRate(html: string, currency: string) {
  for (const match of html.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/gi)) {
    const rowHtml = match[1];
    const rowText = toPlainText(rowHtml);

    if (!rowText.includes(currency)) {
      continue;
    }

    const cells = parseCells(rowHtml);

    if (cells.length < 9) {
      continue;
    }

    return {
      currency,
      currencyName: cells[0],
      sendRate: cells[5],
      baseRate: cells[8]
    };
  }

  throw new Error(`${currency} 환율 행을 찾지 못했습니다.`);
}

async function fetchHanaRate(currency: string) {
  const ymd = kstYmd();
  const params = new URLSearchParams({
    curCd: currency,
    tmpInqStrDt: ymd,
    pbldDvCd: "3",
    pbldSqn: "",
    hid_key_data: "",
    inqStrDt: ymd,
    inqKindCd: "1",
    requestTarget: "searchContentDiv"
  });
  const response = await fetch(`${HANA_RATE_URL}?${params.toString()}`, {
    headers: {
      "Accept": "text/html, */*; q=0.01",
      "Referer": "https://www.kebhana.com/cont/mall/mall15/mall1501/index.jsp",
      "User-Agent": "Mozilla/5.0",
      "X-Requested-With": "XMLHttpRequest"
    }
  });

  if (!response.ok) {
    throw new Error(`하나은행 응답 오류: ${response.status}`);
  }

  const html = await response.text();
  const rate = extractRate(html, currency);
  const standard = parseStandardAt(html);

  return {
    ok: true,
    source: "Hana Bank",
    ...rate,
    ...standard,
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

  const body = await req.json().catch(() => ({}));
  const currency = String(body.currency || "USD").trim().toUpperCase();

  if (!/^[A-Z]{3}$/.test(currency)) {
    return json({ ok: false, error: "currency must be a 3-letter code." }, 400);
  }

  try {
    return json(await fetchHanaRate(currency));
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 502);
  }
});
