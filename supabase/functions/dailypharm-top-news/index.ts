import "@supabase/functions-js/edge-runtime.d.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};
const JSON_HEADERS = {
  ...CORS_HEADERS,
  "Content-Type": "application/json",
  "Cache-Control": "public, max-age=300, s-maxage=300"
};
const DAILY_PHARM_URL = "https://www.dailypharm.com/";

type NewsItem = {
  rank: number;
  title: string;
  url: string;
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: JSON_HEADERS
  });
}

function decodeHtml(value: string) {
  return value
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/gi, "'");
}

function toPlainText(value: string) {
  return decodeHtml(value)
    .replace(/<br\s*\/?>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function absoluteDailyPharmUrl(value: string) {
  return new URL(decodeHtml(value), DAILY_PHARM_URL).toString();
}

function extractDailyTopNews(html: string) {
  const blockMatch = html.match(/<div\b[^>]*class="[^"]*today-top-5[^"]*"[^>]*>[\s\S]*?<ul\b[^>]*class="[^"]*today_list[^"]*"[^>]*>([\s\S]*?)<\/ul>/i);

  if (!blockMatch) {
    throw new Error("오늘의 TOP 10 영역을 찾지 못했습니다.");
  }

  const news: NewsItem[] = [];
  const itemPattern = /<li\b[^>]*>\s*<a\b[^>]*href="([^"]+)"[^>]*>\s*<span>\s*(\d+)\s*<\/span>([\s\S]*?)<\/a>\s*<\/li>/gi;

  for (const match of blockMatch[1].matchAll(itemPattern)) {
    const [, rawUrl, rawRank, rawTitle] = match;
    const title = toPlainText(rawTitle);
    const rank = Number(rawRank);

    if (!title || !Number.isFinite(rank)) {
      continue;
    }

    news.push({
      rank,
      title,
      url: absoluteDailyPharmUrl(rawUrl)
    });
  }

  if (news.length === 0) {
    throw new Error("오늘의 TOP 10 뉴스를 찾지 못했습니다.");
  }

  return news.slice(0, 10);
}

async function fetchDailyTopNews() {
  const response = await fetch(DAILY_PHARM_URL, {
    headers: {
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
      "User-Agent": "Mozilla/5.0"
    }
  });

  if (!response.ok) {
    throw new Error(`데일리팜 응답 오류: ${response.status}`);
  }

  const html = await response.text();

  return {
    ok: true,
    source: DAILY_PHARM_URL,
    fetchedAt: new Date().toISOString(),
    news: extractDailyTopNews(html)
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
    return json(await fetchDailyTopNews());
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 502);
  }
});
