import "@supabase/functions-js/edge-runtime.d.ts";

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS"
};

const COM_CODE = Deno.env.get("ECOUNT_COM_CODE") ?? "";
const USER_ID = Deno.env.get("ECOUNT_USER_ID") ?? "";
const PASSWORD = Deno.env.get("ECOUNT_PASSWORD") ?? "";
const LOGIN_BASE = "https://loginbb.ecount.com";
const PAGE_SIZE = 25;

type EcountLogin = {
  sessionId: string;
  cookie: string;
};

type DateRange = {
  from: string;
  to: string;
};

type EcountSlipRow = Record<string, unknown>;

type ScanMode = {
  label: string;
  dateRange: DateRange;
  remarks: string;
  maxPages: number;
};

type ScanDebug = {
  mode: string;
  dateRange: DateRange;
  totalPages: number;
  scannedPages: number;
  rowsSeen: number;
  poRowsSeen: number;
  firstPageKeys: string[];
  firstPageDataKeys: string[];
  samplePoRowKeys: string[];
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS_HEADERS, "Content-Type": "application/json" }
  });
}

function decodeMaybeBase64(text: string): unknown {
  const trimmed = text.trim();

  if (/^[A-Za-z0-9+/=]+$/.test(trimmed) && trimmed.length > 20) {
    try {
      const decoded = new TextDecoder().decode(Uint8Array.from(atob(trimmed), (char) => char.charCodeAt(0)));
      return JSON.parse(decoded);
    } catch {
      // ECOUNT sometimes returns plain text/HTML instead of encoded JSON.
    }
  }

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

async function tryDecompress(bytes: Uint8Array, format: CompressionFormat | "deflate-raw") {
  try {
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream(format as CompressionFormat));
    const buffer = await new Response(stream).arrayBuffer();
    return new TextDecoder().decode(buffer);
  } catch {
    return "";
  }
}

function bytesToHex(bytes: Uint8Array, length = 32) {
  return [...bytes.slice(0, length)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
}

async function decodeResponse(response: Response) {
  const bytes = new Uint8Array(await response.arrayBuffer());
  const plainText = new TextDecoder().decode(bytes);
  const candidates = [
    plainText,
    await tryDecompress(bytes, "gzip"),
    await tryDecompress(bytes, "deflate"),
    await tryDecompress(bytes, "deflate-raw")
  ].filter(Boolean);

  for (const candidate of candidates) {
    const decoded = decodeMaybeBase64(candidate);

    if (decoded && typeof decoded === "object") {
      return {
        text: candidate,
        data: decoded,
        rawHex: bytesToHex(bytes)
      };
    }
  }

  return {
    text: candidates[0] ?? "",
    data: decodeMaybeBase64(candidates[0] ?? ""),
    rawHex: bytesToHex(bytes)
  };
}

function toUtf8Bytes(value: string) {
  return [...new TextEncoder().encode(value)];
}

function encryptPassword(value: string, key: string) {
  return toUtf8Bytes(value)
    .map((byte, index) => (byte ^ key.charCodeAt(index % key.length)).toString(16).padStart(2, "0"))
    .join("");
}

function buildEncryptedPassword(password: string) {
  const token = String(Date.now()).slice(0, 7);
  return `${token.length}${token}${encryptPassword(password, token)}`;
}

class CookieJar {
  #cookies = new Map<string, string>();

  set(name: string, value: string) {
    this.#cookies.set(name, value);
  }

  addFromResponse(response: Response) {
    const headers = response.headers as Headers & { getSetCookie?: () => string[] };
    const setCookies = headers.getSetCookie ? headers.getSetCookie() : [];

    for (const cookie of setCookies) {
      const first = cookie.split(";")[0];
      const separator = first.indexOf("=");

      if (separator > 0) {
        this.#cookies.set(first.slice(0, separator), first.slice(separator + 1));
      }
    }
  }

  get(name: string) {
    return this.#cookies.get(name) ?? "";
  }

  header() {
    return [...this.#cookies.entries()].map(([key, value]) => `${key}=${value}`).join("; ");
  }
}

async function ecountFetch(jar: CookieJar, url: string, options: RequestInit = {}) {
  const headers = new Headers(options.headers);
  const cookie = jar.header();

  if (cookie) {
    headers.set("Cookie", cookie);
  }

  const response = await fetch(url, {
    ...options,
    headers,
    redirect: "manual"
  });
  jar.addFromResponse(response);

  const decoded = await decodeResponse(response);

  if (response.status >= 400) {
    throw new Error(`ECOUNT request failed: ${response.status}`);
  }

  return {
    response,
    headers: Object.fromEntries(response.headers.entries()),
    text: decoded.text,
    data: decoded.data,
    rawHex: decoded.rawHex
  };
}

function getDataObject(value: unknown) {
  if (value && typeof value === "object" && "Data" in value) {
    return (value as { Data?: Record<string, unknown> }).Data ?? {};
  }

  return {};
}

function parseHiddenInputs(html: string) {
  const inputs: Record<string, string> = {};
  const matches = html.matchAll(/<input[^>]+>/g);

  for (const match of matches) {
    const tag = match[0];
    const name = tag.match(/name="([^"]*)"/)?.[1];

    if (!name) {
      continue;
    }

    const value = tag.match(/value="([^"]*)"/)?.[1] ?? "";
    inputs[name] = value
      .replace(/&quot;/g, "\"")
      .replace(/&#x2B;/g, "+")
      .replace(/&amp;/g, "&");
  }

  return inputs;
}

function getSessionIdFromCookie(cookie: string) {
  return decodeURIComponent(cookie).split("=")[0] ?? "";
}

function getSessionIdFromLocation(location: string) {
  if (!location) {
    return "";
  }

  const decoded = location.replace(/&amp;/g, "&");
  const match = decoded.match(/[?&]ec_req_sid=([^&#]+)/);
  return match ? decodeURIComponent(match[1]) : "";
}

function toHex(value: string) {
  return [...new TextEncoder().encode(value)]
    .map((byte) => byte.toString(16).padStart(2, "0"))
    .join("");
}

function buildSessionCookieValue(sessionId: string) {
  const encodedUserId = encodeURIComponent(USER_ID).toLowerCase();
  return `${sessionId}=${toHex(`${COM_CODE}|${encodedUserId}`)}`;
}

function formatYmd(date: Date) {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}${month}${day}`;
}

function kstDate() {
  return new Date(Date.now() + 9 * 60 * 60 * 1000);
}

function getRecentDateRange() {
  const from = kstDate();
  const to = kstDate();
  from.setUTCDate(from.getUTCDate() - 30);
  to.setUTCDate(to.getUTCDate() + 30);

  return {
    from: formatYmd(from),
    to: formatYmd(to)
  };
}

function getCurrentYearDateRange() {
  const now = kstDate();
  const year = now.getUTCFullYear();

  return {
    from: `${year}0101`,
    to: `${year}1231`
  };
}

async function continueNewDeviceLogin(jar: CookieJar, hidden: Record<string, string>) {
  const zoneInfo = await ecountFetch(jar, `${LOGIN_BASE}/ec5/api/app.login/action/GetZoneInfoAction?&xce=none`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Referer": `${LOGIN_BASE}/ec5/view/app.login/erp_login`
    },
    body: "{}"
  });
  const zoneInfoData = getDataObject(zoneInfo.data);
  const body = new URLSearchParams({
    ...hidden,
    process_ing: hidden.process_ing || "N",
    ts: String(zoneInfoData.ts ?? hidden.ts ?? ""),
    lan_type: hidden.lan_type || hidden.login_lantype || "ko-KR",
    access_site: hidden.access_site || "ECOUNT",
    adminChecked: hidden.adminChecked || "N",
    deviceFlag: "N",
    deviceKey: hidden.deviceKey || "",
    closeSessionFlag: hidden.closeSessionFlag || "",
    forceLogout: hidden.forceLogout || "",
    qrNumber: hidden.qrNumber || "",
    platformVersion: hidden.platformVersion || "19"
  });

  const response = await ecountFetch(jar, `${LOGIN_BASE}/ec5/view/app.login/erp_login_processor`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "Origin": LOGIN_BASE,
      "Referer": `${LOGIN_BASE}/ec5/view/app.login/erp_login`
    },
    body: body.toString()
  });
  const location = response.response.headers.get("location") ?? "";
  let sessionId = getSessionIdFromLocation(location);

  if (!sessionId) {
    sessionId = getSessionIdFromCookie(jar.get("ECOUNT_SessionId"));
  }

  if (sessionId && !jar.get("ECOUNT_SessionId")) {
    jar.set("ECOUNT_SessionId", buildSessionCookieValue(sessionId));
  }

  return sessionId;
}

async function ecountWebLogin(): Promise<EcountLogin> {
  const jar = new CookieJar();

  await ecountFetch(jar, `${LOGIN_BASE}/ec5/view/app.login/erp_login?lan_type=ko-KR`);

  const device = await ecountFetch(jar, `${LOGIN_BASE}/ec5/api/app.login/action/SetDeviceKeyCookieAction`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ DOMAIN: ".ecount.com" })
  });
  const deviceKey = String(getDataObject(device.data).DEVICEKEY ?? "");

  if (deviceKey) {
    jar.set("EcNewDeviceKey", encodeURIComponent(deviceKey));
  }

  const zone = await ecountFetch(jar, `${LOGIN_BASE}/ec5/api/app.login/action/GetZoneAction?&xce=none`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ com_code: COM_CODE })
  });
  const zoneData = getDataObject(zone.data);

  const zoneInfo = await ecountFetch(jar, `${LOGIN_BASE}/ec5/api/app.login/action/GetZoneInfoAction?&xce=none`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: "{}"
  });
  const zoneInfoData = getDataObject(zoneInfo.data);
  const loginBody = new URLSearchParams({
    com_code: COM_CODE,
    id: USER_ID,
    spasswd: buildEncryptedPassword(PASSWORD),
    loginck: "",
    logintimeinck: "",
    lan_type: "ko-KR",
    process_ing: "N",
    ts: String(zoneInfoData.ts ?? ""),
    zone: String(zoneData.ZONE ?? "BB"),
    db_shard_no: String(zoneData.DB_SHARD_NO ?? "1"),
    domain: String(zoneData.DOMAIN ?? ".ecount.com"),
    m_flag: "",
    error_code: "",
    error_msg: "",
    con_type: "",
    uname: "",
    login_days: "",
    brn_cd: "",
    xurl_rd: "",
    login_type: "0",
    access_site: "ECOUNT",
    hidExSID: "",
    hidFromCs: "",
    hidCnFlag: "",
    adminChecked: "N",
    noHistoryIpReason: "",
    deviceFlag: "",
    deviceKey: "",
    forceLogout: "",
    closeSessionFlag: "",
    sip: String(zoneData.SIP ?? ""),
    qrNumber: "",
    s_second_auth_sid: "",
    platformVersion: "19"
  });

  const login = await ecountFetch(jar, `${LOGIN_BASE}/ec5/view/app.login/erp_login_processor`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "Origin": LOGIN_BASE,
      "Referer": `${LOGIN_BASE}/ec5/view/app.login/erp_login?lan_type=ko-KR`
    },
    body: loginBody.toString()
  });

  const sessionCookie = jar.get("ECOUNT_SessionId");
  const sessionId = getSessionIdFromCookie(sessionCookie);

  if (sessionId) {
    return { sessionId, cookie: jar.header() };
  }

  const hidden = parseHiddenInputs(login.text);

  if (hidden.error_code === "505") {
    const newDeviceSessionId = await continueNewDeviceLogin(jar, hidden);

    if (newDeviceSessionId) {
      return { sessionId: newDeviceSessionId, cookie: jar.header() };
    }

    throw new Error("ECOUNT new device step could not be completed automatically.");
  }

  throw new Error(`ECOUNT login did not return a session. error_code=${hidden.error_code || "unknown"}`);
}

function buildSlipSearchPayload(pageCurrent = 1, dateRange = getRecentDateRange(), remarks = "") {
  return {
    Request: {
      Data: {
        FORM_GUBUN: "AR990",
        FORM_TYPE: "AR990",
        FORM_SEQ: "1",
        PAGE_CURRENT: pageCurrent,
        BASE_DATE_FROM: dateRange.from,
        BASE_DATE_TO: dateRange.to,
        PARAM: "",
        MENU_SEQ: "490",
        PRG_ID: "E060208",
        PAGE_SIZE,
        GB_TYPE: "Y",
        SLIP_TYPE: "1",
        TRADE_GUBUN: "Y",
        SEARCH_CRITERIA: ["CUST"],
        SEARCH_RANGE: "A",
        IsFromZaOnly: false,
        CANCEL: "",
        LIMIT_FLAG: 1,
        LIMIT_COUNT: 5001,
        DELETE_RANGE: "A",
        TabCd: "A1",
        SORT_TYPE: "",
        SEND_FLAG: "",
        P_FLAG: "",
        ALL_YN: "N",
        BASE_DATE_CALC: "",
        BASE_DATE_CALC_SELECT: "",
        BASE_DATE_SELECT: "recentDateMonth",
        CUST: "",
        CUST_DES: "",
        CUST_GROUP1: "",
        CUST_GROUP2: "",
        CUST_LEVEL_GROUP: "",
        CUST_LEVEL_GROUP_CHK: "1",
        GYE_CODE: "",
        SITE: "",
        SITE_LEVEL_GROUP: "",
        SITE_LEVEL_GROUP_CHK: "1",
        PJT_CD: "",
        PJT_CODE1: "",
        PJT_CODE2: "",
        ACCCASE: "",
        BILL_NO: "",
        AMT_F: "",
        AMT_T: "",
        REMARKS: remarks,
        DOC_NO: "",
        WRITER_ID: "",
        LAST_ID: "",
        BASE_DATE_CHK: "0"
      }
    },
    PAGE_CURRENT: pageCurrent,
    PAGE_SIZE
  };
}

function getRows(payload: unknown): EcountSlipRow[] {
  const data = getDataObject(payload);
  const rows = data.Data;
  return Array.isArray(rows) ? rows as EcountSlipRow[] : [];
}

function getTotalPages(payload: unknown, pageSize: number) {
  const data = getDataObject(payload);
  const totalCount = Number(data.TotalCount ?? data.TOTAL_COUNT ?? 0);

  if (!Number.isFinite(totalCount) || totalCount <= 0) {
    return 1;
  }

  return Math.max(1, Math.ceil(totalCount / pageSize));
}

function getTopLevelKeys(value: unknown) {
  return value && typeof value === "object" ? Object.keys(value as Record<string, unknown>).slice(0, 30) : [];
}

function textFromRow(row: EcountSlipRow) {
  return Object.values(row).map((value) => String(value ?? "")).join(" ");
}

function normalizeRate(value: unknown) {
  const rate = String(value ?? "").replace(/,/g, "").trim();
  return rate && rate !== "0" && rate !== "0.0000" ? rate : "";
}

function getExchangeRate(row: EcountSlipRow) {
  const directRate = normalizeRate(row["ACC101.EXCHANGE_RATE"]);

  if (directRate) {
    return directRate;
  }

  const explicit = textFromRow(row).match(/@\s*([\d,]+(?:\.\d+)?)/);
  return explicit ? explicit[1].replace(/,/g, "") : "";
}

function findExchangeRate(rows: EcountSlipRow[], poNo: string) {
  const row = rows.find((item) => textFromRow(item).includes(poNo) && Boolean(getExchangeRate(item)));

  if (!row) {
    return { row: null, exchangeRate: "" };
  }

  return {
    row,
    exchangeRate: getExchangeRate(row)
  };
}

async function scanSlipPages(jar: CookieJar, login: EcountLogin, poNo: string, mode: ScanMode) {
  const debug: ScanDebug = {
    mode: mode.label,
    dateRange: mode.dateRange,
    totalPages: 1,
    scannedPages: 0,
    rowsSeen: 0,
    poRowsSeen: 0,
    firstPageKeys: [],
    firstPageDataKeys: [],
    samplePoRowKeys: []
  };

  let totalPages = 1;

  for (let pageCurrent = 1; pageCurrent <= totalPages; pageCurrent += 1) {
    const response = await ecountFetch(
      jar,
      `${LOGIN_BASE}/ECAPI/SVC/Common/SlipTransfer/GetListXFormSlipAccountCommon?ec_req_sid=${encodeURIComponent(login.sessionId)}&xce=none`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "Referer": `${LOGIN_BASE}/ec5/view/erp?w_flag=1&ec_req_sid=${encodeURIComponent(login.sessionId)}`
        },
        body: JSON.stringify(buildSlipSearchPayload(pageCurrent, mode.dateRange, mode.remarks))
      }
    );
    const rows = getRows(response.data);
    const poRows = rows.filter((item) => textFromRow(item).includes(poNo));
    const found = findExchangeRate(rows, poNo);

    debug.scannedPages = pageCurrent;
    debug.rowsSeen += rows.length;
    debug.poRowsSeen += poRows.length;

    if (pageCurrent === 1) {
      totalPages = Math.min(getTotalPages(response.data, PAGE_SIZE), mode.maxPages);
      debug.totalPages = totalPages;
      debug.firstPageKeys = rows[0] ? Object.keys(rows[0]).slice(0, 20) : [];
      debug.firstPageDataKeys = getTopLevelKeys(getDataObject(response.data));
    }

    if (poRows.length && debug.samplePoRowKeys.length === 0) {
      debug.samplePoRowKeys = Object.keys(poRows[0]).slice(0, 20);
    }

    if (found.exchangeRate) {
      return {
        found: {
          poNo,
          exchangeRate: found.exchangeRate,
          pageCurrent,
          mode: mode.label
        },
        debug
      };
    }
  }

  return { found: null, debug };
}

async function lookupExchangeRate(poNo: string) {
  const login = await ecountWebLogin();
  const jar = new CookieJar();

  for (const cookie of login.cookie.split("; ")) {
    const separator = cookie.indexOf("=");
    if (separator > 0) {
      jar.set(cookie.slice(0, separator), cookie.slice(separator + 1));
    }
  }

  const recentDateRange = getRecentDateRange();
  const currentYearDateRange = getCurrentYearDateRange();
  const modes: ScanMode[] = [
    { label: "remarks-current-year", dateRange: currentYearDateRange, remarks: poNo, maxPages: 40 },
    { label: "recent-scan", dateRange: recentDateRange, remarks: "", maxPages: 80 },
    { label: "current-year-scan", dateRange: currentYearDateRange, remarks: "", maxPages: 160 }
  ];
  const debug: ScanDebug[] = [];

  for (const mode of modes) {
    const result = await scanSlipPages(jar, login, poNo, mode);
    debug.push(result.debug);

    if (result.found) {
      return result.found;
    }
  }

  throw new Error(`${poNo} exchange rate not found. debug=${JSON.stringify(debug)}`);
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") {
    return new Response(null, { headers: CORS_HEADERS });
  }

  if (req.method !== "POST") {
    return json({ ok: false, error: "POST requests only." }, 405);
  }

  const { poNo } = await req.json().catch(() => ({ poNo: "" }));
  const normalizedPoNo = typeof poNo === "string" ? poNo.trim() : "";

  if (!normalizedPoNo) {
    return json({ ok: false, error: "PO number is required." }, 400);
  }

  if (!COM_CODE || !USER_ID || !PASSWORD) {
    return json({ ok: false, error: "ECOUNT_COM_CODE, ECOUNT_USER_ID, and ECOUNT_PASSWORD secrets are required." }, 500);
  }

  try {
    const result = await lookupExchangeRate(normalizedPoNo);
    return json({ ok: true, ...result });
  } catch (error) {
    return json({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 502);
  }
});
