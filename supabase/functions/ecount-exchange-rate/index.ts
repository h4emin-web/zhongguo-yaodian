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

type EcountLogin = {
  sessionId: string;
  cookie: string;
};

type EcountSlipRow = Record<string, unknown>;

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
      // Some ECOUNT view responses are base64 encoded HTML snippets, not JSON.
    }
  }

  try {
    return JSON.parse(text);
  } catch {
    return text;
  }
}

function toUtf8Bytes(value: string) {
  return [...new TextEncoder().encode(value)];
}

function encryptPassword(value: string, key: string) {
  return toUtf8Bytes(value)
    .map((byte, index) => {
      const hex = (byte ^ key.charCodeAt(index % key.length)).toString(16);
      return `${hex.length < 2 ? "0" : ""}${hex}`;
    })
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
    const getSetCookie = (response.headers as Headers & { getSetCookie?: () => string[] }).getSetCookie;
    const setCookies = getSetCookie ? getSetCookie.call(response.headers) : [];

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

  const text = await response.text();

  if (!response.ok && response.status >= 300) {
    throw new Error(`ECOUNT 요청 실패: ${response.status}`);
  }

  return {
    response,
    text,
    data: decodeMaybeBase64(text)
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
  const spasswd = buildEncryptedPassword(PASSWORD);
  const loginBody = new URLSearchParams({
    com_code: COM_CODE,
    id: USER_ID,
    spasswd,
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
    throw new Error("ECOUNT 새 기기 로그인 알림 단계에서 멈췄습니다. Supabase 서버 IP/기기를 ERP에서 한 번 등록해야 완전 자동 조회가 가능합니다.");
  }

  throw new Error(`ECOUNT 웹 로그인 세션을 만들지 못했습니다. error_code=${hidden.error_code || "unknown"}`);
}

function buildSlipSearchPayload(poNo: string, pageCurrent = 1) {
  return {
    Request: {
      Data: {
        FORM_GUBUN: "AR990",
        FORM_TYPE: "AR990",
        FORM_SEQ: "1",
        PAGE_CURRENT: pageCurrent,
        BASE_DATE_FROM: "20260101",
        BASE_DATE_TO: "20261231",
        PARAM: "",
        MENU_SEQ: "490",
        PRG_ID: "E060208",
        PAGE_SIZE: 25,
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
        BILL_NO: poNo,
        AMT_F: "",
        AMT_T: "",
        REMARKS: poNo,
        DOC_NO: "",
        WRITER_ID: "",
        LAST_ID: "",
        BASE_DATE_CHK: "0"
      }
    },
    PAGE_CURRENT: pageCurrent,
    PAGE_SIZE: 25
  };
}

function getRows(payload: unknown): EcountSlipRow[] {
  const data = getDataObject(payload);
  const rows = data.Data;
  return Array.isArray(rows) ? rows as EcountSlipRow[] : [];
}

function textFromRow(row: EcountSlipRow) {
  return Object.values(row).map((value) => String(value ?? "")).join(" ");
}

function findExchangeRate(rows: EcountSlipRow[], poNo: string) {
  const row = rows.find((item) => {
    const text = textFromRow(item);
    return text.includes(poNo) && text.includes("물품대") && text.includes("미지급");
  }) ?? rows.find((item) => {
    const text = textFromRow(item);
    return text.includes(poNo) && text.includes("물품대");
  });

  if (!row) {
    return { row: null, exchangeRate: "" };
  }

  const rate = String(row["ACC101.EXCHANGE_RATE"] ?? "").replace(/,/g, "");

  if (rate && rate !== "0.0000") {
    return { row, exchangeRate: rate };
  }

  const explicit = textFromRow(row).match(/@\s*([\d,]+(?:\.\d+)?)/);
  return {
    row,
    exchangeRate: explicit ? explicit[1].replace(/,/g, "") : ""
  };
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

  const response = await ecountFetch(
    jar,
    `${LOGIN_BASE}/ECAPI/SVC/Common/SlipTransfer/GetListXFormSlipAccountCommon?ec_req_sid=${encodeURIComponent(login.sessionId)}&xce=none`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Referer": `${LOGIN_BASE}/ec5/view/erp?w_flag=1&ec_req_sid=${encodeURIComponent(login.sessionId)}`
      },
      body: JSON.stringify(buildSlipSearchPayload(poNo))
    }
  );
  const rows = getRows(response.data);
  const found = findExchangeRate(rows, poNo);

  if (!found.exchangeRate) {
    throw new Error(`${poNo} 물품대 행에서 환율을 찾지 못했습니다.`);
  }

  return {
    poNo,
    exchangeRate: found.exchangeRate,
    row: found.row
  };
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

  if (!COM_CODE || !USER_ID || !PASSWORD) {
    return json({ ok: false, error: "ECOUNT_COM_CODE, ECOUNT_USER_ID, ECOUNT_PASSWORD secret이 필요합니다." }, 500);
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
