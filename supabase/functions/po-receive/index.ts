// PO 입고처리: PO 번호를 받아 shipping MariaDB의 상태를 "입고"로 바꾸고,
// 이카운트 ERP 연결(Zone 조회 + 로그인)까지 확인한다.
// 이카운트 쪽 실제 저장 호출은 정확한 화면/엔드포인트가 정해지면 채워 넣는다.

import "@supabase/functions-js/edge-runtime.d.ts";
import { withSupabase } from "@supabase/server";
import mysql from "npm:mysql2@3/promise";

const SHIPPING_DB_HOST = Deno.env.get("SHIPPING_DB_HOST")!;
const SHIPPING_DB_PORT = Number(Deno.env.get("SHIPPING_DB_PORT") ?? "3306");
const SHIPPING_DB_NAME = Deno.env.get("SHIPPING_DB_NAME")!;
const SHIPPING_DB_USER = Deno.env.get("SHIPPING_DB_USER")!;
const SHIPPING_DB_PASSWORD = Deno.env.get("SHIPPING_DB_PASSWORD")!;

const ECOUNT_COM_CODE = Deno.env.get("ECOUNT_COM_CODE")!;
const ECOUNT_USER_ID = Deno.env.get("ECOUNT_USER_ID")!;
const ECOUNT_API_CERT_KEY = Deno.env.get("ECOUNT_API_CERT_KEY")!;

function todayKST(): string {
  const now = new Date(Date.now() + 9 * 60 * 60 * 1000);
  return now.toISOString().slice(0, 10);
}

async function markReceivedInShippingDb(poNo: string, receivedDate: string) {
  const conn = await mysql.createConnection({
    host: SHIPPING_DB_HOST,
    port: SHIPPING_DB_PORT,
    database: SHIPPING_DB_NAME,
    user: SHIPPING_DB_USER,
    password: SHIPPING_DB_PASSWORD,
  });

  try {
    const [rows] = await conn.execute(
      "SELECT PONo FROM offer_main2 WHERE PONo = ?",
      [poNo],
    );
    if (!Array.isArray(rows) || rows.length === 0) {
      return { found: false as const };
    }

    await conn.execute(
      "UPDATE offer_main2 SET 진행상태 = '입고', Instock = ? WHERE PONo = ?",
      [receivedDate, poNo],
    );

    return { found: true as const };
  } finally {
    await conn.end();
  }
}

async function ecountLogin() {
  const zoneRes = await fetch(
    `https://sboapi.ecount.com/OAPI/V2/Zone?COM_CODE=${encodeURIComponent(ECOUNT_COM_CODE)}`,
  );
  const zoneJson = await zoneRes.json();
  const zone = zoneJson?.Data?.ZONE;
  if (!zone) {
    return { loggedIn: false, error: `ZONE 조회 실패: ${JSON.stringify(zoneJson)}` };
  }

  const loginRes = await fetch(
    `https://sboapi${zone}.ecount.com/OAPI/V2/OAPILogin`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        COM_CODE: ECOUNT_COM_CODE,
        USER_ID: ECOUNT_USER_ID,
        API_CERT_KEY: ECOUNT_API_CERT_KEY,
        LAN_TYPE: "ko-KR",
        ZONE: zone,
      }),
    },
  );
  const loginJson = await loginRes.json();
  const sessionId = loginJson?.Data?.Datas?.SESSION_ID;
  if (!sessionId) {
    return { loggedIn: false, error: `로그인 실패: ${JSON.stringify(loginJson)}` };
  }

  // TODO: ECOUNT_SAVE_ENDPOINT 확정되면 여기서 구매 화면 입고일자 저장 호출 추가
  // (zone, sessionId 를 사용해 이카운트 구매 저장 API 호출)

  return { loggedIn: true, zone, pending: "구매 화면 입고일자 저장 API는 아직 미확정" };
}

export default {
  fetch: withSupabase({ auth: ["publishable", "secret"] }, async (req) => {
    if (req.method !== "POST") {
      return Response.json({ ok: false, error: "POST 요청만 지원합니다." }, { status: 405 });
    }

    let poNo: string | undefined;
    let receivedDate: string | undefined;
    try {
      const body = await req.json();
      poNo = typeof body.poNo === "string" ? body.poNo.trim() : undefined;
      receivedDate = typeof body.receivedDate === "string" ? body.receivedDate : undefined;
    } catch {
      return Response.json({ ok: false, error: "요청 본문이 올바르지 않습니다." }, { status: 400 });
    }

    if (!poNo) {
      return Response.json({ ok: false, error: "PO 번호를 입력하세요." }, { status: 400 });
    }

    const date = receivedDate ?? todayKST();

    let mariadbResult;
    try {
      mariadbResult = await markReceivedInShippingDb(poNo, date);
    } catch (err) {
      return Response.json(
        { ok: false, error: `shipping DB 연결/처리 실패: ${err instanceof Error ? err.message : String(err)}` },
        { status: 502 },
      );
    }

    if (!mariadbResult.found) {
      return Response.json(
        { ok: false, error: `PO 번호 '${poNo}'를 shipping DB에서 찾을 수 없습니다.` },
        { status: 404 },
      );
    }

    let ecountResult;
    try {
      ecountResult = await ecountLogin();
    } catch (err) {
      ecountResult = { loggedIn: false, error: err instanceof Error ? err.message : String(err) };
    }

    return Response.json({
      ok: true,
      poNo,
      receivedDate: date,
      mariadb: "updated",
      ecount: ecountResult,
    });
  }),
};

/* 로컬 테스트:

  1. supabase start
  2. supabase secrets set --env-file ../../.env  (또는 개별 env 지정)
  3. curl -i --location --request POST 'http://127.0.0.1:54321/functions/v1/po-receive' \
       --header 'apiKey: <anon key>' \
       --header 'Content-Type: application/json' \
       --data '{"poNo":"Z26-00013"}'

*/
