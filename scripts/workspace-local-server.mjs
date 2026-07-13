import { createServer } from "node:http";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { spawn, spawnSync } from "node:child_process";
import { extname, join, normalize } from "node:path";
import { homedir, tmpdir } from "node:os";
import XLSX from "xlsx";
import { lookupExchangeRate } from "./ecount-exchange-rate.mjs";

const PORT = Number(process.env.PORT || "4173");
const ROOT = process.cwd();
const POWERSHELL = process.env.POWERSHELL || "powershell.exe";
const DEFAULT_OFFER_LIST_PATHS = [
  join(process.env.USERPROFILE || homedir(), "Desktop", "오퍼발행내역C8-2(양식수정).xlsm"),
  "Z:\\동진파마\\공유문서(Main)\\A60-오퍼리스트\\오퍼발행내역C8-2(양식수정).xlsm"
];
const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With",
  "Access-Control-Allow-Private-Network": "true"
};

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg"
};

function json(res, body, status = 200) {
  res.writeHead(status, { ...CORS_HEADERS, "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(body));
}

async function readBody(req) {
  const chunks = [];

  for await (const chunk of req) {
    chunks.push(chunk);
  }

  return JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
}

async function readBuffer(req) {
  const chunks = [];

  for await (const chunk of req) {
    chunks.push(chunk);
  }

  return Buffer.concat(chunks);
}

function parseMultipart(buffer, contentType) {
  const boundary = contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[1] || contentType.match(/boundary=(?:"([^"]+)"|([^;]+))/)?.[2];

  if (!boundary) {
    throw new Error("multipart boundary를 찾을 수 없습니다.");
  }

  const fields = {};
  const files = {};
  const body = buffer.toString("binary");
  const parts = body.split(`--${boundary}`);

  for (const part of parts) {
    if (!part || part === "--\r\n" || part === "--") {
      continue;
    }

    const cleanPart = part.startsWith("\r\n") ? part.slice(2) : part;
    const headerEnd = cleanPart.indexOf("\r\n\r\n");

    if (headerEnd < 0) {
      continue;
    }

    const rawHeaders = cleanPart.slice(0, headerEnd);
    let rawValue = cleanPart.slice(headerEnd + 4);

    if (rawValue.endsWith("\r\n")) {
      rawValue = rawValue.slice(0, -2);
    }

    const disposition = rawHeaders.match(/content-disposition:[^\r\n]+/i)?.[0] || "";
    const name = disposition.match(/name="([^"]+)"/)?.[1];
    const filename = disposition.match(/filename="([^"]*)"/)?.[1];

    if (!name) {
      continue;
    }

    if (filename) {
      files[name] = {
        filename,
        buffer: Buffer.from(rawValue, "binary")
      };
    } else {
      fields[name] = Buffer.from(rawValue, "binary").toString("utf8");
    }
  }

  return { fields, files };
}

function runPowerShell(scriptPath, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(POWERSHELL, [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      scriptPath,
      ...args
    ], {
      cwd: ROOT,
      windowsHide: true
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) {
        reject(new Error(stderr.trim() || stdout.trim() || `PowerShell exited with ${code}`));
        return;
      }

      resolve(stdout.trim());
    });
  });
}

function resolveOfferListPath() {
  const candidates = [
    process.env.OFFER_LIST_PATH,
    ...DEFAULT_OFFER_LIST_PATHS
  ].filter(Boolean);
  const offerListPath = candidates.find((candidate) => existsSync(candidate));

  if (!offerListPath) {
    throw new Error(`오퍼발행내역 파일을 찾지 못했습니다: ${candidates.join(" / ")}`);
  }

  return offerListPath;
}

function formatOfferDate(value, label) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);

    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, "0")}-${String(parsed.d).padStart(2, "0")}`;
    }
  }

  const text = String(value || "").trim();
  const match = text.replace(/[./]/g, "-").match(/(\d{4})-(\d{1,2})-(\d{1,2})/);

  if (match) {
    return `${match[1]}-${String(Number(match[2])).padStart(2, "0")}-${String(Number(match[3])).padStart(2, "0")}`;
  }

  throw new Error(`${label} 날짜를 해석하지 못했습니다: ${text || "(빈 값)"}`);
}

function lookupOfferDates(poNo) {
  try {
    const args = [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      join(ROOT, "scripts", "offer-list-lookup.ps1"),
      "-PoNo",
      poNo
    ];

    if (process.env.OFFER_LIST_PATH) {
      args.push("-WorkbookPath", process.env.OFFER_LIST_PATH);
    }

    const child = spawnSync(POWERSHELL, args, {
      cwd: ROOT,
      encoding: "utf8",
      windowsHide: true,
      timeout: 120000
    });

    if (child.status === 0 && child.stdout.trim()) {
      const output = child.stdout.trim().split(/\r?\n/).filter(Boolean).pop();
      return JSON.parse(output);
    }
  } catch {
  }

  const offerListPath = resolveOfferListPath();
  const workbook = XLSX.readFile(offerListPath, {
    cellDates: true,
    cellNF: false,
    cellText: false
  });
  const sheet = workbook.Sheets["PO메인"] || workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: true
  });

  for (let index = 2; index < rows.length; index += 1) {
    const row = rows[index];
    const candidate = String(row[0] || "").trim();

    if (candidate === poNo) {
      return {
        boardingDate: formatOfferDate(row[23], "Boarding"),
        instockDate: formatOfferDate(row[24], "Instock"),
        offerListPath,
        offerListRow: index + 1
      };
    }
  }

  throw new Error(`오퍼발행내역에서 PO '${poNo}' 를 찾지 못했습니다.`);
}

function runNode(scriptPath, args, env = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [
      scriptPath,
      ...args
    ], {
      cwd: ROOT,
      windowsHide: false,
      env: {
        ...process.env,
        ...env
      }
    });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });
    child.on("error", reject);
    child.on("close", (code) => {
      if (code !== 0) {
        reject(new Error(stderr.trim() || stdout.trim() || `Node exited with ${code}`));
        return;
      }

      resolve(stdout.trim());
    });
  });
}

function buildEcountPurchasePayload(result, autoSave = false) {
  return {
    productCode: result.productCode,
    quantity: result.quantity,
    exchangeRate: result.exchangeRate,
    unitPrice: result.purchaseUnitPrice || result.unitPrice,
    foreignAmount: result.foreignAmount,
    krwAmount: result.krwAmount,
    autoSave
  };
}

function buildInoutOrderPayload(result) {
  return {
    productCode: result.productCode,
    quantity: result.quantity,
    dueDate: result.instockDate,
    unitPrice: result.unitPrice,
    amount: result.krwAmount
  };
}

async function runEcountPurchaseFill(payload, options = {}) {
  requireEcountEnv();
  const required = ["productCode", "quantity", "exchangeRate", "unitPrice", "foreignAmount", "krwAmount"];
  const missing = required.filter((key) => payload[key] === undefined || payload[key] === null || payload[key] === "");

  if (missing.length) {
    throw new Error(`${missing.join(", ")} values are required.`);
  }

  const tempDir = await mkdtemp(join(tmpdir(), "haemin-ecount-purchase-"));

  try {
    const payloadPath = join(tempDir, "payload.json");
    await writeFile(payloadPath, JSON.stringify(payload), "utf8");
    const output = await runNode(join(ROOT, "scripts", "ecount-purchase-fill.mjs"), [payloadPath], {
      HEADLESS: "false",
      KEEP_OPEN_MS: options.keepOpenMs ?? "0",
      ECOUNT_PURCHASE_AUTO_SAVE: payload.autoSave ? "1" : "0"
    });

    return JSON.parse(output.split(/\r?\n/).filter(Boolean).pop() || "{}");
  } finally {
    await rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function runInoutOrderSave(payload, options = {}) {
  const required = ["productCode", "quantity", "dueDate", "unitPrice", "amount"];
  const missing = required.filter((key) => payload[key] === undefined || payload[key] === null || payload[key] === "");

  if (missing.length) {
    throw new Error(`${missing.join(", ")} values are required.`);
  }

  const args = [
    "-ProductCode", String(payload.productCode),
    "-Quantity", String(payload.quantity),
    "-DueDate", String(payload.dueDate),
    "-UnitPrice", String(payload.unitPrice),
    "-Amount", String(payload.amount)
  ];

  if (process.env.INOUT_ORDER_PATH) {
    args.push("-WorkbookPath", process.env.INOUT_ORDER_PATH);
  }

  if (options.commit) {
    args.push("-Commit");
  }

  const output = await runPowerShell(join(ROOT, "scripts", "inout-order-save.ps1"), args);
  return JSON.parse(output);
}

async function saveAutoSettlement(req, res) {
  let tempDir = "";

  try {
    const contentType = req.headers["content-type"] || "";
    const body = await readBuffer(req);
    const { fields, files } = parseMultipart(body, contentType);
    const file = files.settlementFile;

    if (!file) {
      throw new Error("Settlement file is required.");
    }

    tempDir = await mkdtemp(join(tmpdir(), "haemin-settlement-"));
    const extension = [".xlsx", ".xlsm", ".xls"].includes(extname(file.filename).toLowerCase())
      ? extname(file.filename).toLowerCase()
      : ".xlsx";
    const settlementPath = join(tempDir, `settlement-upload${extension}`);
    await writeFile(settlementPath, file.buffer);
    const offerDates = (!fields.boardingDate || !fields.instockDate)
      ? lookupOfferDates(fields.poNo || "")
      : null;
    const boardingDate = fields.boardingDate || offerDates?.boardingDate || "";
    const instockDate = fields.instockDate || offerDates?.instockDate || "";

    const output = await runPowerShell(join(ROOT, "scripts", "auto-settlement-save.ps1"), [
      "-SettlementPath", settlementPath,
      "-ManagerName", fields.managerName || "",
      "-PoNo", fields.poNo || "",
      "-BoardingDate", boardingDate,
      "-InstockDate", instockDate,
      "-ExchangeRate", fields.exchangeRate || "",
      "-Quantity", fields.quantity || ""
    ]);
    const result = JSON.parse(output);
    result.offerListPath = offerDates?.offerListPath || "";
    result.offerListRow = offerDates?.offerListRow || null;
    let ecount = null;

    if (process.env.AUTO_SETTLEMENT_ECOUNT_SAVE === "1") {
      try {
        ecount = await runEcountPurchaseFill(buildEcountPurchasePayload(result, true), {
          keepOpenMs: process.env.ECOUNT_PURCHASE_KEEP_OPEN_MS || "0"
        });
      } catch (error) {
        json(res, {
          ok: false,
          excelSaved: true,
          ...result,
          error: `수입정산서 원본은 저장됐지만 ERP 구매 저장에 실패했습니다: ${error instanceof Error ? error.message : String(error)}`
        }, 500);
        return;
      }
    }

    let inout = null;

    if (process.env.AUTO_SETTLEMENT_INOUT_SAVE === "1") {
      try {
        inout = await runInoutOrderSave(buildInoutOrderPayload(result), {
          commit: true
        });
      } catch (error) {
        json(res, {
          ok: false,
          excelSaved: true,
          ecountSaved: Boolean(ecount?.saved),
          ...result,
          ecount,
          error: `수입정산서 원본 저장${ecount?.saved ? "과 ERP 구매 저장" : ""}은 완료됐지만 입출고 지시서 저장에 실패했습니다: ${error instanceof Error ? error.message : String(error)}`
        }, 500);
        return;
      }
    }

    json(res, { ok: true, ...result, ecount, inout });
  } catch (error) {
    json(res, {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 500);
  } finally {
    if (tempDir) {
      await rm(tempDir, { recursive: true, force: true }).catch(() => {});
    }
  }
}

function requireEcountEnv() {
  const missing = ["ECOUNT_COM_CODE", "ECOUNT_USER_ID", "ECOUNT_PASSWORD"].filter((key) => !process.env[key]);

  if (missing.length) {
    throw new Error(`${missing.join(", ")} environment variables are required.`);
  }
}

async function startEcountPurchaseFill(req, res) {
  try {
    const payload = await readBody(req);
    const result = await runEcountPurchaseFill({
      ...payload,
      autoSave: payload.autoSave === true || process.env.ECOUNT_PURCHASE_AUTO_SAVE === "1"
    }, {
      keepOpenMs: process.env.ECOUNT_PURCHASE_KEEP_OPEN_MS || "0"
    });

    json(res, {
      ok: true,
      ...result
    });
  } catch (error) {
    json(res, {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 500);
  }
}

async function serveFile(req, res) {
  const url = new URL(req.url, `http://localhost:${PORT}`);
  const pathname = decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname);
  const filePath = normalize(join(ROOT, pathname));

  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  try {
    const data = await readFile(filePath);
    res.writeHead(200, { "Content-Type": TYPES[extname(filePath)] || "application/octet-stream" });
    res.end(data);
  } catch {
    res.writeHead(404);
    res.end("Not found");
  }
}

const server = createServer(async (req, res) => {
  if (req.method === "OPTIONS") {
    res.writeHead(204, CORS_HEADERS);
    res.end();
    return;
  }

  if (req.method === "POST" && req.url === "/api/ecount-rate") {
    try {
      const { poNo } = await readBody(req);
      const result = await lookupExchangeRate({ poNo, headless: process.env.HEADLESS === "true" });

      json(res, result);
    } catch (error) {
      json(res, {
        ok: false,
        error: error instanceof Error ? error.message : String(error)
      }, 500);
    }
    return;
  }

  if (req.method === "POST" && req.url === "/api/auto-settlement/save") {
    await saveAutoSettlement(req, res);
    return;
  }

  if (req.method === "POST" && req.url === "/api/ecount-purchase/fill") {
    await startEcountPurchaseFill(req, res);
    return;
  }

  if (req.method !== "GET") {
    res.writeHead(405);
    res.end("Method not allowed");
    return;
  }

  await serveFile(req, res);
});

server.listen(PORT, () => {
  console.log(`WorkSpace local server: http://localhost:${PORT}`);
});
