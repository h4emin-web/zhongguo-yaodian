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

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
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

function normalizePoNo(value) {
  return String(value || "")
    .normalize("NFKC")
    .replace(/[\u2010-\u2015\u2212]/g, "-")
    .replace(/\s+/g, "")
    .trim()
    .toUpperCase();
}

function lookupOfferDates(poNo) {
  let powerShellLookupError = "";
  const normalizedPoNo = normalizePoNo(poNo);

  try {
    const args = [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      join(ROOT, "scripts", "offer-list-lookup.ps1"),
      "-PoNo",
      normalizedPoNo
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

    if (child.stdout.trim()) {
      const output = child.stdout.trim()
        .split(/\r?\n/)
        .filter(Boolean)
        .reverse()
        .find((line) => line.trim().startsWith("{") && line.trim().endsWith("}"));

      if (output) {
        return JSON.parse(output);
      }
    }

    powerShellLookupError = (child.stderr || child.stdout || "").trim();
  } catch {
    powerShellLookupError = "PowerShell 오퍼 조회 실행 중 예외가 발생했습니다.";
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
    const candidate = [row[0], row[1], row[2]]
      .map((value) => String(value || "").trim().toUpperCase())
      .find((value) => normalizePoNo(value) === normalizedPoNo);

    if (candidate) {
      return {
        boardingDate: formatOfferDate(row[23], "Boarding"),
        instockDate: formatOfferDate(row[24], "Instock"),
        offerListPath,
        offerListRow: index + 1
      };
    }
  }

  throw new Error(`오퍼발행내역에서 PO '${poNo}' 를 찾지 못했습니다.${powerShellLookupError ? ` PowerShell 조회 오류: ${powerShellLookupError}` : ""}`);
}

function runNode(scriptPath, args, env = {}, options = {}) {
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
    let settled = false;
    const timeoutMs = Number(options.timeoutMs || 0);
    const timer = timeoutMs > 0
      ? setTimeout(() => {
        if (settled) {
          return;
        }

        settled = true;
        spawn("taskkill", ["/PID", String(child.pid), "/T", "/F"], { windowsHide: true }).on("error", () => {});
        reject(new Error(`Node automation timed out after ${timeoutMs} ms. ${stderr.trim() || stdout.trim()}`.trim()));
      }, timeoutMs)
      : null;

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString("utf8");
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString("utf8");
    });
    child.on("error", (error) => {
      if (settled) {
        return;
      }

      settled = true;
      if (timer) {
        clearTimeout(timer);
      }
      reject(error);
    });
    child.on("close", (code) => {
      if (settled) {
        return;
      }

      settled = true;
      if (timer) {
        clearTimeout(timer);
      }

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
    unitPrice: result.erpUnitPrice ?? result.unitPrice,
    foreignAmount: result.erpForeignAmount ?? result.foreignAmount,
    krwAmount: result.erpKrwAmount ?? result.krwAmount,
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
    }, {
      timeoutMs: Number(process.env.ECOUNT_PURCHASE_TIMEOUT_MS || "180000")
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

function parseJsonOutput(output) {
  const line = String(output || "")
    .split(/\r?\n/)
    .filter(Boolean)
    .reverse()
    .find((entry) => entry.trim().startsWith("{") && entry.trim().endsWith("}"));

  if (!line) {
    throw new Error(`Automation did not return JSON. ${String(output || "").slice(-1000)}`.trim());
  }

  return JSON.parse(line);
}

function pendingReceiptShouldCommit() {
  return process.env.PENDING_RECEIPT_COMMIT !== "0";
}

async function runPendingOfferUpdate(item, instockDate) {
  const args = [
    "-PoNo", String(item.poNo || ""),
    "-InstockDate", String(instockDate || "")
  ];

  if (process.env.OFFER_LIST_PATH) {
    args.push("-WorkbookPath", process.env.OFFER_LIST_PATH);
  }

  if (pendingReceiptShouldCommit()) {
    args.push("-Commit");
  }

  const output = await runPowerShell(join(ROOT, "scripts", "pending-receipt-offer-update.ps1"), args);
  return parseJsonOutput(output);
}

async function runPendingInoutCopy(offer) {
  const args = [
    "-ProductCode", String(offer.productCode || ""),
    "-Quantity", String(offer.quantity || ""),
    "-PoDate", String(offer.poDate || ""),
    "-InstockDate", String(offer.instockDate || "")
  ];

  if (process.env.INOUT_ORDER_PATH) {
    args.push("-WorkbookPath", process.env.INOUT_ORDER_PATH);
  }

  if (pendingReceiptShouldCommit()) {
    args.push("-Commit");
  }

  const output = await runPowerShell(join(ROOT, "scripts", "pending-receipt-inout-copy.ps1"), args);
  return parseJsonOutput(output);
}

async function runPendingCoaCopy(offer) {
  const quantityLabel = `${offer.quantity || ""}${offer.unit || ""}`.trim();
  const args = [
    "-PoNo", String(offer.poNo || ""),
    "-ProductName", String(offer.productName || ""),
    "-InstockDate", String(offer.instockDate || ""),
    "-Quantity", quantityLabel,
    "-Messrs", String(offer.messrs || "")
  ];

  if (process.env.PENDING_RECEIPT_COA_SOURCE_ROOT) {
    args.push("-SourceRoot", process.env.PENDING_RECEIPT_COA_SOURCE_ROOT);
  }

  if (process.env.PENDING_RECEIPT_COA_TARGET_ROOT) {
    args.push("-TargetRoot", process.env.PENDING_RECEIPT_COA_TARGET_ROOT);
  }

  if (pendingReceiptShouldCommit()) {
    args.push("-Commit");
  }

  const output = await runPowerShell(join(ROOT, "scripts", "pending-receipt-coa-copy.ps1"), args);
  return parseJsonOutput(output);
}

async function runPendingEcountCopy(offer) {
  requireEcountEnv();
  const tempDir = await mkdtemp(join(tmpdir(), "haemin-pending-ecount-"));

  try {
    const payloadPath = join(tempDir, "payload.json");
    const payload = {
      productCode: offer.productCode,
      quantity: offer.quantity,
      instockDate: offer.instockDate,
      autoSave: process.env.ECOUNT_PENDING_AUTO_SAVE === "1" || process.env.ECOUNT_PURCHASE_AUTO_SAVE === "1"
    };
    await writeFile(payloadPath, JSON.stringify(payload), "utf8");
    const output = await runNode(join(ROOT, "scripts", "ecount-pending-receipt-copy.mjs"), [payloadPath], {
      HEADLESS: "false",
      KEEP_OPEN_MS: process.env.ECOUNT_PENDING_KEEP_OPEN_MS || "0",
      ECOUNT_PENDING_AUTO_SAVE: payload.autoSave ? "1" : "0"
    }, {
      timeoutMs: Number(process.env.ECOUNT_PENDING_TIMEOUT_MS || "180000")
    });

    return parseJsonOutput(output);
  } finally {
    await rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

async function processPendingReceipt(req, res) {
  try {
    const payload = await readBody(req);
    const instockDate = String(payload.instockDate || "").trim();
    const items = Array.isArray(payload.items) ? payload.items : [];

    if (!instockDate) {
      throw new Error("Instock date is required.");
    }

    if (!items.length) {
      throw new Error("At least one PO number is required.");
    }

    const shouldRunInout = process.env.PENDING_RECEIPT_INOUT_SAVE !== "0";
    const shouldRunEcount = process.env.PENDING_RECEIPT_ECOUNT_SAVE === "1" || (
      process.env.PENDING_RECEIPT_ECOUNT_SAVE !== "0" && process.env.AUTO_SETTLEMENT_ECOUNT_SAVE === "1"
    );
    const shouldRunCoa = process.env.PENDING_RECEIPT_COA_COPY !== "0";
    const results = [];

    for (const item of items) {
      const poNo = String(item.poNo || "").trim();
      if (!poNo) {
        throw new Error("PO number is empty.");
      }

      const itemResult = {
        poNo,
        ok: false,
        offer: null,
        inout: null,
        ecount: null,
        coa: null,
        warnings: []
      };

      try {
        const offer = await runPendingOfferUpdate({ poNo }, instockDate);
        itemResult.offer = offer;
        Object.assign(itemResult, {
          poNo: offer.poNo || poNo,
          productCode: offer.productCode,
          productName: offer.productName,
          quantity: offer.quantity,
          unit: offer.unit,
          poDate: offer.poDate,
          instockDate: offer.instockDate,
          messrs: offer.messrs
        });

        if (shouldRunInout) {
          itemResult.inout = await runPendingInoutCopy(offer);
          if (itemResult.inout?.warning) {
            itemResult.warnings.push(itemResult.inout.warning);
          }
        }

        if (shouldRunEcount) {
          itemResult.ecount = await runPendingEcountCopy(offer);
        }

        if (shouldRunCoa) {
          itemResult.coa = await runPendingCoaCopy(offer);
        }

        itemResult.ok = true;
      } catch (error) {
        itemResult.error = error instanceof Error ? error.message : String(error);
      }

      results.push(itemResult);
    }

    const failed = results.filter((item) => !item.ok);

    json(res, {
      ok: failed.length === 0,
      committed: pendingReceiptShouldCommit(),
      mode: payload.mode === "batch" ? "batch" : "single",
      instockDate,
      count: results.length,
      items: results,
      error: failed.length ? failed.map((item) => `${item.poNo}: ${item.error}`).join(" / ") : ""
    }, failed.length ? 500 : 200);
  } catch (error) {
    json(res, {
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, 500);
  }
}

function parseBatchItems(value) {
  if (!value) {
    return [];
  }

  const parsed = JSON.parse(value);

  if (!Array.isArray(parsed)) {
    throw new Error("batchItems must be an array.");
  }

  return parsed.map((item) => ({
    poNo: String(item.poNo || "").trim(),
    quantity: String(item.quantity || "").trim(),
    vat: String(item.vat || "").trim(),
    duty: String(item.duty || "").trim(),
    ratioBasis: String(item.ratioBasis || "").trim()
  })).filter((item) => item.poNo);
}

function resolveBatchRatioBasis(fields, batchItems) {
  if (fields.batchRatioBasis === "quantity") {
    return "quantity";
  }

  const hasMissingVat = batchItems.some((item) => !item.vat);
  const allHaveQuantity = batchItems.length > 0 && batchItems.every((item) => item.quantity);

  return hasMissingVat && allHaveQuantity ? "quantity" : "tax";
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
    const settlementMode = fields.settlementMode === "batch" ? "batch" : "single";
    const batchItems = settlementMode === "batch" ? parseBatchItems(fields.batchItems) : [];
    const batchRatioBasis = settlementMode === "batch" ? resolveBatchRatioBasis(fields, batchItems) : "tax";
    const batchItemsWithDates = settlementMode === "batch"
      ? batchItems.map((item) => {
        const dates = lookupOfferDates(item.poNo);
        return {
          ...item,
          ratioBasis: batchRatioBasis,
          boardingDate: dates.boardingDate,
          instockDate: dates.instockDate
        };
      })
      : [];
    const offerDates = settlementMode === "single" && (!fields.boardingDate || !fields.instockDate)
      ? lookupOfferDates(fields.poNo || "")
      : null;
    const boardingDate = fields.boardingDate || offerDates?.boardingDate || "";
    const instockDate = fields.instockDate || offerDates?.instockDate || "";

    const powerShellArgs = [
      "-SettlementPath", settlementPath,
      "-ManagerName", fields.managerName || "",
      "-PoNo", fields.poNo || (settlementMode === "batch" ? "__BATCH__" : ""),
      "-BoardingDate", boardingDate,
      "-InstockDate", instockDate,
      "-ExchangeRate", fields.exchangeRate || "",
      "-Quantity", fields.quantity || ""
    ];

    if (settlementMode === "batch") {
      const batchItemsPath = join(tempDir, "batch-items.json");
      await writeFile(batchItemsPath, JSON.stringify(batchItemsWithDates), "utf8");
      powerShellArgs.push("-BatchItemsPath", batchItemsPath, "-BatchRatioBasis", batchRatioBasis);
    }

    const output = await runPowerShell(join(ROOT, "scripts", "auto-settlement-save.ps1"), powerShellArgs);
    const result = JSON.parse(output);
    result.offerListPath = offerDates?.offerListPath || "";
    result.offerListRow = offerDates?.offerListRow || null;
    let ecount = null;
    let ecountError = "";
    let inout = null;
    let inoutError = "";
    const resultItems = Array.isArray(result.items) ? result.items : [result];

    if (process.env.AUTO_SETTLEMENT_INOUT_SAVE === "1") {
      try {
        await delay(Number(process.env.AUTO_SETTLEMENT_EXCEL_SETTLE_MS || "3000"));
        inout = [];

        for (const item of resultItems) {
          inout.push(await runInoutOrderSave(buildInoutOrderPayload(item), {
            commit: true
          }));
        }

        if (inout.length === 1 && !Array.isArray(result.items)) {
          inout = inout[0];
        }
      } catch (error) {
        inoutError = error instanceof Error ? error.message : String(error);
      }
    }

    if (process.env.AUTO_SETTLEMENT_ECOUNT_SAVE === "1") {
      try {
        ecount = [];

        for (const item of resultItems) {
          ecount.push(await runEcountPurchaseFill(buildEcountPurchasePayload(item, true), {
            keepOpenMs: process.env.ECOUNT_PURCHASE_KEEP_OPEN_MS || "0"
          }));
        }

        if (ecount.length === 1 && !Array.isArray(result.items)) {
          ecount = ecount[0];
        }
      } catch (error) {
        ecountError = error instanceof Error ? error.message : String(error);
      }
    }

    if (inoutError) {
      json(res, {
        ok: false,
        excelSaved: true,
        ecountSaved: Boolean(ecount?.saved),
        ...result,
        ecount,
        ecountError,
        error: `수입정산서 원본 저장${ecount?.saved ? "과 ERP 구매 저장" : ""}은 완료됐지만 입출고 지시서 저장에 실패했습니다: ${inoutError}`
      }, 500);
      return;
    }

    json(res, {
      ok: true,
      ...result,
      ecount,
      inout,
      warning: ecountError ? `ERP 구매 저장에 실패했습니다: ${ecountError}` : ""
    });
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

  if (req.method === "POST" && req.url === "/api/pending-receipt/process") {
    await processPendingReceipt(req, res);
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
