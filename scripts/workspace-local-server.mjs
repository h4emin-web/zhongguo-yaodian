import { createServer } from "node:http";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { spawn } from "node:child_process";
import { extname, join, normalize } from "node:path";
import { tmpdir } from "node:os";
import { lookupExchangeRate } from "./ecount-exchange-rate.mjs";

const PORT = Number(process.env.PORT || "4173");
const ROOT = process.cwd();
const POWERSHELL = process.env.POWERSHELL || "powershell.exe";

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
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
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

    const output = await runPowerShell(join(ROOT, "scripts", "auto-settlement-save.ps1"), [
      "-SettlementPath", settlementPath,
      "-ManagerName", fields.managerName || "",
      "-PoNo", fields.poNo || "",
      "-BoardingDate", fields.boardingDate || "",
      "-InstockDate", fields.instockDate || "",
      "-ExchangeRate", fields.exchangeRate || "",
      "-Quantity", fields.quantity || ""
    ]);
    const result = JSON.parse(output);
    json(res, { ok: true, ...result });
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
