import { spawn } from "node:child_process";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { homedir } from "node:os";

const ROOT = process.cwd();

function loadDotEnv(text) {
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();

    if (!trimmed || trimmed.startsWith("#")) {
      continue;
    }

    const separator = trimmed.indexOf("=");

    if (separator <= 0) {
      continue;
    }

    const key = trimmed.slice(0, separator).trim();
    let value = trimmed.slice(separator + 1).trim();

    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }

    if (!process.env[key]) {
      process.env[key] = value;
    }
  }
}

async function loadLocalEnv() {
  try {
    loadDotEnv(await readFile(join(ROOT, ".env.local"), "utf8"));
  } catch {
    // Optional local-only file.
  }
}

function requireEnv(key) {
  const value = process.env[key];

  if (!value) {
    throw new Error(`${key} is required.`);
  }

  return value;
}

function runPowerShell(scriptPath, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.env.POWERSHELL || "powershell.exe", [
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

function parseJsonOutput(output) {
  const line = String(output || "")
    .split(/\r?\n/)
    .filter(Boolean)
    .reverse()
    .find((entry) => entry.trim().startsWith("{") && entry.trim().endsWith("}"));

  if (!line) {
    throw new Error("Password workbook reader did not return JSON.");
  }

  return JSON.parse(line);
}

await loadLocalEnv();

const supabaseUrl = requireEnv("SUPABASE_URL").replace(/\/$/, "");
const anonKey = requireEnv("SUPABASE_ANON_KEY");
const adminKey = requireEnv("PASSWORD_VAULT_ADMIN_KEY");
const excelPassword = requireEnv("PASSWORD_VAULT_EXCEL_PASSWORD");
const workbookPath = process.env.PASSWORD_VAULT_PATH || join(homedir(), "Desktop", "151515.xlsx");
const output = await runPowerShell(join(ROOT, "scripts", "password-vault-read.ps1"), [
  "-WorkbookPath", workbookPath,
  "-Password", excelPassword
]);
const vault = parseJsonOutput(output);
const response = await fetch(`${supabaseUrl}/functions/v1/password-vault`, {
  method: "POST",
  headers: {
    "Content-Type": "application/json",
    "Authorization": `Bearer ${anonKey}`,
    "apikey": anonKey
  },
  body: JSON.stringify({
    action: "save",
    adminKey,
    vault
  })
});
const data = await response.json().catch(() => ({}));

if (!response.ok || !data.ok) {
  throw new Error(data.error || `Upload failed with HTTP ${response.status}`);
}

const sheetCount = Array.isArray(vault.sheets) ? vault.sheets.length : 0;
const rowCount = Array.isArray(vault.sheets)
  ? vault.sheets.reduce((sum, sheet) => sum + (Array.isArray(sheet.rows) ? sheet.rows.length : 0), 0)
  : 0;

console.log(JSON.stringify({
  ok: true,
  file: vault.file,
  sheetCount,
  rowCount
}));
