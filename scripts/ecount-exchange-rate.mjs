import { chromium } from "playwright";
import { pathToFileURL } from "node:url";

const {
  ECOUNT_COM_CODE,
  ECOUNT_USER_ID,
  ECOUNT_PASSWORD,
  PO_NO,
  HEADLESS = "false"
} = process.env;

const IMPORT_EXPENSE_HASH =
  "#menuType=MENUTREE_000005&menuSeq=MENUTREE_000037&groupSeq=MENUTREE_000037&prgId=C000037&depth=2";

function requireEnv(name, value) {
  if (!value) {
    throw new Error(`${name} 환경변수가 필요합니다.`);
  }
}

function extractRate(line) {
  const explicit = line.match(/@\s*([\d,]+(?:\.\d+)?)/);

  if (explicit) {
    return explicit[1].replace(/,/g, "");
  }

  const parts = line.split("\t").map((part) => part.trim()).filter(Boolean);
  const lastNumber = [...parts].reverse().find((part) => /^[\d,]+(?:\.\d+)?$/.test(part));

  return lastNumber ? lastNumber.replace(/,/g, "") : "";
}

function findExchangeLine(text, poNo) {
  const rows = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.includes(poNo) && line.includes("물품대"));

  return rows.find((line) => line.includes("미지급")) || rows[0] || "";
}

async function fillAndLogin(page, credentials) {
  await page.goto("https://loginbb.ecount.com/ec5/view/app.login/erp_login?lan_type=ko-KR", {
    waitUntil: "domcontentloaded",
    timeout: 60_000
  });

  await page.fill("#com_code", credentials.comCode);
  await page.fill("#id", credentials.userId);
  await page.fill("#passwd", credentials.password);
  await page.locator("#passwd").press("Tab");
  await page.waitForTimeout(500);
  await page.click("#save");
  await page.waitForTimeout(7_000);

  await page.evaluate(() => {
    const button = Array.from(document.querySelectorAll("button")).find((element) => {
      const rect = element.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0 && element.textContent?.trim() === "등록안함";
    });

    button?.click();
  });
  await page.waitForTimeout(8_000);

  if (!page.url().includes("/ec5/view/erp")) {
    const text = await page.locator("body").innerText().catch(() => "");
    throw new Error(`ECOUNT 로그인 후 ERP 화면으로 이동하지 못했습니다. 현재 URL: ${page.url()} / 화면: ${text.slice(0, 300)}`);
  }
}

async function openImportExpense(page) {
  const erpUrl = page.url().split("#")[0];

  await page.goto(`${erpUrl}${IMPORT_EXPENSE_HASH}`, {
    waitUntil: "domcontentloaded",
    timeout: 60_000
  });
  await page.waitForTimeout(6_000);

  const text = await page.locator("body").innerText();

  if (!text.includes("수입비용전표검색")) {
    throw new Error("수입비용전표검색 화면을 열지 못했습니다.");
  }
}

async function goToPage(page, pageNumber) {
  const pageInput = page.locator("input.form-control").filter({
    hasNotText: ""
  });

  const candidates = await page.locator("input").evaluateAll((inputs) => {
    return inputs
      .map((input, index) => {
        const rect = input.getBoundingClientRect();
        return {
          index,
          value: input.value,
          visible: rect.width > 0 && rect.height > 0,
          width: rect.width,
          height: rect.height
        };
      })
      .filter((item) => item.visible && item.width <= 50 && item.height <= 32);
  });
  const target = candidates[0];

  if (target) {
    const inputs = page.locator("input");
    await inputs.nth(target.index).fill(String(pageNumber));
    await inputs.nth(target.index).press("Enter");
    await page.waitForTimeout(2_500);
    return;
  }

  await page.evaluate((pageNumber) => {
    const link = Array.from(document.querySelectorAll("a")).find((anchor) => {
      const rect = anchor.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0 && anchor.textContent?.trim() === String(pageNumber);
    });
    link?.click();
  }, pageNumber);
  await page.waitForTimeout(2_500);
}

async function findRate(page, poNo) {
  const initialText = await page.locator("body").innerText();
  const totalPages = Number(initialText.match(/\/\s*(\d+)/)?.[1] || "1");

  for (let pageNumber = 1; pageNumber <= totalPages; pageNumber += 1) {
    await goToPage(page, pageNumber);

    const gridText = await page.locator("#grid-main").innerText().catch(async () => {
      return page.locator("body").innerText();
    });
    const line = findExchangeLine(gridText, poNo);

    if (line) {
      const rate = extractRate(line);

      if (rate) {
        return { rate, line, pageNumber };
      }
    }
  }

  throw new Error(`${poNo} 물품대 행에서 환율을 찾지 못했습니다.`);
}

export async function lookupExchangeRate({
  poNo,
  comCode = ECOUNT_COM_CODE,
  userId = ECOUNT_USER_ID,
  password = ECOUNT_PASSWORD,
  headless = HEADLESS === "true"
}) {
  requireEnv("ECOUNT_COM_CODE", comCode);
  requireEnv("ECOUNT_USER_ID", userId);
  requireEnv("ECOUNT_PASSWORD", password);
  requireEnv("PO_NO", poNo);
  const browser = await chromium.launch({
    channel: "chrome",
    headless
  });

  try {
    const page = await browser.newPage({ viewport: { width: 1365, height: 900 } });

    await fillAndLogin(page, { comCode, userId, password });
    await openImportExpense(page);

    const result = await findRate(page, poNo);

    return {
      ok: true,
      poNo,
      exchangeRate: result.rate,
      pageNumber: result.pageNumber,
      line: result.line
    };
  } finally {
    await browser.close();
  }
}

async function main() {
  const result = await lookupExchangeRate({ poNo: PO_NO });
  console.log(JSON.stringify(result, null, 2));
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(JSON.stringify({
      ok: false,
      error: error instanceof Error ? error.message : String(error)
    }, null, 2));
    process.exit(1);
  });
}
