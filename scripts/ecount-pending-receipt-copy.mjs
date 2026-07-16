import { chromium } from "playwright";
import { readFile, writeFile } from "node:fs/promises";

const {
  ECOUNT_COM_CODE,
  ECOUNT_USER_ID,
  ECOUNT_PASSWORD,
  HEADLESS = "false",
  ECOUNT_PENDING_AUTO_SAVE = "0",
  KEEP_OPEN_MS = "0"
} = process.env;

const payloadPath = process.argv[2];
let currentStep = "init";
const dialogMessages = [];

function setStep(step) {
  currentStep = step;
  console.error(`[ecount-pending-receipt-copy] ${step}`);
}

function required(name, value) {
  if (!value) {
    throw new Error(`${name} is required`);
  }
}

function toNumber(value) {
  const number = Number(String(value || "").replace(/,/g, "").trim());
  return Number.isFinite(number) ? number : 0;
}

function normalizeDate(value) {
  const text = String(value || "").replace(/[./-]/g, "").trim();
  const match = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  return match ? `${match[1]}${match[2]}${match[3]}` : text;
}

function normalizeAmount(value) {
  return String(value || "").replace(/,/g, "").trim();
}

function parseSearchResponse(text) {
  try {
    return JSON.parse(text);
  } catch {}

  try {
    return JSON.parse(Buffer.from(text, "base64").toString("utf8"));
  } catch {
    return null;
  }
}

function sortPurchaseRows(rows) {
  return [...rows].sort((a, b) => {
    const aDate = String(a.IO_DATE || "").padEnd(8, "0");
    const bDate = String(b.IO_DATE || "").padEnd(8, "0");
    if (aDate !== bDate) {
      return bDate.localeCompare(aDate);
    }

    return toNumber(b.IO_NO) - toNumber(a.IO_NO);
  });
}

async function dismissDeviceRegistration(page) {
  const notNow = "\ub4f1\ub85d\uc548\ud568";
  await page.evaluate((label) => {
    const target = [...document.querySelectorAll("button, a, input")].find((element) => {
      const rect = element.getBoundingClientRect();
      const text = element.value || element.textContent || "";
      return rect.width > 0 && rect.height > 0 && text.trim() === label;
    });
    target?.click();
  }, notNow).catch(() => {});
  await page.waitForTimeout(2500);

  if (!page.url().includes("/ec5/view/erp")) {
    await page.mouse.click(548, 650).catch(() => {});
    await page.waitForTimeout(5000);
  }
}

async function login(page) {
  await page.goto("https://loginbb.ecount.com/ec5/view/app.login/erp_login?lan_type=ko-KR", {
    waitUntil: "domcontentloaded",
    timeout: 60_000
  });

  await page.fill("#com_code", ECOUNT_COM_CODE);
  await page.fill("#id", ECOUNT_USER_ID);
  await page.fill("#passwd", ECOUNT_PASSWORD);
  await page.locator("#passwd").press("Tab");
  await page.waitForTimeout(500);
  await page.click("#save");
  await page.waitForTimeout(6000);
  await dismissDeviceRegistration(page);

  if (!page.url().includes("/ec5/view/erp")) {
    throw new Error(`ECOUNT ERP screen did not open: ${page.url()}`);
  }
}

async function collectVisibleInputs(page) {
  const inputs = [];

  for (const frame of page.frames()) {
    const handles = await frame.locator("input").elementHandles().catch(() => []);

    for (let index = 0; index < handles.length; index += 1) {
      const handle = handles[index];
      const box = await handle.boundingBox().catch(() => null);

      if (!box || box.width <= 0 || box.height <= 0) {
        continue;
      }

      const meta = await handle.evaluate((input) => ({
        value: input.value || "",
        placeholder: input.getAttribute("placeholder") || "",
        title: input.getAttribute("title") || "",
        name: input.getAttribute("name") || "",
        id: input.id || "",
        type: input.getAttribute("type") || ""
      })).catch(() => null);

      if (!meta) {
        continue;
      }

      inputs.push({
        frame,
        handle,
        index,
        ...meta,
        x: box.x,
        y: box.y,
        width: box.width,
        height: box.height
      });
    }
  }

  return inputs;
}

async function collectVisibleControls(page) {
  const controls = [];

  for (const frame of page.frames()) {
    const handles = await frame.locator("button, a, input, [role='button'], .btn, .btn-primary, .btn-default").elementHandles().catch(() => []);

    for (let index = 0; index < handles.length; index += 1) {
      const handle = handles[index];
      const box = await handle.boundingBox().catch(() => null);

      if (!box || box.width <= 0 || box.height <= 0) {
        continue;
      }

      const meta = await handle.evaluate((element) => ({
        text: (element.innerText || element.textContent || element.value || "").trim(),
        title: element.getAttribute("title") || "",
        ariaLabel: element.getAttribute("aria-label") || "",
        id: element.id || "",
        className: typeof element.className === "string" ? element.className : ""
      })).catch(() => null);

      if (!meta) {
        continue;
      }

      controls.push({
        frame,
        handle,
        index,
        ...meta,
        x: box.x,
        y: box.y,
        width: box.width,
        height: box.height,
        centerY: box.y + box.height / 2
      });
    }
  }

  return controls;
}

async function fillHandle(input, value) {
  await input.handle.click();
  await input.frame.waitForTimeout(100).catch(() => {});
  await input.frame.page().keyboard.press("Control+A");
  if (value) {
    await input.frame.page().keyboard.type(String(value));
  } else {
    await input.frame.page().keyboard.press("Backspace");
  }
  await input.frame.page().keyboard.press("Tab");
  await input.frame.waitForTimeout(150).catch(() => {});
}

async function clickCopyButton(page) {
  const controls = await collectVisibleControls(page);
  const controlText = (control) => `${control.text} ${control.title} ${control.ariaLabel}`.replace(/\s+/g, "");
  const saveCandidates = controls
    .filter((control) => controlText(control).includes("저장"))
    .sort((a, b) => a.y - b.y || a.x - b.x);

  for (const save of saveCandidates) {
    const sameToolbarControls = controls
      .filter((control) => {
        const text = controlText(control);
        return control.x > save.x &&
          Math.abs(control.centerY - save.centerY) <= 36 &&
          !text.includes("저장") &&
          !text.includes("사원") &&
          !text.includes("검색");
      })
      .sort((a, b) => a.x - b.x);
    const copyControl = sameToolbarControls.find((control) => controlText(control).includes("복사"));

    if (copyControl) {
      await copyControl.handle.click();
      await page.waitForTimeout(1000);
      return `toolbar:${copyControl.text || copyControl.title || copyControl.ariaLabel || "copy"}`;
    }
  }

  await writeFile("tmp/ecount-pending-copy-controls.json", JSON.stringify(controls.map(({ frame, handle, ...control }) => control), null, 2), "utf8").catch(() => {});
  throw new Error("Could not locate the toolbar copy button next to Save.");
}

async function acceptCopyConfirmation(page, previousDialogCount) {
  const startedAt = Date.now();
  let clickedDomConfirm = false;

  while (Date.now() - startedAt < 12_000) {
    if (dialogMessages.length > previousDialogCount) {
      await page.waitForTimeout(1500);
      return {
        accepted: true,
        method: "native-dialog",
        message: dialogMessages[dialogMessages.length - 1] || ""
      };
    }

    for (const frame of page.frames()) {
      const frameText = await frame.locator("body").innerText({ timeout: 500 }).catch(() => "");
      const hasCopyPrompt =
        frameText.includes("복사하겠습니까") ||
        frameText.includes("시리얼") ||
        frameText.includes("로트No") ||
        frameText.includes("로트No.");

      if (!hasCopyPrompt) {
        continue;
      }

      const handles = await frame.locator("button, a, input, [role='button'], .btn, .btn-primary, .btn-default").elementHandles().catch(() => []);

      for (const handle of handles) {
        const box = await handle.boundingBox().catch(() => null);
        if (!box || box.width <= 0 || box.height <= 0) {
          continue;
        }

        const text = await handle.evaluate((element) => (
          element.innerText ||
          element.textContent ||
          element.value ||
          element.getAttribute("title") ||
          element.getAttribute("aria-label") ||
          ""
        ).trim()).catch(() => "");

        if (["확인", "예", "OK", "Ok", "ok"].includes(text) || text.includes("확인")) {
          await handle.click();
          clickedDomConfirm = true;
          await page.waitForTimeout(2500);
          return {
            accepted: true,
            method: "dom-confirm",
            message: frameText.replace(/\s+/g, " ").trim().slice(0, 160)
          };
        }
      }
    }

    await page.waitForTimeout(500);
  }

  return {
    accepted: clickedDomConfirm || dialogMessages.length > previousDialogCount,
    method: clickedDomConfirm ? "dom-confirm" : "",
    message: dialogMessages.slice(previousDialogCount).join(" / ")
  };
}

async function getVisibleProductInput(page, productCode) {
  const normalizedCode = String(productCode).trim().toUpperCase();
  let inputs = [];

  for (let attempt = 0; attempt < 8; attempt += 1) {
    inputs = await collectVisibleInputs(page);
    const productInput = inputs
      .filter((input) => {
        const value = String(input.value || "").trim().toUpperCase();
        return value === normalizedCode && input.y > 430;
      })
      .sort((a, b) => Math.abs(a.y - 690) - Math.abs(b.y - 690) || a.y - b.y)[0] || null;

    if (productInput) {
      return productInput;
    }

    await page.waitForTimeout(750);
  }

  await writeFile("tmp/ecount-pending-inputs.json", JSON.stringify(inputs.map(({ frame, handle, ...input }) => input), null, 2), "utf8").catch(() => {});
  throw new Error(`Could not locate product row input for ${productCode}.`);
}

async function fillTextAt(page, x, y, value) {
  await page.mouse.click(x, y);
  await page.waitForTimeout(200);
  await page.keyboard.press("Control+A");
  if (value) {
    await page.keyboard.type(String(value));
  } else {
    await page.keyboard.press("Backspace");
  }
  await page.keyboard.press("Tab");
  await page.waitForTimeout(250);
}

async function clearPendingFields(page, payload) {
  const yyyymmdd = normalizeDate(payload.instockDate);
  const quantity = normalizeAmount(payload.quantity);
  const inputs = await collectVisibleInputs(page);
  let dateFilled = false;
  let serialCleared = 0;
  let moneyCleared = 0;
  let quantityFilled = false;

  const dateCandidates = inputs
    .filter((input) => {
      const meta = `${input.placeholder} ${input.title} ${input.name} ${input.id}`.toLowerCase();
      const value = String(input.value || "").replace(/[./-]/g, "");
      return input.y < 260 && (meta.includes("date") || meta.includes("일자") || /^\d{8}$/.test(value));
    })
    .sort((a, b) => a.y - b.y || a.x - b.x);

  if (dateCandidates[0]) {
    await fillHandle(dateCandidates[0], yyyymmdd);
    dateFilled = true;
  }

  for (const input of inputs) {
    const meta = `${input.placeholder} ${input.title} ${input.name} ${input.id}`.toLowerCase();

    if (meta.includes("serial") || meta.includes("시리얼") || meta.includes("lot") || meta.includes("batch")) {
      await fillHandle(input, "");
      serialCleared += 1;
    }

    if (meta.includes("환율") || meta.includes("exchange") || meta.includes("단가") || meta.includes("금액") || meta.includes("공급") || meta.includes("부가")) {
      await fillHandle(input, "");
      moneyCleared += 1;
    }
  }

  const productInput = await getVisibleProductInput(page, payload.productCode);
  const rowY = productInput.y + productInput.height / 2;
  const quantityOffset = Number(process.env.ECOUNT_PENDING_QTY_OFFSET || "545");

  if (quantity) {
    await fillTextAt(page, productInput.x + quantityOffset, rowY, quantity);
    quantityFilled = true;
  }

  await fillTextAt(page, productInput.x + 615, rowY, "");
  await fillTextAt(page, productInput.x + 685, rowY, "");
  await fillTextAt(page, productInput.x + 755, rowY, "");

  return { dateFilled, serialCleared, moneyCleared, quantityFilled };
}

async function savePurchaseForm(page) {
  const saveText = "\uc800\uc7a5(F8)";

  const saveButton = page.getByText(saveText, { exact: true }).first();
  const clicked = await saveButton.click({ timeout: 5000 }).then(() => true).catch(() => false);

  if (!clicked) {
    await page.keyboard.press("F8").catch(() => {});
    await page.waitForTimeout(800);
  }

  await page.waitForTimeout(6000);
  return {};
}

async function main() {
  required("ECOUNT_COM_CODE", ECOUNT_COM_CODE);
  required("ECOUNT_USER_ID", ECOUNT_USER_ID);
  required("ECOUNT_PASSWORD", ECOUNT_PASSWORD);
  required("payload path", payloadPath);

  const payload = JSON.parse(await readFile(payloadPath, "utf8"));
  required("payload.productCode", payload.productCode);
  required("payload.instockDate", payload.instockDate);

  const browser = await chromium.launch({
    channel: "chrome",
    headless: HEADLESS === "true"
  });
  const page = await browser.newPage({ viewport: { width: 1420, height: 900 } });
  let searchRows = [];

  page.on("dialog", async (dialog) => {
    dialogMessages.push(dialog.message());
    await dialog.accept();
  });

  try {
    page.on("response", async (response) => {
      if (!response.url().includes("GetListXFormSlipCommon")) {
        return;
      }

      const requestData = response.request().postData() || "";
      if (!requestData.includes(payload.productCode)) {
        return;
      }

      const bodyText = await response.text().catch(() => "");
      const decoded = parseSearchResponse(bodyText);
      searchRows = decoded?.Data?.Data || [];
      console.error(`[ecount-pending-receipt-copy] captured ${searchRows.length} purchase rows`);
    });

    setStep("login");
    await login(page);
    setStep("open purchase search");
    await page.mouse.click(153, 424);
    await page.waitForTimeout(4500);

    setStep("search product");
    await page.locator("input").nth(1).click();
    await page.locator("input").nth(1).fill(payload.productCode);
    await page.locator("input").nth(1).press("Enter");
    await page.waitForTimeout(4500);

    setStep("select latest row");
    const quantity = toNumber(payload.quantity);
    const candidates = quantity > 0
      ? searchRows.filter((row) => Math.abs(toNumber(row.QTY || row["BUY.QTY"]) - quantity) < 0.0001)
      : searchRows;
    const targetRow = sortPurchaseRows(candidates)[0] || sortPurchaseRows(searchRows)[0];

    if (!targetRow) {
      throw new Error(`No purchase row found for ${payload.productCode}.`);
    }

    const purchaseNo = targetRow["BUY.IO_DATE_NO"];
    const visiblePurchaseNo = String(purchaseNo).replace(/^20(\d{2}\/)/, "$1");
    await page.getByText(visiblePurchaseNo, { exact: true }).first().dblclick();
    await page.waitForTimeout(7000);

    setStep("copy and clear form");
    const dialogCountBeforeCopy = dialogMessages.length;
    const copyButton = await clickCopyButton(page);
    const copyConfirm = await acceptCopyConfirmation(page, dialogCountBeforeCopy);
    const clearResult = await clearPendingFields(page, payload);
    await page.screenshot({ path: "tmp/ecount-pending-receipt-filled.png", fullPage: true }).catch(() => {});

    let saveResult = null;
    const autoSave = payload.autoSave === true || ECOUNT_PENDING_AUTO_SAVE === "1";

    if (autoSave) {
      setStep("save form");
      saveResult = await savePurchaseForm(page);
    }

    const result = {
      ok: true,
      saved: autoSave,
      purchaseNo,
      visiblePurchaseNo,
      productCode: payload.productCode,
      quantity: payload.quantity || "",
      instockDate: payload.instockDate,
      copyButton,
      copyConfirm,
      ...clearResult,
      dialogMessage: dialogMessages.join(" / "),
      dialogMessages,
      note: autoSave
        ? "Pending receipt purchase form was copied/cleared and saved."
        : "Pending receipt purchase form was copied/cleared. Save was not clicked."
    };

    await writeFile("tmp/ecount-pending-receipt-result.json", JSON.stringify(result, null, 2), "utf8");
    console.log(JSON.stringify(result));

    const keepOpenMs = Number(KEEP_OPEN_MS);
    if (keepOpenMs > 0) {
      await page.waitForTimeout(keepOpenMs);
    }
  } finally {
    await browser.close();
  }
}

main().catch(async (error) => {
  await writeFile("tmp/ecount-pending-receipt-error-dialogs.json", JSON.stringify(dialogMessages, null, 2), "utf8").catch(() => {});
  const result = {
    ok: false,
    step: currentStep,
    error: error instanceof Error ? error.message : String(error),
    dialogMessage: dialogMessages.join(" / "),
    dialogMessages
  };
  await writeFile("tmp/ecount-pending-receipt-result.json", JSON.stringify(result, null, 2), "utf8").catch(() => {});
  console.error(JSON.stringify(result));
  process.exit(1);
});
