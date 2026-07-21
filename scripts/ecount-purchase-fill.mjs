import { chromium } from "playwright";
import { readFile, writeFile } from "node:fs/promises";

const {
  ECOUNT_COM_CODE,
  ECOUNT_USER_ID,
  ECOUNT_PASSWORD,
  HEADLESS = "false",
  ECOUNT_PURCHASE_AUTO_SAVE = "0",
  KEEP_OPEN_MS = "1800000"
} = process.env;

const payloadPath = process.argv[2];
let currentStep = "init";
const PURCHASE_TRANSACTION_TYPE_LABEL = "\uac70\ub798\uc720\ud615";
const PURCHASE_TRANSACTION_TYPE_VALUE = "\ubd80\uac00\uc138\uc728 \ubbf8\uc801\uc6a9";

function setStep(step) {
  currentStep = step;
  console.error(`[ecount-purchase-fill] ${step}`);
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

function normalizeAmount(value) {
  return String(Math.round(toNumber(value)));
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
    const handles = await frame.locator("input, select, textarea, [role='combobox']").elementHandles().catch(() => []);

    for (let index = 0; index < handles.length; index += 1) {
      const handle = handles[index];
      const box = await handle.boundingBox().catch(() => null);

      if (!box || box.width <= 0 || box.height <= 0) {
        continue;
      }

      const meta = await handle.evaluate((input) => ({
        tagName: input.tagName || "",
        value: input.value || "",
        selectedText: input.tagName === "SELECT"
          ? (input.options?.[input.selectedIndex]?.textContent || "").trim()
          : "",
        text: (input.innerText || input.textContent || "").trim(),
        placeholder: input.getAttribute("placeholder") || "",
        title: input.getAttribute("title") || "",
        name: input.getAttribute("name") || "",
        id: input.id || "",
        ariaLabel: input.getAttribute("aria-label") || "",
        type: input.getAttribute("type") || ""
      })).catch(() => null);

      if (!meta) {
        continue;
      }

      inputs.push({
        frameUrl: frame.url(),
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

async function fillHandle(input, value) {
  await input.handle.evaluate((element, nextValue) => {
    const tagName = String(element.tagName || "").toUpperCase();
    if (tagName === "SELECT") {
      const desired = String(nextValue ?? "").trim();
      const option = [...element.options].find((entry) => {
        const optionValue = String(entry.value || "").trim();
        const optionText = String(entry.textContent || "").trim();
        return optionValue === desired || optionText === desired || optionText.includes(desired);
      });
      if (!option) {
        throw new Error(`select option was not found: ${desired}`);
      }
      element.value = option.value;
    } else {
      element.focus?.();
      element.value = String(nextValue ?? "");
    }

    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
    element.blur?.();
  }, value);
  await input.frame.waitForTimeout(200).catch(() => {});
}

async function readHandleText(handle) {
  return await handle.evaluate((element) => {
    const tagName = String(element.tagName || "").toUpperCase();
    if (tagName === "SELECT") {
      return (element.options?.[element.selectedIndex]?.textContent || "").trim();
    }

    return String(element.value || element.innerText || element.textContent || "").trim();
  }).catch(() => "");
}

async function findControlRightOfLabel(page, labelText) {
  for (const frame of page.frames()) {
    const handle = await frame.evaluateHandle((label) => {
      const normalize = (value) => String(value || "").replace(/\s+/g, "");
      const targetLabel = normalize(label);
      const visible = (element) => {
        const rect = element.getBoundingClientRect();
        const style = window.getComputedStyle(element);
        return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
      };
      const textOf = (element) => normalize(element.innerText || element.textContent || element.value || "");
      const labelElements = [...document.querySelectorAll("label, span, div, td, th, p")]
        .filter((element) => {
          if (!visible(element)) {
            return false;
          }

          const rect = element.getBoundingClientRect();
          const text = textOf(element);
          return text.includes(targetLabel) && rect.width <= 170 && rect.height <= 42;
        })
        .sort((a, b) => {
          const aRect = a.getBoundingClientRect();
          const bRect = b.getBoundingClientRect();
          return aRect.top - bRect.top || aRect.left - bRect.left;
        });

      const controls = [...document.querySelectorAll("input, select, textarea, [role='combobox']")]
        .filter((element) => {
          if (!visible(element)) {
            return false;
          }

          const rect = element.getBoundingClientRect();
          const type = String(element.getAttribute("type") || "").toLowerCase();
          return type !== "hidden" && rect.width >= 80 && rect.height >= 18;
        });

      for (const labelElement of labelElements) {
        const labelRect = labelElement.getBoundingClientRect();
        const labelCenterY = labelRect.top + labelRect.height / 2;
        const candidates = controls
          .map((element) => ({ element, rect: element.getBoundingClientRect() }))
          .filter(({ rect }) => {
            const centerY = rect.top + rect.height / 2;
            return rect.left > labelRect.right - 4 && Math.abs(centerY - labelCenterY) <= 20;
          })
          .sort((a, b) => a.rect.left - b.rect.left);

        if (candidates[0]) {
          return candidates[0].element;
        }
      }

      return null;
    }, labelText).catch(() => null);

    const element = handle?.asElement?.();
    if (element) {
      return { frame, handle: element };
    }
  }

  return null;
}

async function clickOptionText(page, optionText) {
  for (const frame of page.frames()) {
    const option = frame.getByText(optionText, { exact: true }).last();
    const clicked = await option.click({ timeout: 1200, force: true }).then(() => true).catch(() => false);
    if (clicked) {
      return true;
    }
  }

  return false;
}

async function setPurchaseTransactionType(page, desiredType = PURCHASE_TRANSACTION_TYPE_VALUE) {
  for (let attempt = 0; attempt < 4; attempt += 1) {
    const inputs = await collectVisibleInputs(page);
    const alreadySelected = inputs.find((input) => {
      const text = `${input.value} ${input.selectedText} ${input.text}`.replace(/\s+/g, " ").trim();
      return text.includes(desiredType);
    });

    if (alreadySelected) {
      return { applied: true, method: "already-selected", value: desiredType };
    }

    const labeledControl = await findControlRightOfLabel(page, PURCHASE_TRANSACTION_TYPE_LABEL);

    if (labeledControl) {
      const box = await labeledControl.handle.boundingBox().catch(() => null);
      const tagName = await labeledControl.handle.evaluate((element) => String(element.tagName || "").toUpperCase()).catch(() => "");

      if (tagName === "SELECT") {
        await fillHandle(labeledControl, desiredType);
        return { applied: true, method: "select", value: desiredType };
      }

      if (box) {
        await page.mouse.click(box.x + box.width - 14, box.y + box.height / 2);
        await page.waitForTimeout(500);
        const clicked = await clickOptionText(page, desiredType);

        if (clicked) {
          await page.waitForTimeout(500);
          const current = await readHandleText(labeledControl.handle);
          if (current.includes(desiredType)) {
            return { applied: true, method: "dropdown-option", value: desiredType };
          }
        }
      }

      await fillHandle(labeledControl, desiredType);
      const current = await readHandleText(labeledControl.handle);
      if (current.includes(desiredType)) {
        return { applied: true, method: "direct-value", value: desiredType };
      }
    }

    await page.waitForTimeout(750);
  }

  await writeFile("tmp/ecount-purchase-transaction-type-inputs.json", JSON.stringify(await collectVisibleInputs(page), null, 2), "utf8").catch(() => {});
  throw new Error(`Could not set purchase transaction type to ${desiredType}.`);
}

async function getVisibleProductInput(page, productCode) {
  const normalizedCode = String(productCode).trim().toUpperCase();
  let inputs = [];
  let productInput = null;

  for (let attempt = 0; attempt < 8; attempt += 1) {
    inputs = await collectVisibleInputs(page);
    productInput = inputs
      .filter((input) => {
        const value = String(input.value || "").trim().toUpperCase();
        return value === normalizedCode && input.y > 430;
      })
      .sort((a, b) => {
        const aRowScore = Math.abs(a.y - 690);
        const bRowScore = Math.abs(b.y - 690);
        return aRowScore - bRowScore || a.y - b.y;
      })[0] || null;

    if (productInput) {
      return productInput;
    }

    await page.waitForTimeout(750);
  }

  if (!productInput) {
    const firstGridRow = inputs
      .filter((input) => {
        return input.type === "checkbox" &&
          input.x >= 180 &&
          input.x <= 240 &&
          input.y > 650;
      })
      .sort((a, b) => a.y - b.y)[0] || null;

    if (firstGridRow) {
      return {
        ...firstGridRow,
        value: normalizedCode,
        x: 259,
        y: firstGridRow.y,
        width: 68,
        height: firstGridRow.height
      };
    }

    await writeFile("tmp/ecount-purchase-fill-inputs.json", JSON.stringify(inputs, null, 2), "utf8").catch(() => {});
    throw new Error(`Could not locate product row input for ${productCode}.`);
  }

  return productInput;
}

async function fillTextAt(page, x, y, value) {
  await page.mouse.click(x, y);
  await page.waitForTimeout(200);
  await page.keyboard.press("Control+A");
  await page.keyboard.type(String(value));
  await page.keyboard.press("Tab");
  await page.waitForTimeout(250);
}

async function fillPurchaseForm(page, payload) {
  const transactionTypeResult = await setPurchaseTransactionType(page, payload.transactionType || PURCHASE_TRANSACTION_TYPE_VALUE);
  const currencyPlaceholder = "\ud1b5\ud654";
  const exchangeInput = page.locator(`input[placeholder="${currencyPlaceholder}"]`).last();
  await exchangeInput.fill(String(payload.exchangeRate));

  const productInput = await getVisibleProductInput(page, payload.productCode);
  const rowY = productInput.y + productInput.height / 2;

  await fillTextAt(page, productInput.x + 615, rowY, normalizeAmount(payload.unitPrice));
  await fillTextAt(page, productInput.x + 685, rowY, normalizeAmount(payload.foreignAmount));
  await fillTextAt(page, productInput.x + 755, rowY, normalizeAmount(payload.krwAmount));

  return { transactionTypeResult };
}

async function savePurchaseForm(page) {
  const saveText = "\uc800\uc7a5(F8)";
  let dialogMessage = "";

  page.on("dialog", async (dialog) => {
    dialogMessage = dialog.message();
    await dialog.accept();
  });

  const saveButton = page.getByText(saveText, { exact: true }).first();
  const clicked = await saveButton.click({ timeout: 5000 }).then(() => true).catch(() => false);

  if (!clicked) {
    await page.keyboard.press("F8").catch(() => {});
    await page.waitForTimeout(800);
    await page.mouse.click(271, 807).catch(() => {});
  }

  await page.waitForTimeout(6000);

  return { dialogMessage };
}

async function main() {
  required("ECOUNT_COM_CODE", ECOUNT_COM_CODE);
  required("ECOUNT_USER_ID", ECOUNT_USER_ID);
  required("ECOUNT_PASSWORD", ECOUNT_PASSWORD);
  required("payload path", payloadPath);

  const payload = JSON.parse(await readFile(payloadPath, "utf8"));
  const requiredPayload = ["productCode", "quantity", "exchangeRate", "unitPrice", "foreignAmount", "krwAmount"];
  const autoSave = payload.autoSave === true || ECOUNT_PURCHASE_AUTO_SAVE === "1";

  for (const key of requiredPayload) {
    required(`payload.${key}`, payload[key]);
  }

  const browser = await chromium.launch({
    channel: "chrome",
    headless: HEADLESS === "true"
  });
  const page = await browser.newPage({ viewport: { width: 1420, height: 900 } });
  let searchRows = [];

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
      console.error(`[ecount-purchase-fill] captured ${searchRows.length} purchase rows`);
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

    setStep("select row");
    const sameQuantityRows = searchRows.filter((row) => {
      return Math.abs(toNumber(row.QTY || row["BUY.QTY"]) - toNumber(payload.quantity)) < 0.0001;
    });
    const targetRow = sortPurchaseRows(sameQuantityRows)[0];

    if (!targetRow) {
      throw new Error(`No purchase row found for ${payload.productCode} with quantity ${payload.quantity}.`);
    }

    const purchaseNo = targetRow["BUY.IO_DATE_NO"];
    const visiblePurchaseNo = String(purchaseNo).replace(/^20(\d{2}\/)/, "$1");
    await page.getByText(visiblePurchaseNo, { exact: true }).first().dblclick();
    await page.waitForTimeout(7000);

    setStep("fill form");
    let fillResult = null;
    try {
      fillResult = await fillPurchaseForm(page, payload);
    } catch (error) {
      await page.screenshot({ path: "tmp/ecount-purchase-fill-error.png", fullPage: true }).catch(() => {});
      throw error;
    }
    await page.screenshot({ path: "tmp/ecount-purchase-filled.png", fullPage: true });

    let saveResult = null;

    if (autoSave) {
      setStep("save form");
      saveResult = await savePurchaseForm(page);
      await page.screenshot({ path: "tmp/ecount-purchase-saved.png", fullPage: true }).catch(() => {});
    }

    const result = {
      ok: true,
      saved: autoSave,
      purchaseNo,
      visiblePurchaseNo,
      productCode: payload.productCode,
      quantity: payload.quantity,
      exchangeRate: payload.exchangeRate,
      unitPrice: normalizeAmount(payload.unitPrice),
      foreignAmount: normalizeAmount(payload.foreignAmount),
      krwAmount: normalizeAmount(payload.krwAmount),
      erpUnitPrice: normalizeAmount(payload.unitPrice),
      erpForeignAmount: normalizeAmount(payload.foreignAmount),
      erpKrwAmount: normalizeAmount(payload.krwAmount),
      transactionType: fillResult?.transactionTypeResult?.value || "",
      transactionTypeMethod: fillResult?.transactionTypeResult?.method || "",
      dialogMessage: saveResult?.dialogMessage || "",
      note: autoSave
        ? "Values were entered and the purchase save button was clicked."
        : "Values were entered on the purchase edit screen. Save was not clicked."
    };
    await writeFile("tmp/ecount-purchase-fill-result.json", JSON.stringify(result, null, 2), "utf8");
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
  const result = {
    ok: false,
    step: currentStep,
    error: error instanceof Error ? error.message : String(error)
  };
  await writeFile("tmp/ecount-purchase-fill-result.json", JSON.stringify(result, null, 2), "utf8").catch(() => {});
  console.error(JSON.stringify(result));
  process.exit(1);
});
