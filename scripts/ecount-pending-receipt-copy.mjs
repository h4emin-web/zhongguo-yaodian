import { chromium } from "playwright";
import { readFile, writeFile } from "node:fs/promises";

const {
  ECOUNT_COM_CODE,
  ECOUNT_USER_ID,
  ECOUNT_PASSWORD,
  HEADLESS = "false",
  ECOUNT_PENDING_AUTO_SAVE = "0",
  ECOUNT_PENDING_EMPLOYEE_NAME = "",
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

function normalizeCode(value) {
  return String(value || "").toUpperCase().replace(/[^0-9A-Z]/g, "");
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

async function openPurchaseSearch(page) {
  const clickedByText = await page.getByText("구매조회", { exact: true }).first().click({ timeout: 5000 }).then(() => true).catch(() => false);
  if (!clickedByText) {
    await page.mouse.click(153, 424);
  }
  await page.waitForTimeout(4500);
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

async function fillProductSearch(page, productCode) {
  let inputs = [];

  for (let attempt = 0; attempt < 12; attempt += 1) {
    inputs = await collectVisibleInputs(page);
    const searchInput = inputs
      .filter((input) => String(input.placeholder || "").includes("입력 후") && input.y > 80)
      .sort((a, b) => b.x - a.x || a.y - b.y)[0] || null;

    if (searchInput) {
      await fillHandle(searchInput, productCode);
      await searchInput.handle.press("Enter").catch(async () => {
        await searchInput.frame.page().keyboard.press("Enter");
      });
      return { method: "placeholder-search", x: searchInput.x, y: searchInput.y };
    }

    const fallback = page.locator("input").nth(1);
    const fallbackVisible = await fallback.isVisible({ timeout: 300 }).catch(() => false);
    if (fallbackVisible) {
      await fallback.click();
      await fallback.fill(productCode);
      await fallback.press("Enter");
      return { method: "nth-input" };
    }

    await page.waitForTimeout(1000);
  }

  await writeFile("tmp/ecount-pending-search-inputs.json", JSON.stringify(inputs.map(({ frame, handle, ...input }) => input), null, 2), "utf8").catch(() => {});
  throw new Error("Could not locate purchase product search input.");
}

async function collectVisibleInputs(page) {
  const inputs = [];

  for (const frame of page.frames()) {
    const handles = await frame.locator("input, select, textarea").elementHandles().catch(() => []);

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
  await input.handle.evaluate((element, nextValue) => {
    const tagName = String(element.tagName || "").toUpperCase();
    if (tagName === "SELECT") {
      const desired = String(nextValue ?? "").trim();
      const normalizedDesired = desired.replace(/^0+(\d)$/, "$1");
      const option = [...element.options].find((entry) => {
        const optionValue = String(entry.value || "").trim();
        const optionText = String(entry.textContent || "").trim();
        return optionValue === desired ||
          optionText === desired ||
          optionValue === normalizedDesired ||
          optionText === normalizedDesired ||
          optionText.includes(desired);
      });
      if (option) {
        element.value = option.value;
      } else {
        element.value = desired;
      }
      element.dispatchEvent(new Event("input", { bubbles: true }));
      element.dispatchEvent(new Event("change", { bubbles: true }));
      element.blur?.();
      return;
    }

    element.focus?.();
    element.value = String(nextValue ?? "");
    element.dispatchEvent(new Event("input", { bubbles: true }));
    element.dispatchEvent(new Event("change", { bubbles: true }));
    element.blur?.();
  }, value);
  await input.frame.waitForTimeout(150).catch(() => {});
}

async function typeHandle(input, value) {
  await input.handle.click({ force: true });
  await input.frame.waitForTimeout(150).catch(() => {});
  const page = input.frame.page();
  await page.keyboard.press("Control+A");
  if (String(value || "") === "") {
    await page.keyboard.press("Backspace");
  } else {
    await page.keyboard.type(String(value));
  }
  await page.keyboard.press("Tab");
  await input.frame.waitForTimeout(250).catch(() => {});
}

async function clickCopyButton(page) {
  const controls = await collectVisibleControls(page);
  await writeFile("tmp/ecount-pending-before-copy-controls.json", JSON.stringify(controls.map(({ frame, handle, ...control }) => control), null, 2), "utf8").catch(() => {});
  const normalize = (value) => String(value || "").replace(/\s+/g, "");
  const controlText = (control) => normalize(`${control.text} ${control.title} ${control.ariaLabel}`);
  const isButtonLike = (control) => /btn|button/i.test(control.className || "") || control.width >= 38;
  const clickControl = async (copyControl, prefix) => {
    const context = page.context();
    const pagesBefore = new Set(context.pages());
    const popupPromise = context.waitForEvent("page", { timeout: 5000 }).catch(() => null);
    let descriptor = `${prefix}:${copyControl.text || copyControl.title || copyControl.ariaLabel || "copy"}:${copyControl.id || "no-id"}`;
    let clickError = null;

    try {
      if (copyControl.id) {
        await copyControl.frame.evaluate((id) => {
          const element = document.getElementById(id);
          if (!element) {
            throw new Error(`copy control #${id} was not found`);
          }
          element.click();
        }, copyControl.id);
        descriptor = `${descriptor}:frame-id-click`;
      } else {
        await copyControl.handle.evaluate((element) => {
          element.click();
        });
        descriptor = `${descriptor}:handle-click`;
      }
    } catch (error) {
      clickError = error;
    }

    const popup = await popupPromise;
    if (popup && !popup.isClosed()) {
      await popup.waitForLoadState("domcontentloaded", { timeout: 15_000 }).catch(() => {});
      await popup.bringToFront().catch(() => {});
      await popup.waitForTimeout(1000).catch(() => {});
      return { descriptor: `${descriptor}:new-page`, page: popup };
    }

    if (page.isClosed()) {
      const replacementPage = context.pages().find((candidate) => !pagesBefore.has(candidate) && !candidate.isClosed()) ||
        context.pages().find((candidate) => !candidate.isClosed());
      if (replacementPage) {
        await replacementPage.bringToFront().catch(() => {});
        await replacementPage.waitForTimeout(1000).catch(() => {});
        return { descriptor: `${descriptor}:recovered-page`, page: replacementPage };
      }
      throw new Error("Copy button closed the ERP page before the copied form could be edited.");
    }

    if (clickError) {
      throw clickError;
    }

    await page.waitForTimeout(1000);
    return { descriptor, page };
  };

  const slipCopy = controls.find((control) => control.id === "slipCopy");
  if (slipCopy) {
    return await clickControl(slipCopy, "slipCopy");
  }

  const formCopyCandidates = controls
    .filter((control) => {
      const text = controlText(control);
      return text === "복사" &&
        isButtonLike(control) &&
        control.width >= 38 &&
        control.height >= 18 &&
        control.y > 450 &&
        control.x < 900 &&
        !text.includes("사원") &&
        !text.includes("검색");
    })
    .sort((a, b) => a.y - b.y || a.x - b.x);
  if (formCopyCandidates[0]) {
    return await clickControl(formCopyCandidates[0], "purchase-form-copy");
  }

  const saveCandidates = controls
    .filter((control) => {
      const text = controlText(control);
      return isButtonLike(control) &&
        (control.id === "group3slipSave" || control.id === "slipSave" || text === "저장(F8)") &&
        control.y > 450;
    })
    .sort((a, b) => a.y - b.y || a.x - b.x);

  for (const save of saveCandidates) {
    const copyControl = controls
      .filter((control) => control.x > save.x &&
        control.x < save.x + 320 &&
        Math.abs(control.centerY - save.centerY) <= 24 &&
        controlText(control) === "복사" &&
        isButtonLike(control))
      .sort((a, b) => a.x - b.x)[0] || null;

    if (copyControl) {
      return await clickControl(copyControl, "save-row-copy");
    }
  }

  await writeFile("tmp/ecount-pending-copy-controls.json", JSON.stringify(controls.map(({ frame, handle, ...control }) => control), null, 2), "utf8").catch(() => {});
  throw new Error("Could not locate slipCopy purchase form copy button.");
}

async function acceptCopyConfirmation(page, previousDialogCount) {
  const startedAt = Date.now();
  let clickedDomConfirm = false;
  let latestFrameText = "";
  let latestControls = [];
  const normalize = (value) => String(value || "").replace(/\s+/g, "");
  const isConfirmText = (value) => {
    const text = normalize(value);
    return text === "확인" || text === "예" || text.toUpperCase() === "OK";
  };
  const clickConfirmControl = async (control, method, message) => {
    await control.handle.evaluate((element) => {
      element.click();
    });
    clickedDomConfirm = true;
    await page.waitForTimeout(3000);
    return {
      accepted: true,
      method,
      message
    };
  };

  while (Date.now() - startedAt < 12_000) {
    if (page.isClosed()) {
      throw new Error("ERP page closed while waiting for copy confirmation.");
    }

    if (dialogMessages.length > previousDialogCount) {
      await page.waitForTimeout(1500);
      return {
        accepted: true,
        method: "native-dialog",
        message: dialogMessages[dialogMessages.length - 1] || ""
      };
    }

    for (const frame of page.frames()) {
      const promptLocator = frame.getByText(/시리얼.*로트.*복사|로트No.*복사|로트번호.*복사/).first();
      const promptVisible = await promptLocator.isVisible({ timeout: 300 }).catch(() => false);
      if (promptVisible) {
        const confirmText = frame.getByText(/^확인$/).last();
        const clickedConfirmText = await confirmText.click({ timeout: 1500, force: true }).then(() => true).catch(() => false);
        if (clickedConfirmText) {
          clickedDomConfirm = true;
          await page.waitForTimeout(3000);
          return {
            accepted: true,
            method: "alert-confirm-text",
            message: await promptLocator.innerText({ timeout: 500 }).catch(() => "")
          };
        }

        const promptBox = await promptLocator.boundingBox().catch(() => null);
        if (promptBox) {
          await page.mouse.click(promptBox.x, promptBox.y + 180);
          clickedDomConfirm = true;
          await page.waitForTimeout(3000);
          return {
            accepted: true,
            method: "alert-confirm-coordinate",
            message: await promptLocator.innerText({ timeout: 500 }).catch(() => "")
          };
        }
      }

      const frameText = await frame.locator("body").innerText({ timeout: 500 }).catch(() => "");
      if (frameText) {
        latestFrameText = frameText;
      }
      const compactFrameText = normalize(frameText);
      const hasCopyPrompt =
        compactFrameText.includes("복사하겠습니까") ||
        compactFrameText.includes("복사하시겠습니까") ||
        (
          compactFrameText.includes("값이있습니다") &&
          compactFrameText.includes("복사") &&
          (compactFrameText.includes("시리얼") || compactFrameText.includes("로트No") || compactFrameText.includes("로트번호"))
        );

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

        if (isConfirmText(text) || text.includes("확인")) {
          await handle.evaluate((element) => {
            element.click();
          });
          clickedDomConfirm = true;
          await page.waitForTimeout(3000);
          return {
            accepted: true,
            method: "dom-confirm",
            message: frameText.replace(/\s+/g, " ").trim().slice(0, 160)
          };
        }
      }

      await page.keyboard.press("Enter").catch(() => {});
      clickedDomConfirm = true;
      await page.waitForTimeout(3000);
      return {
        accepted: true,
        method: "keyboard-enter",
        message: frameText.replace(/\s+/g, " ").trim().slice(0, 160)
      };
    }

    latestControls = await collectVisibleControls(page).catch(() => []);
    const centeredConfirm = latestControls
      .filter((control) => isConfirmText(`${control.text} ${control.title} ${control.ariaLabel}`) &&
        control.y > 180 &&
        control.y < 760 &&
        control.x > 300 &&
        control.x < 1120 &&
        control.width >= 36 &&
        control.height >= 18)
      .sort((a, b) => {
        const aCenterScore = Math.abs((a.x + a.width / 2) - 710) + Math.abs((a.y + a.height / 2) - 450);
        const bCenterScore = Math.abs((b.x + b.width / 2) - 710) + Math.abs((b.y + b.height / 2) - 450);
        return aCenterScore - bCenterScore;
      })[0] || null;

    if (centeredConfirm) {
      return await clickConfirmControl(
        centeredConfirm,
        "centered-confirm",
        latestFrameText.replace(/\s+/g, " ").trim().slice(0, 160)
      );
    }

    await page.waitForTimeout(500);
  }

  await writeFile("tmp/ecount-pending-confirm-controls.json", JSON.stringify(latestControls.map(({ frame, handle, ...control }) => control), null, 2), "utf8").catch(() => {});
  return {
    accepted: clickedDomConfirm || dialogMessages.length > previousDialogCount,
    method: clickedDomConfirm ? "dom-confirm" : "",
    message: dialogMessages.slice(previousDialogCount).join(" / ") || latestFrameText.replace(/\s+/g, " ").trim().slice(0, 160)
  };
}

async function getVisibleProductInput(page, productCode) {
  const normalizedCode = normalizeCode(productCode);
  let inputs = [];
  let controls = [];

  for (let attempt = 0; attempt < 8; attempt += 1) {
    inputs = await collectVisibleInputs(page);
    const productInput = inputs
      .filter((input) => {
        const value = normalizeCode(input.value);
        return value === normalizedCode && input.y > 430;
      })
      .sort((a, b) => Math.abs(a.y - 690) - Math.abs(b.y - 690) || a.y - b.y)[0] || null;

    if (productInput) {
      return productInput;
    }

    controls = await collectVisibleControls(page);
    const productControl = controls
      .filter((control) => {
        const value = normalizeCode(`${control.text} ${control.title} ${control.ariaLabel}`);
        return value === normalizedCode && control.y > 430;
      })
      .sort((a, b) => Math.abs(a.y - 500) - Math.abs(b.y - 500) || a.y - b.y)[0] || null;

    if (productControl) {
      return productControl;
    }

    const formSave = controls
      .filter((control) => control.id === "group3slipSave" && control.y > 600)
      .sort((a, b) => b.y - a.y)[0] || null;
    if (formSave) {
      return {
        index: -1,
        value: normalizedCode,
        placeholder: "",
        title: "",
        name: "",
        id: "fallback-grid-product-code",
        type: "fallback",
        x: Number(process.env.ECOUNT_PENDING_PRODUCT_X || "259.40625"),
        y: formSave.y - Number(process.env.ECOUNT_PENDING_PRODUCT_SAVE_Y_OFFSET || "216.71875"),
        width: Number(process.env.ECOUNT_PENDING_PRODUCT_WIDTH || "68.3125"),
        height: Number(process.env.ECOUNT_PENDING_PRODUCT_HEIGHT || "28")
      };
    }

    await page.waitForTimeout(750);
  }

  await writeFile("tmp/ecount-pending-inputs.json", JSON.stringify(inputs.map(({ frame, handle, ...input }) => input), null, 2), "utf8").catch(() => {});
  await writeFile("tmp/ecount-pending-product-controls.json", JSON.stringify(controls.map(({ frame, handle, ...control }) => control), null, 2), "utf8").catch(() => {});
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

async function editGridCellAt(page, x, y, value, label) {
  await page.mouse.dblclick(x, y);
  await page.waitForTimeout(250);
  await page.keyboard.press("Control+A");
  if (String(value || "") === "") {
    await page.keyboard.press("Backspace");
  } else {
    await page.keyboard.type(String(value));
  }
  await page.keyboard.press("Tab");
  await page.waitForTimeout(350);
  return { label, x, y };
}

async function fillInputNear(inputs, targetX, rowY, value, label) {
  const target = inputs
    .filter((input) => Math.abs((input.y + input.height / 2) - rowY) <= 24)
    .sort((a, b) => Math.abs((a.x + a.width / 2) - targetX) - Math.abs((b.x + b.width / 2) - targetX))[0] || null;

  if (!target) {
    await writeFile("tmp/ecount-pending-inputs.json", JSON.stringify(inputs.map(({ frame, handle, ...input }) => input), null, 2), "utf8").catch(() => {});
    throw new Error(`Could not locate ${label} input near x=${targetX}, y=${rowY}.`);
  }

  await fillHandle(target, value);
  return true;
}

async function fillEmployeeSearchName(page, payload) {
  const employeeName = String(payload.employeeName || ECOUNT_PENDING_EMPLOYEE_NAME || ECOUNT_USER_ID || "강해민").trim();
  if (!employeeName) {
    return false;
  }

  for (let attempt = 0; attempt < 8; attempt += 1) {
    const inputs = await collectVisibleInputs(page);
    const managerCodeInput = inputs
      .filter((input) =>
        input.y > 0 &&
        input.y < 220 &&
        input.x > 300 &&
        input.x < 540 &&
        (String(input.placeholder || "").includes("담당자") || /^\d{6,}$/.test(String(input.value || "").trim())))
      .sort((a, b) => a.y - b.y || a.x - b.x)[0] || null;

    if (managerCodeInput) {
      const nameInput = inputs
        .filter((input) =>
          Math.abs((input.y + input.height / 2) - (managerCodeInput.y + managerCodeInput.height / 2)) <= 10 &&
          input.x > managerCodeInput.x + managerCodeInput.width &&
          input.x < managerCodeInput.x + managerCodeInput.width + 260)
        .sort((a, b) => a.x - b.x)[0] || null;

      if (nameInput) {
        if (String(nameInput.value || "").trim() === employeeName) {
          return true;
        }
        await typeHandle(nameInput, employeeName);
        return true;
      }
    }

    await page.waitForTimeout(500);
  }

  return false;
}

async function fillPurchaseDate(page, yyyymmdd) {
  const match = String(yyyymmdd || "").match(/^(\d{4})(\d{2})(\d{2})$/);
  if (!match) {
    return false;
  }

  const [, year, month, day] = match;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    const inputs = await collectVisibleInputs(page);
    const dayInput = inputs
      .filter((input) => {
        const value = String(input.value || "").trim();
        return String(input.tagName || "").toUpperCase() !== "SELECT" &&
          input.y > 0 &&
          input.y < 200 &&
          input.x > 440 &&
          input.x < 540 &&
          /^\d{1,2}$/.test(value);
      })
      .sort((a, b) => a.y - b.y || a.x - b.x)[0] || null;

    if (dayInput) {
      const rowCenter = dayInput.y + dayInput.height / 2;
      const sameRow = inputs
        .filter((input) => Math.abs((input.y + input.height / 2) - rowCenter) <= 10)
        .sort((a, b) => a.x - b.x);
      const selects = sameRow.filter((input) => String(input.tagName || "").toUpperCase() === "SELECT" && input.x < dayInput.x);

      if (selects[0]) {
        await fillHandle(selects[0], year);
      }
      if (selects[1]) {
        await fillHandle(selects[1], month);
      }
      await typeHandle(dayInput, day);
      return true;
    }

    await page.waitForTimeout(500);
  }

  return false;
}

async function clearPendingFields(page, payload) {
  const yyyymmdd = normalizeDate(payload.instockDate);
  const quantity = normalizeAmount(payload.quantity);
  const inputs = await collectVisibleInputs(page);
  await writeFile("tmp/ecount-pending-clear-inputs.json", JSON.stringify(inputs.map(({ frame, handle, ...input }) => input), null, 2), "utf8").catch(() => {});
  let dateFilled = false;
  let serialCleared = 0;
  let moneyCleared = 0;
  let quantityFilled = false;
  let employeeFilled = false;

  employeeFilled = await fillEmployeeSearchName(page, payload);
  dateFilled = await fillPurchaseDate(page, yyyymmdd);

  for (const input of inputs) {
    const meta = `${input.placeholder} ${input.title} ${input.name} ${input.id}`.toLowerCase();

    if (meta.includes("serial") || meta.includes("시리얼") || meta.includes("lot") || meta.includes("batch")) {
      await fillHandle(input, "");
      serialCleared += 1;
    }

    if (meta.includes("환율") || meta.includes("exchange") || meta.includes("통화") || meta.includes("단가") || meta.includes("금액") || meta.includes("공급") || meta.includes("부가")) {
      await fillHandle(input, "");
      moneyCleared += 1;
    }
  }

  const productInput = await getVisibleProductInput(page, payload.productCode);
  const rowY = productInput.y + productInput.height / 2;
  const quantityOffset = Number(process.env.ECOUNT_PENDING_QTY_OFFSET || "525");

  if (quantity) {
    await editGridCellAt(page, productInput.x + quantityOffset, rowY, quantity, "quantity");
    quantityFilled = true;
  }

  await editGridCellAt(page, productInput.x + 615, rowY, "", "unit price");
  await editGridCellAt(page, productInput.x + 685, rowY, "", "foreign amount");
  await editGridCellAt(page, productInput.x + 755, rowY, "", "supply amount");
  await editGridCellAt(page, productInput.x + 925, rowY, "", "serial lot");
  moneyCleared += 4;
  serialCleared += 1;

  return { dateFilled, serialCleared, moneyCleared, quantityFilled, employeeFilled };
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
  let activePage = page;
  let searchRows = [];

  const dialogHandledPages = new WeakSet();
  const attachDialogHandler = (targetPage) => {
    if (!targetPage || dialogHandledPages.has(targetPage)) {
      return;
    }

    dialogHandledPages.add(targetPage);
    targetPage.on("dialog", async (dialog) => {
      dialogMessages.push(dialog.message());
      await dialog.accept();
    });
  };
  attachDialogHandler(page);
  page.context().on("page", attachDialogHandler);

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
    await openPurchaseSearch(page);

    setStep("search product");
    await fillProductSearch(page, payload.productCode);
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

    setStep("click purchase copy button");
    const dialogCountBeforeCopy = dialogMessages.length;
    const copyResult = await clickCopyButton(activePage);
    activePage = copyResult.page;
    if (activePage !== page) {
      attachDialogHandler(activePage);
    }
    const copyButton = copyResult.descriptor;
    setStep("confirm purchase copy dialog");
    const copyConfirm = await acceptCopyConfirmation(activePage, dialogCountBeforeCopy);
    await activePage.screenshot({ path: "tmp/ecount-pending-after-copy-confirm.png", fullPage: true }).catch(() => {});
    await collectVisibleControls(activePage)
      .then((controls) => writeFile("tmp/ecount-pending-after-copy-confirm-controls.json", JSON.stringify(controls.map(({ frame, handle, ...control }) => control), null, 2), "utf8"))
      .catch(() => {});
    setStep("clear copied purchase form");
    const clearResult = await clearPendingFields(activePage, payload);
    await activePage.screenshot({ path: "tmp/ecount-pending-receipt-filled.png", fullPage: true }).catch(() => {});

    let saveResult = null;
    const autoSave = payload.autoSave === true || ECOUNT_PENDING_AUTO_SAVE === "1";

    if (autoSave) {
      setStep("save form");
      saveResult = await savePurchaseForm(activePage);
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
