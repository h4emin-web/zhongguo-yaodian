const cards = document.querySelectorAll(".tool-box");
const detailView = document.querySelector(".detail-view");
const detailPanel = document.querySelector(".detail-panel");
const closeButton = document.querySelector(".detail-close");
const detailKicker = document.querySelector(".detail-kicker");
const detailTitle = document.querySelector("#detail-title");
const detailDescription = document.querySelector(".detail-description");
const searchInput = document.querySelector("#workspace-search");
const wcSearchPanel = document.querySelector(".wc-search-panel");
const wcSearchInput = document.querySelector("#wc-search-input");
const wcResults = document.querySelector(".wc-results");
const excelExportButton = document.querySelector(".wc-excel-export");
const poReceivePanel = document.querySelector(".po-receive-panel");
const poReceiveInput = document.querySelector("#po-receive-input");
const poReceiveSubmit = document.querySelector(".po-receive-submit");
const poReceiveResult = document.querySelector(".po-receive-result");
const mfdsSearchPanel = document.querySelector(".mfds-search-panel");
const mfdsSearchInput = document.querySelector("#mfds-search-input");
const mfdsResults = document.querySelector(".mfds-results");
const mfdsExcelExportButton = document.querySelector(".mfds-excel-export");
const cnphSearchPanel = document.querySelector(".cnph-search-panel");
const cnphSearchInput = document.querySelector("#cnph-search-input");
const cnphResults = document.querySelector(".cnph-results");
const cnphExcelExportButton = document.querySelector(".cnph-excel-export");
const usdmfSearchPanel = document.querySelector(".usdmf-search-panel");
const usdmfSearchInput = document.querySelector("#usdmf-search-input");
const usdmfResults = document.querySelector(".usdmf-results");
const usdmfExcelExportButton = document.querySelector(".usdmf-excel-export");
const cnphCriteriaToggle = document.querySelector(".cnph-criteria-toggle");
const cnphCriteria = document.querySelector(".cnph-criteria");
const globalSearchHead = document.querySelector(".global-search-head");
const globalSearchResults = document.querySelector(".global-search-results");
const globalExcelExportButton = document.querySelector(".global-excel-export");
const worklogFullscreen = document.querySelector(".worklog-fullscreen");
const worklogBoxDate = document.querySelector(".worklog-box-date");
const worklogDateMeta = document.querySelector(".worklog-date-meta");
const worklogClose = document.querySelector(".worklog-close");
const worklogLoadButton = document.querySelector(".worklog-load-button");
const worklogFileInput = document.querySelector(".worklog-file-input");
const worklogFilterInput = document.querySelector("#worklog-filter-input");
const worklogClearButton = document.querySelector(".worklog-clear-button");
const toolsMenu = document.querySelector(".tools-menu");
const toolsTrigger = document.querySelector(".tools-trigger");
const toolsDropdown = document.querySelector(".tools-dropdown");
const marginCalculatorItem = document.querySelector('.tools-item[data-tool="margin-calculator"]');
const marginCalcPanel = document.querySelector(".margin-calc-panel");
const marginCostInput = document.querySelector("#margin-cost-input");
const marginPriceInput = document.querySelector("#margin-price-input");
const marginResult = document.querySelector(".margin-calc-result");
const importCostItem = document.querySelector('.tools-item[data-tool="import-cost-calculator"]');
const importCostPanel = document.querySelector(".import-cost-panel");
const importPriceInput = document.querySelector("#import-price-input");
const importPriceCurrency = document.querySelector("#import-price-currency");
const importRateInput = document.querySelector("#import-rate-input");
const importQtyInput = document.querySelector("#import-qty-input");
const importQtyUnit = document.querySelector("#import-qty-unit");
const importShippingSelect = document.querySelector("#import-shipping-select");
const importDutyInput = document.querySelector("#import-duty-input");
const importCostResult = document.querySelector(".import-cost-result");
const worklogDropzone = document.querySelector(".worklog-dropzone");
const worklogResults = document.querySelector(".worklog-results");
const supabaseConfig = window.HAEMIN_SUPABASE || {};
const hasSupabaseKey = Boolean(
  supabaseConfig.anonKey && supabaseConfig.anonKey !== "YOUR_SUPABASE_ANON_KEY"
);
const supabaseClient = hasSupabaseKey && window.supabase
  ? window.supabase.createClient(supabaseConfig.url, supabaseConfig.anonKey)
  : null;

let lastFocusedCard = null;
let wcRows = [];
let wcLoaded = false;
let currentWcResults = [];
let mfdsRows = [];
let mfdsLoaded = false;
let currentMfdsResults = [];
let cnphRows = [];
let cnphLoaded = false;
let currentCnphResults = [];
let usdmfRows = [];
let usdmfLoaded = false;
let currentUsdmfResults = [];
let globalSearchToken = 0;
let currentGlobalKeyword = "";
let currentGlobalMatches = { wc: [], mfds: [], cnph: [] };
let worklogPeople = [];
let worklogDateLabel = "";

function openDetail(card) {
  const title = card.querySelector("h2").textContent;

  if (title.includes("업무일지")) {
    openWorklog(card);
    return;
  }

  const kicker = card.querySelector(".tool-kicker").textContent;
  const description = card.querySelector("p:last-child").textContent;

  lastFocusedCard = card;
  detailKicker.textContent = kicker;
  detailTitle.textContent = title;
  detailDescription.textContent = description;
  showWcSearch(title.includes("중국 WC & COPP"));
  showPoReceive(title.includes("PO 입고처리"));
  showMfdsSearch(title.includes("K-DMF 검색"));
  showCnphSearch(title.includes("중국 약전"));
  showUsdmfSearch(title.includes("미국 DMF"));
  showMarginCalc(false);
  showImportCostCalc(false);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");

  if (title.includes("중국 WC & COPP")) {
    loadWcData();
    wcSearchInput.focus();
  } else if (title.includes("PO 입고처리")) {
    poReceiveInput.focus();
  } else if (title.includes("K-DMF 검색")) {
    loadMfdsData();
    mfdsSearchInput.focus();
  } else if (title.includes("중국 약전")) {
    loadCnphData();
    cnphSearchInput.focus();
  } else if (title.includes("미국 DMF")) {
    loadUsdmfData();
    usdmfSearchInput.focus();
  } else {
    closeButton.focus();
  }
}

function closeDetail() {
  detailView.classList.remove("is-open");
  detailView.setAttribute("aria-hidden", "true");
  document.body.classList.remove("detail-open");

  if (lastFocusedCard) {
    lastFocusedCard.focus();
  }
}

cards.forEach((card) => {
  card.addEventListener("click", () => openDetail(card));
  card.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      openDetail(card);
    }
  });
});

async function shareCardAsImage(card, button) {
  if (!window.html2canvas) {
    alert("이미지 캡처 도구를 불러오지 못했습니다.");
    return;
  }

  const title = card.querySelector("h2").textContent.trim();

  button.disabled = true;

  try {
    const canvas = await window.html2canvas(card, { backgroundColor: "#ffffff", scale: 2 });

    canvas.toBlob(async (blob) => {
      button.disabled = false;

      if (!blob) {
        alert("이미지를 만들지 못했습니다.");
        return;
      }

      const file = new File([blob], `${title}.png`, { type: "image/png" });

      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        try {
          await navigator.share({ files: [file], title });
          return;
        } catch (shareError) {
          if (shareError.name === "AbortError") {
            return;
          }
          console.error(shareError);
        }
      }

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");

      link.href = url;
      link.download = `${title}.png`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);
      alert("이 브라우저는 공유가 지원되지 않아 이미지를 다운로드했습니다. 카카오톡에서 사진 첨부로 보내주세요.");
    }, "image/png");
  } catch (error) {
    button.disabled = false;
    console.error(error);
    alert("카드 이미지를 캡처하지 못했습니다.");
  }
}

document.querySelectorAll(".card-share-button").forEach((button) => {
  button.addEventListener("click", (event) => {
    event.stopPropagation();
    shareCardAsImage(button.closest(".tool-box"), button);
  });
});

closeButton.addEventListener("click", closeDetail);

detailView.addEventListener("click", (event) => {
  if (!detailPanel.contains(event.target)) {
    closeDetail();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && detailView.classList.contains("is-open")) {
    closeDetail();
  }

  if (event.key === "Escape" && worklogFullscreen.classList.contains("is-open")) {
    closeWorklog();
  }

  if (event.key === "Escape" && document.activeElement && document.activeElement.closest(".tools-menu")) {
    closeToolsDropdown();
  }
});

toolsMenu.addEventListener("mouseenter", () => {
  toolsTrigger.setAttribute("aria-expanded", "true");
});

toolsMenu.addEventListener("mouseleave", () => {
  toolsTrigger.setAttribute("aria-expanded", "false");
});

marginCalculatorItem.addEventListener("click", () => {
  closeToolsDropdown();
  openMarginCalculator();
});

marginCostInput.addEventListener("input", renderMarginResult);
marginPriceInput.addEventListener("input", renderMarginResult);

importCostItem.addEventListener("click", () => {
  closeToolsDropdown();
  openImportCostCalculator();
});

[importPriceInput, importPriceCurrency, importRateInput, importQtyInput, importQtyUnit, importShippingSelect, importDutyInput]
  .forEach((el) => {
    el.addEventListener("input", renderImportCostResult);
    el.addEventListener("change", renderImportCostResult);
  });

let globalSearchDebounce = null;

searchInput.addEventListener("input", () => {
  const keyword = searchInput.value.trim().toLowerCase();

  cards.forEach((card) => {
    const cardText = card.textContent.toLowerCase();
    card.hidden = keyword.length > 0 && !cardText.includes(keyword);
  });

  clearTimeout(globalSearchDebounce);
  globalSearchDebounce = setTimeout(() => renderGlobalSearch(keyword), 300);
});

async function ensureGlobalSearchData() {
  await Promise.all([loadWcData(), loadMfdsData(), loadCnphData()]);
}

const GLOBAL_ITEM_RENDER_LIMIT = 150;

const FLAG_CN_SVG = `<svg class="flag-cn" viewBox="0 0 30 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
  <rect width="30" height="20" fill="#DE2910"/>
  <polygon fill="#FFDE00" points="5.00,2.00 5.71,4.03 7.85,4.07 6.14,5.37 6.76,7.43 5.00,6.20 3.24,7.43 3.86,5.37 2.15,4.07 4.29,4.03"/>
  <polygon fill="#FFDE00" points="10.39,1.08 10.34,1.79 11.00,2.09 10.30,2.26 10.22,2.97 9.84,2.37 9.14,2.52 9.60,1.97 9.25,1.34 9.91,1.61"/>
  <polygon fill="#FFDE00" points="12.32,3.31 12.00,3.94 12.48,4.47 11.78,4.36 11.43,4.98 11.31,4.28 10.61,4.14 11.25,3.81 11.16,3.10 11.67,3.61"/>
  <polygon fill="#FFDE00" points="12.53,6.24 11.99,6.70 12.23,7.38 11.62,7.00 11.06,7.44 11.23,6.74 10.63,6.34 11.35,6.29 11.55,5.60 11.82,6.26"/>
  <polygon fill="#FFDE00" points="11.00,8.60 10.32,8.84 10.31,9.55 9.88,8.98 9.19,9.19 9.60,8.60 9.19,8.01 9.88,8.22 10.31,7.65 10.32,8.36"/>
</svg>`;

const FLAG_KR_SVG = `<svg class="flag-kr" viewBox="0 0 30 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
  <rect width="30" height="20" fill="#FFFFFF"/>
  <circle cx="15" cy="10" r="5.5" fill="#CD2E3A"/>
  <path d="M15,4.5 A5.5,5.5 0 0,1 15,15.5 A2.75,2.75 0 0,1 15,10 A2.75,2.75 0 0,0 15,4.5 Z" fill="#0047A0"/>
  <g fill="#000000">
    <g transform="translate(4.2 3.6) rotate(-45)">
      <rect x="-2" y="-1.75" width="4" height="0.7"/>
      <rect x="-2" y="-0.35" width="4" height="0.7"/>
      <rect x="-2" y="1.05" width="4" height="0.7"/>
    </g>
    <g transform="translate(25.8 3.6) rotate(45)">
      <rect x="-2" y="-1.75" width="1.7" height="0.7"/>
      <rect x="0.3" y="-1.75" width="1.7" height="0.7"/>
      <rect x="-2" y="-0.35" width="4" height="0.7"/>
      <rect x="-2" y="1.05" width="1.7" height="0.7"/>
      <rect x="0.3" y="1.05" width="1.7" height="0.7"/>
    </g>
    <g transform="translate(4.2 16.4) rotate(45)">
      <rect x="-2" y="-1.75" width="4" height="0.7"/>
      <rect x="-2" y="-0.35" width="1.7" height="0.7"/>
      <rect x="0.3" y="-0.35" width="1.7" height="0.7"/>
      <rect x="-2" y="1.05" width="4" height="0.7"/>
    </g>
    <g transform="translate(25.8 16.4) rotate(-45)">
      <rect x="-2" y="-1.75" width="1.7" height="0.7"/>
      <rect x="0.3" y="-1.75" width="1.7" height="0.7"/>
      <rect x="-2" y="-0.35" width="1.7" height="0.7"/>
      <rect x="0.3" y="-0.35" width="1.7" height="0.7"/>
      <rect x="-2" y="1.05" width="1.7" height="0.7"/>
      <rect x="0.3" y="1.05" width="1.7" height="0.7"/>
    </g>
  </g>
</svg>`;

function globalSearchGroup(label, matches, itemRenderer, flagSvg) {
  const isRegistered = matches.length > 0;
  const shown = matches.slice(0, GLOBAL_ITEM_RENDER_LIMIT);

  return `
    <div class="global-search-group">
      <div class="global-search-group-head">
        <h3>${escapeHtml(label)}${flagSvg ? ` ${flagSvg}` : ""}</h3>
        <span class="global-search-status ${isRegistered ? "is-registered" : "is-empty"}">
          ${isRegistered ? `등록됨 (${matches.length}건)` : "등록 안됨"}
        </span>
      </div>
      ${isRegistered ? `<div class="global-search-items">${shown.map(itemRenderer).join("")}</div>` : ""}
      ${matches.length > shown.length ? `<p class="global-search-more">화면에는 ${shown.length}건만 표시됩니다. 전체 ${matches.length}건은 엑셀 추출로 확인하세요.</p>` : ""}
    </div>
  `;
}

function globalMfdsItem(row) {
  return `
    <article class="wc-result-item">
      <p><strong>성분명</strong><span>${escapeHtml(row.ingredient)}</span></p>
      <p><strong>신청인</strong><span>${escapeHtml(row.applicant)}</span></p>
      <p><strong>제조소명</strong><span>${renderMultiList(row.manufacturer, row.country)}</span></p>
      <p><strong>등록일</strong><span>${escapeHtml(row.regDate)}</span></p>
      <p><strong>취소/취하</strong><span>${escapeHtml(row.cancelStatus)}</span></p>
    </article>
  `;
}

function globalCnphItem(row) {
  return `
    <article class="wc-result-item">
      <p><strong>이름</strong><span>${escapeHtml(row.chinese)} / ${escapeHtml(row.korean)} / ${escapeHtml(row.english)}</span></p>
      <p><strong>분류</strong><span>${escapeHtml(row.category)} (${escapeHtml(row.part)})</span></p>
    </article>
  `;
}

function globalWcItem(row) {
  return `
    <article class="wc-result-item">
      <p><strong>원료명</strong><span>${escapeHtml(row.chinese)} / ${escapeHtml(row.korean)} / ${escapeHtml(row.english)}</span></p>
      <p><strong>제조사</strong><span>${escapeHtml(row.manufacturer)}</span></p>
      <p><strong>유효기간</strong><span>${formatValidity(row.validity)}</span></p>
      <p><strong>이메일</strong><span>${escapeHtml(row.email)}</span></p>
      <p><strong>연락처</strong><span>${escapeHtml(row.phone)}</span></p>
    </article>
  `;
}

function bySource(rows, source) {
  return rows.filter((row) => row.source === source);
}

async function renderGlobalSearch(keyword) {
  const token = ++globalSearchToken;
  currentGlobalKeyword = keyword;

  if (!keyword) {
    globalSearchHead.hidden = true;
    globalSearchResults.hidden = true;
    globalSearchResults.innerHTML = "";
    currentGlobalMatches = { wc: [], mfds: [], cnph: [] };
    return;
  }

  globalSearchHead.hidden = false;
  globalSearchResults.hidden = false;
  globalSearchResults.innerHTML = '<p class="empty-result">검색 중입니다.</p>';

  await ensureGlobalSearchData();

  if (token !== globalSearchToken) {
    return;
  }

  const wcMatches = dedupeWcRows(wcRows.filter((row) => row._s.includes(keyword)));
  const mfdsMatches = mfdsRows.filter((row) => row._s.includes(keyword));
  const cnphMatches = cnphRows.filter((row) => row._s.includes(keyword));

  currentGlobalMatches = { wc: wcMatches, mfds: mfdsMatches, cnph: cnphMatches };

  globalSearchResults.innerHTML = [
    globalSearchGroup("K-DMF", mfdsMatches, globalMfdsItem, FLAG_KR_SVG),
    globalSearchGroup("중국 약전", cnphMatches, globalCnphItem, FLAG_CN_SVG),
    globalSearchGroup("WC", bySource(wcMatches, "WC"), globalWcItem, FLAG_CN_SVG),
    globalSearchGroup("COPP", bySource(wcMatches, "COPP"), globalWcItem, FLAG_CN_SVG)
  ].join("");
}

const EXPORT_FONT_NAME = "나눔고딕";
const EXPORT_COLUMN_COUNT = 7;

function styleSectionTitleCell(cell) {
  cell.font = { name: EXPORT_FONT_NAME, size: 13, bold: true, color: { argb: "FFFFFFFF" } };
  cell.alignment = { vertical: "middle", horizontal: "left" };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF3E6EA8" } };
}

function styleHeaderCell(cell) {
  cell.font = { name: EXPORT_FONT_NAME, size: 10, bold: true };
  cell.alignment = { vertical: "middle", wrapText: true };
  cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE9F1FB" } };
  cell.border = {
    top: { style: "thin", color: { argb: "FFD8E0EA" } },
    left: { style: "thin", color: { argb: "FFD8E0EA" } },
    bottom: { style: "thin", color: { argb: "FFD8E0EA" } },
    right: { style: "thin", color: { argb: "FFD8E0EA" } }
  };
}

function styleDataCell(cell) {
  cell.font = { name: EXPORT_FONT_NAME, size: 10 };
  cell.alignment = { vertical: "middle", wrapText: true };
  cell.border = {
    top: { style: "thin", color: { argb: "FFD8E0EA" } },
    left: { style: "thin", color: { argb: "FFD8E0EA" } },
    bottom: { style: "thin", color: { argb: "FFD8E0EA" } },
    right: { style: "thin", color: { argb: "FFD8E0EA" } }
  };
}

function addExportSection(worksheet, label, headers, rows, mapRow) {
  const titleRow = worksheet.addRow([`${label} (${rows.length}건)`]);
  worksheet.mergeCells(titleRow.number, 1, titleRow.number, EXPORT_COLUMN_COUNT);
  titleRow.height = 26;
  titleRow.eachCell({ includeEmpty: true }, styleSectionTitleCell);

  const headerRow = worksheet.addRow(headers);
  headerRow.eachCell(styleHeaderCell);

  if (rows.length === 0) {
    const emptyRow = worksheet.addRow(["등록된 항목이 없습니다."]);
    worksheet.mergeCells(emptyRow.number, 1, emptyRow.number, EXPORT_COLUMN_COUNT);
    emptyRow.eachCell({ includeEmpty: true }, styleDataCell);
  } else {
    rows.forEach((row) => {
      const dataRow = worksheet.addRow(mapRow(row));
      dataRow.eachCell({ includeEmpty: true }, styleDataCell);
    });
  }

  worksheet.addRow([]);
}

async function exportGlobalResults() {
  if (!window.ExcelJS) {
    alert("엑셀 추출 도구를 불러오지 못했습니다.");
    return;
  }

  if (!currentGlobalKeyword) {
    alert("검색어를 입력하세요.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WorkSpace";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("종합검색 결과");
  worksheet.columns = [
    { width: 30 }, { width: 30 }, { width: 34 }, { width: 20 }, { width: 20 }, { width: 22 }, { width: 22 }
  ];

  addExportSection(
    worksheet,
    "K-DMF",
    ["성분명", "신청인", "제조소명", "제조국가", "등록일", "취소/취하구분"],
    currentGlobalMatches.mfds,
    (row) => [
      row.ingredient,
      row.applicant,
      splitMulti(row.manufacturer).map((name, index) => `${index + 1}. ${name}`).join("\n"),
      splitMulti(row.country).join("\n"),
      row.regDate,
      row.cancelStatus
    ]
  );

  addExportSection(
    worksheet,
    "중국 약전",
    ["중국어이름", "한국어이름", "영어이름", "분류", "구분"],
    currentGlobalMatches.cnph,
    (row) => [row.chinese, row.korean, row.english, row.category, row.part]
  );

  const wcColumnHeaders = ["원료명(중국어)", "원료명(한국어)", "원료명(영어)", "제조사", "유효기간", "이메일", "연락처"];
  const wcRowMapper = (row) => [
    row.chinese,
    row.korean,
    row.english,
    row.manufacturer,
    isExpired(row.validity) ? `${row.validity} 만료` : row.validity,
    row.email,
    row.phone
  ];

  addExportSection(worksheet, "WC", wcColumnHeaders, bySource(currentGlobalMatches.wc, "WC"), wcRowMapper);
  addExportSection(worksheet, "COPP", wcColumnHeaders, bySource(currentGlobalMatches.wc, "COPP"), wcRowMapper);

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const dateText = new Date().toISOString().slice(0, 10);

  link.href = url;
  link.download = `종합검색결과_${currentGlobalKeyword}_${dateText}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function showWcSearch(isWcSearch) {
  wcSearchPanel.hidden = !isWcSearch;

  if (isWcSearch) {
    wcSearchInput.value = "";
    wcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  }
}

function showPoReceive(isPoReceive) {
  poReceivePanel.hidden = !isPoReceive;

  if (isPoReceive) {
    poReceiveInput.value = "";
    poReceiveResult.innerHTML = "";
  }
}

function showMfdsSearch(isMfdsSearch) {
  mfdsSearchPanel.hidden = !isMfdsSearch;

  if (isMfdsSearch) {
    mfdsSearchInput.value = "";
    mfdsResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  }
}

async function loadMfdsData() {
  if (mfdsLoaded) {
    return;
  }

  mfdsResults.innerHTML = '<p class="empty-result">데이터를 불러오는 중입니다.</p>';

  try {
    const dataUrl = supabaseConfig.mfdsDataUrl || "data/mfds-review.json";
    const response = await fetch(dataUrl);

    if (!response.ok) {
      throw new Error(`Failed to load MFDS data: ${response.status}`);
    }

    mfdsRows = await response.json();
    mfdsRows.forEach((row) => {
      row._s = row.ingredient.toLowerCase();
    });
    mfdsLoaded = true;
    mfdsResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  } catch (error) {
    console.error(error);
    mfdsResults.innerHTML = '<p class="empty-result">데이터를 불러오지 못했습니다.</p>';
  }
}

function splitMulti(value) {
  return String(value || "")
    .split("@")
    .map((part) => part.trim())
    .filter(Boolean);
}

function renderMultiList(manufacturer, country) {
  const manufacturers = splitMulti(manufacturer);
  const countries = splitMulti(country);

  if (manufacturers.length <= 1) {
    return `${escapeHtml(manufacturer)} (${escapeHtml(country)})`;
  }

  return manufacturers
    .map((name, index) => `${index + 1}. ${escapeHtml(name)} (${escapeHtml(countries[index] || "")})`)
    .join("<br>");
}

function renderMfdsResults() {
  const keyword = mfdsSearchInput.value.trim().toLowerCase();

  if (!keyword) {
    mfdsResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
    return;
  }

  const results = mfdsRows
    .filter((row) => row._s.includes(keyword))
    .slice(0, 80);

  currentMfdsResults = results;

  if (results.length === 0) {
    mfdsResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  mfdsResults.innerHTML = results.map((row) => `
    <article class="wc-result-item">
      <p><strong>성분명</strong><span>${escapeHtml(row.ingredient)}</span></p>
      <p><strong>신청인</strong><span>${escapeHtml(row.applicant)}</span></p>
      <p><strong>제조소명</strong><span>${renderMultiList(row.manufacturer, row.country)}</span></p>
      <p><strong>등록일</strong><span>${escapeHtml(row.regDate)}</span></p>
      <p><strong>취소/취하</strong><span>${escapeHtml(row.cancelStatus)}</span></p>
    </article>
  `).join("");
}

async function exportMfdsResults() {
  if (!window.ExcelJS) {
    alert("엑셀 추출 도구를 불러오지 못했습니다.");
    return;
  }

  if (currentMfdsResults.length === 0) {
    alert("추출할 검색 결과가 없습니다.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WorkSpace";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("성분명 검색결과");
  worksheet.columns = [
    { header: "성분명", key: "ingredient", width: 30 },
    { header: "신청인", key: "applicant", width: 26 },
    { header: "제조소명", key: "manufacturer", width: 44 },
    { header: "제조국가", key: "country", width: 20 },
    { header: "등록일", key: "regDate", width: 16 },
    { header: "취소/취하구분", key: "cancelStatus", width: 16 }
  ];

  currentMfdsResults.forEach((row) => worksheet.addRow({
    ...row,
    manufacturer: splitMulti(row.manufacturer)
      .map((name, index) => `${index + 1}. ${name}`)
      .join("\n"),
    country: splitMulti(row.country).join("\n")
  }));

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = {
        name: EXPORT_FONT_NAME,
        size: rowNumber === 1 ? 11 : 10,
        bold: rowNumber === 1
      };
      cell.alignment = {
        vertical: "middle",
        wrapText: true
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFD8E0EA" } },
        left: { style: "thin", color: { argb: "FFD8E0EA" } },
        bottom: { style: "thin", color: { argb: "FFD8E0EA" } },
        right: { style: "thin", color: { argb: "FFD8E0EA" } }
      };
    });
  });

  worksheet.getRow(1).height = 24;
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE9F1FB" }
  };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const dateText = new Date().toISOString().slice(0, 10);

  link.href = url;
  link.download = `성분명_검색결과_${dateText}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function showCnphSearch(isCnphSearch) {
  cnphSearchPanel.hidden = !isCnphSearch;

  if (isCnphSearch) {
    cnphSearchInput.value = "";
    cnphResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
    cnphCriteria.hidden = true;
  }
}

function showMarginCalc(isMargin) {
  marginCalcPanel.hidden = !isMargin;

  if (isMargin) {
    marginCostInput.value = "";
    marginPriceInput.value = "";
    renderMarginResult();
  }
}

function renderMarginResult() {
  const cost = parseFloat(marginCostInput.value);
  const price = parseFloat(marginPriceInput.value);

  if (!Number.isFinite(cost) || cost <= 0 || !Number.isFinite(price)) {
    marginResult.innerHTML = '<p class="empty-result">예상수입원가와 납품가를 입력하세요.</p>';
    return;
  }

  const marginPercent = ((price - cost) / cost) * 100;
  const isHigh = marginPercent >= 10;
  const sign = marginPercent > 0 ? "+" : "";

  marginResult.innerHTML = `<p class="${isHigh ? "margin-high" : "margin-low"}">마진율 ${sign}${marginPercent.toFixed(2)}%</p>`;
}

function openMarginCalculator() {
  lastFocusedCard = toolsTrigger;
  detailKicker.textContent = "Tools";
  detailTitle.textContent = "마진율 계산기";
  detailDescription.textContent = "예상수입원가와 납품가를 입력해 마진율을 확인하세요.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showImportCostCalc(false);
  showMarginCalc(true);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  marginCostInput.focus();
}

const SHIPPING_RATES = {
  선박: { base: 600000, threshold: 100, step: 100 },
  항공: { base: 300000, threshold: 20, step: 20 }
};

const IMPORT_COST_FLOOR_CAP = 5000000;

function showImportCostCalc(isImportCost) {
  importCostPanel.hidden = !isImportCost;

  if (isImportCost) {
    importPriceInput.value = "";
    importRateInput.value = "";
    importQtyInput.value = "";
    importDutyInput.value = "";
    renderImportCostResult();
  }
}

function formatWonFloor100(value) {
  return Math.floor(value / 100) * 100;
}

function renderImportCostResult() {
  const price = parseFloat(importPriceInput.value);
  const rate = parseFloat(importRateInput.value);
  const qtyRaw = parseFloat(importQtyInput.value);
  const dutyPercent = parseFloat(importDutyInput.value);

  if (!Number.isFinite(price) || price <= 0 || !Number.isFinite(rate) || rate <= 0 || !Number.isFinite(qtyRaw) || qtyRaw <= 0) {
    importCostResult.innerHTML = '<p class="empty-result">단가, 환율, 수량을 입력하세요.</p>';
    return;
  }

  const duty = Number.isFinite(dutyPercent) ? dutyPercent : 0;
  const qtyKg = importQtyUnit.value === "G" ? qtyRaw / 1000 : qtyRaw;

  const baseKrw = price * rate;
  const dutyApplied = baseKrw * (1 + duty / 100);

  const shipping = SHIPPING_RATES[importShippingSelect.value];
  const extraBlocks = Math.ceil(Math.max(0, qtyKg - shipping.threshold) / shipping.step);
  const shippingTotal = shipping.base + extraBlocks * (shipping.base * 0.1);
  const shippingPerKg = shippingTotal / qtyKg;

  const candidate = dutyApplied + shippingPerKg;
  const floor = baseKrw * 1.03;
  const applyFloor = baseKrw <= IMPORT_COST_FLOOR_CAP;
  const finalPrice = applyFloor ? Math.max(candidate, floor) : candidate;
  const rounded = formatWonFloor100(finalPrice);

  importCostResult.innerHTML = `
    <div class="import-cost-breakdown">
      <p>원화단가<span>${Math.round(baseKrw).toLocaleString()}원/kg</span></p>
      <p>관세반영<span>${Math.round(dutyApplied).toLocaleString()}원/kg</span></p>
      <p>운송비<span>${Math.round(shippingPerKg).toLocaleString()}원/kg (총 ${Math.round(shippingTotal).toLocaleString()}원)</span></p>
    </div>
    <p class="import-cost-final">예상수입원가 ${rounded.toLocaleString()}원/kg</p>
  `;
}

function openImportCostCalculator() {
  lastFocusedCard = toolsTrigger;
  detailKicker.textContent = "Tools";
  detailTitle.textContent = "예상수입원가 계산기";
  detailDescription.textContent = "단가, 환율, 수량, 운송방식, 관세를 입력해 kg당 예상수입원가를 확인하세요.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showMarginCalc(false);
  showImportCostCalc(true);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  importPriceInput.focus();
}

function closeToolsDropdown() {
  toolsTrigger.setAttribute("aria-expanded", "false");

  if (document.activeElement && document.activeElement.closest(".tools-menu")) {
    document.activeElement.blur();
  }
}

async function loadCnphData() {
  if (cnphLoaded) {
    return;
  }

  cnphResults.innerHTML = '<p class="empty-result">데이터를 불러오는 중입니다.</p>';

  try {
    const dataUrl = supabaseConfig.cnphDataUrl || "data/china-pharmacopoeia.json";
    const response = await fetch(dataUrl);

    if (!response.ok) {
      throw new Error(`Failed to load China Pharmacopoeia data: ${response.status}`);
    }

    cnphRows = await response.json();
    cnphRows.forEach((row) => {
      row._s = `${row.chinese} ${row.korean} ${row.english}`.toLowerCase();
    });
    cnphLoaded = true;
    cnphResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  } catch (error) {
    console.error(error);
    cnphResults.innerHTML = '<p class="empty-result">데이터를 불러오지 못했습니다.</p>';
  }
}

function renderCnphResults() {
  const keyword = cnphSearchInput.value.trim().toLowerCase();

  if (!keyword) {
    cnphResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
    return;
  }

  const results = cnphRows
    .filter((row) => row._s.includes(keyword))
    .slice(0, 80);

  currentCnphResults = results;

  if (results.length === 0) {
    cnphResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  cnphResults.innerHTML = results.map((row) => `
    <article class="wc-result-item">
      <p><strong>이름</strong><span>${escapeHtml(row.chinese)} / ${escapeHtml(row.korean)} / ${escapeHtml(row.english)}</span></p>
      <p><strong>분류</strong><span>${escapeHtml(row.category)} (${escapeHtml(row.part)})</span></p>
    </article>
  `).join("");
}

async function exportCnphResults() {
  if (!window.ExcelJS) {
    alert("엑셀 추출 도구를 불러오지 못했습니다.");
    return;
  }

  if (currentCnphResults.length === 0) {
    alert("추출할 검색 결과가 없습니다.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WorkSpace";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("중국 약전 검색결과");
  worksheet.columns = [
    { header: "중국어이름", key: "chinese", width: 28 },
    { header: "한국어이름", key: "korean", width: 34 },
    { header: "영어이름", key: "english", width: 38 },
    { header: "분류", key: "category", width: 30 },
    { header: "구분", key: "part", width: 12 }
  ];

  currentCnphResults.forEach((row) => worksheet.addRow(row));

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = {
        name: EXPORT_FONT_NAME,
        size: rowNumber === 1 ? 11 : 10,
        bold: rowNumber === 1
      };
      cell.alignment = {
        vertical: "middle",
        wrapText: true
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFD8E0EA" } },
        left: { style: "thin", color: { argb: "FFD8E0EA" } },
        bottom: { style: "thin", color: { argb: "FFD8E0EA" } },
        right: { style: "thin", color: { argb: "FFD8E0EA" } }
      };
    });
  });

  worksheet.getRow(1).height = 24;
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE9F1FB" }
  };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const dateText = new Date().toISOString().slice(0, 10);

  link.href = url;
  link.download = `중국약전_검색결과_${dateText}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function showUsdmfSearch(isUsdmfSearch) {
  usdmfSearchPanel.hidden = !isUsdmfSearch;

  if (isUsdmfSearch) {
    usdmfSearchInput.value = "";
    usdmfResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  }
}

async function loadUsdmfData() {
  if (usdmfLoaded) {
    return;
  }

  usdmfResults.innerHTML = '<p class="empty-result">데이터를 불러오는 중입니다.</p>';

  try {
    const dataUrl = supabaseConfig.usdmfDataUrl || "data/us-dmf.json";
    const response = await fetch(dataUrl);

    if (!response.ok) {
      throw new Error(`Failed to load US DMF data: ${response.status}`);
    }

    usdmfRows = await response.json();
    usdmfRows.forEach((row) => {
      row._s = `${row.korean} ${row.english}`.toLowerCase();
    });
    usdmfLoaded = true;
    usdmfResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  } catch (error) {
    console.error(error);
    usdmfResults.innerHTML = '<p class="empty-result">데이터를 불러오지 못했습니다.</p>';
  }
}

function renderUsdmfResults() {
  const keyword = usdmfSearchInput.value.trim().toLowerCase();

  if (!keyword) {
    usdmfResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
    return;
  }

  const results = usdmfRows
    .filter((row) => row._s.includes(keyword))
    .slice(0, 80);

  currentUsdmfResults = results;

  if (results.length === 0) {
    usdmfResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  usdmfResults.innerHTML = results.map((row) => `
    <article class="wc-result-item">
      <p><strong>원료명</strong><span>${escapeHtml(row.korean)} / ${escapeHtml(row.english)}</span></p>
      <p><strong>제조사</strong><span>${escapeHtml(row.manufacturer)}</span></p>
      <p><strong>DMF번호</strong><span>${escapeHtml(row.dmfNumber)}</span></p>
      <p><strong>등록일</strong><span>${escapeHtml(row.regDate)}</span></p>
    </article>
  `).join("");
}

async function exportUsdmfResults() {
  if (!window.ExcelJS) {
    alert("엑셀 추출 도구를 불러오지 못했습니다.");
    return;
  }

  if (currentUsdmfResults.length === 0) {
    alert("추출할 검색 결과가 없습니다.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WorkSpace";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("미국 DMF 검색결과");
  worksheet.columns = [
    { header: "한국어명", key: "korean", width: 30 },
    { header: "영어명", key: "english", width: 38 },
    { header: "제조사", key: "manufacturer", width: 34 },
    { header: "DMF번호", key: "dmfNumber", width: 18 },
    { header: "등록일", key: "regDate", width: 16 }
  ];

  currentUsdmfResults.forEach((row) => worksheet.addRow(row));

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = {
        name: EXPORT_FONT_NAME,
        size: rowNumber === 1 ? 11 : 10,
        bold: rowNumber === 1
      };
      cell.alignment = {
        vertical: "middle",
        wrapText: true
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFD8E0EA" } },
        left: { style: "thin", color: { argb: "FFD8E0EA" } },
        bottom: { style: "thin", color: { argb: "FFD8E0EA" } },
        right: { style: "thin", color: { argb: "FFD8E0EA" } }
      };
    });
  });

  worksheet.getRow(1).height = 24;
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE9F1FB" }
  };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const dateText = new Date().toISOString().slice(0, 10);

  link.href = url;
  link.download = `미국DMF_검색결과_${dateText}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function submitPoReceive() {
  const poNo = poReceiveInput.value.trim();

  if (!poNo) {
    poReceiveResult.innerHTML = '<p class="status-error">PO 번호를 입력하세요.</p>';
    return;
  }

  if (!supabaseClient) {
    poReceiveResult.innerHTML = '<p class="status-error">Supabase 연결 설정이 없어 요청을 보낼 수 없습니다.</p>';
    return;
  }

  poReceiveSubmit.disabled = true;
  poReceiveResult.innerHTML = '<p class="empty-result">처리 중입니다.</p>';

  try {
    const { data, error } = await supabaseClient.functions.invoke("po-receive", {
      body: { poNo }
    });

    if (error) {
      throw error;
    }

    if (!data || !data.ok) {
      const message = (data && data.error) || "알 수 없는 오류가 발생했습니다.";
      poReceiveResult.innerHTML = `<p class="status-error">${escapeHtml(message)}</p>`;
      return;
    }

    const ecountNote = data.ecount && data.ecount.loggedIn
      ? `이카운트 연결 확인됨 (${escapeHtml(data.ecount.pending || "")})`
      : `이카운트 연결 실패: ${escapeHtml((data.ecount && data.ecount.error) || "알 수 없음")}`;

    poReceiveResult.innerHTML = `
      <p class="status-ok">
        PO ${escapeHtml(data.poNo)} → shipping DB 입고 처리 완료 (${escapeHtml(data.receivedDate)})<br>
        ${ecountNote}
      </p>
    `;
  } catch (err) {
    console.error(err);
    poReceiveResult.innerHTML = '<p class="status-error">요청 처리 중 오류가 발생했습니다.</p>';
  } finally {
    poReceiveSubmit.disabled = false;
  }
}

async function loadWcData() {
  if (wcLoaded) {
    return;
  }

  wcResults.innerHTML = '<p class="empty-result">데이터를 불러오는 중입니다.</p>';

  try {
    const dataUrl = supabaseConfig.wcDataUrl || "data/wc-copp.json";
    const response = await fetch(dataUrl);

    if (!response.ok) {
      throw new Error(`Failed to load WC data: ${response.status}`);
    }

    wcRows = await response.json();
    wcRows.forEach((row) => {
      row._s = `${row.chinese} ${row.korean} ${row.english}`.toLowerCase();
    });
    wcLoaded = true;
    wcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  } catch (error) {
    console.error(error);
    wcResults.innerHTML = '<p class="empty-result">데이터를 불러오지 못했습니다.</p>';
  }
}

function renderWcSourceGroup(label, rows) {
  const isRegistered = rows.length > 0;
  const shown = rows.slice(0, GLOBAL_ITEM_RENDER_LIMIT);

  return `
    <div class="global-search-group">
      <div class="global-search-group-head">
        <h3>${escapeHtml(label)}</h3>
        <span class="global-search-status ${isRegistered ? "is-registered" : "is-empty"}">
          ${isRegistered ? `등록됨 (${rows.length}건)` : "등록 안됨"}
        </span>
      </div>
      ${isRegistered ? `<div class="wc-source-items">${shown.map(globalWcItem).join("")}</div>` : ""}
      ${rows.length > shown.length ? `<p class="global-search-more">화면에는 ${shown.length}건만 표시됩니다. 전체 ${rows.length}건은 엑셀 추출로 확인하세요.</p>` : ""}
    </div>
  `;
}

function dedupeWcRows(rows) {
  const seen = new Set();

  return rows.filter((row) => {
    const key = [row.chinese, row.korean, row.english, row.manufacturer, row.validity, row.email, row.phone, row.source].join("|");

    if (seen.has(key)) {
      return false;
    }

    seen.add(key);
    return true;
  });
}

function renderWcResults() {
  const keyword = wcSearchInput.value.trim().toLowerCase();

  if (!keyword) {
    wcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
    return;
  }

  const matches = dedupeWcRows(wcRows.filter((row) => row._s.includes(keyword)));
  currentWcResults = matches;

  if (matches.length === 0) {
    wcResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  wcResults.innerHTML = [
    renderWcSourceGroup("WC", bySource(matches, "WC")),
    renderWcSourceGroup("COPP", bySource(matches, "COPP"))
  ].join("");
}

function isExpired(validity) {
  if (!validity) {
    return false;
  }

  const validityDate = new Date(`${validity}T00:00:00`);
  const today = new Date();
  const searchDate = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  return !Number.isNaN(validityDate.getTime()) && validityDate < searchDate;
}

function formatValidity(validity) {
  const safeValidity = escapeHtml(validity);

  if (!isExpired(validity)) {
    return safeValidity;
  }

  return `${safeValidity} <span class="expired-label">만료</span>`;
}

async function exportWcResults() {
  if (!window.ExcelJS) {
    alert("엑셀 추출 도구를 불러오지 못했습니다.");
    return;
  }

  if (currentWcResults.length === 0) {
    alert("추출할 검색 결과가 없습니다.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WorkSpace";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("WC COPP 검색결과");
  worksheet.columns = [
    { header: "구분", key: "source", width: 10 },
    { header: "원료명(중국어)", key: "chinese", width: 28 },
    { header: "원료명(한국어)", key: "korean", width: 34 },
    { header: "원료명(영어)", key: "english", width: 38 },
    { header: "제조사", key: "manufacturer", width: 36 },
    { header: "유효기간", key: "validity", width: 22 },
    { header: "이메일", key: "email", width: 30 },
    { header: "연락처", key: "phone", width: 22 }
  ];

  currentWcResults.forEach((row) => worksheet.addRow({
    ...row,
    validity: isExpired(row.validity) ? `${row.validity} 만료` : row.validity
  }));

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = {
        name: EXPORT_FONT_NAME,
        size: rowNumber === 1 ? 11 : 10,
        bold: rowNumber === 1
      };
      cell.alignment = {
        vertical: "middle",
        wrapText: true
      };
      cell.border = {
        top: { style: "thin", color: { argb: "FFD8E0EA" } },
        left: { style: "thin", color: { argb: "FFD8E0EA" } },
        bottom: { style: "thin", color: { argb: "FFD8E0EA" } },
        right: { style: "thin", color: { argb: "FFD8E0EA" } }
      };
    });

    if (rowNumber > 1 && String(row.getCell("validity").value || "").includes("만료")) {
      row.getCell("validity").font = {
        name: EXPORT_FONT_NAME,
        size: 10,
        color: { argb: "FFD43838" },
        bold: true
      };
    }
  });

  worksheet.getRow(1).height = 24;
  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE9F1FB" }
  };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const dateText = new Date().toISOString().slice(0, 10);

  link.href = url;
  link.download = `WC_COPP_검색결과_${dateText}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function parseWorklogDateFromFilename(filename) {
  const match = String(filename || "").match(/(\d{2,4})[-_](\d{2})[-_](\d{2})/);

  if (!match) {
    return "";
  }

  let [, year, month, day] = match;

  if (year.length === 2) {
    year = `20${year}`;
  }

  return `${year}년 ${month}월 ${day}일자`;
}

function updateWorklogDateDisplay() {
  let label = "불러온 업무일지 없음";

  if (worklogPeople.length > 0) {
    label = worklogDateLabel || "업무일지 불러옴 (날짜 확인 불가)";
  }

  worklogBoxDate.textContent = label;
  worklogDateMeta.textContent = label;
}

async function saveWorklogToStorage() {
  if (!supabaseClient) {
    return;
  }

  try {
    await supabaseClient.functions.invoke("worklog-sync", {
      body: {
        action: "save",
        people: worklogPeople,
        dateLabel: worklogDateLabel
      }
    });
  } catch (error) {
    console.error(error);
  }
}

async function loadWorklogFromStorage() {
  if (!supabaseClient) {
    return false;
  }

  try {
    const { data, error } = await supabaseClient.functions.invoke("worklog-sync", {
      body: { action: "load" }
    });

    if (error || !data || !data.ok || !data.data || !Array.isArray(data.data.people)) {
      return false;
    }

    const saved = data.data;

    worklogPeople = saved.people;
    worklogDateLabel = saved.dateLabel || "";
    updateWorklogDateDisplay();

    return worklogPeople.length > 0;
  } catch (error) {
    console.error(error);
    return false;
  }
}

async function clearWorklogStorage() {
  if (!supabaseClient) {
    return;
  }

  try {
    await supabaseClient.functions.invoke("worklog-sync", {
      body: { action: "clear" }
    });
  } catch (error) {
    console.error(error);
  }
}

async function clearWorklog() {
  if (worklogPeople.length === 0) {
    return;
  }

  if (!confirm("불러온 업무일지를 삭제하시겠습니까? (다른 기기에서도 함께 삭제됩니다)")) {
    return;
  }

  worklogPeople = [];
  worklogDateLabel = "";
  worklogFilterInput.value = "";
  renderWorklog();
  updateWorklogDateDisplay();
  await clearWorklogStorage();
}

async function openWorklog(card) {
  lastFocusedCard = card;
  worklogFullscreen.classList.add("is-open");
  worklogFullscreen.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  worklogFilterInput.value = "";
  worklogResults.innerHTML = '<p class="empty-result">불러오는 중입니다.</p>';
  worklogLoadButton.focus();

  await loadWorklogFromStorage();
  renderWorklog();
}

function closeWorklog() {
  worklogFullscreen.classList.remove("is-open");
  worklogFullscreen.setAttribute("aria-hidden", "true");
  document.body.classList.remove("detail-open");

  if (lastFocusedCard) {
    lastFocusedCard.focus();
  }
}

async function loadWorklogFile(file) {
  worklogResults.innerHTML = '<p class="empty-result">엑셀을 불러오는 중입니다.</p>';

  try {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    worklogPeople = workbook.worksheets
      .map((worksheet, personIndex) => {
        const rows = [];

        worksheet.eachRow((row, rowNumber) => {
          if (rowNumber <= 2) {
            return;
          }

          const v = row.values;
          const company = String(v[1] ?? "").trim();
          const cas = String(v[2] ?? "").trim();
          const material = String(v[3] ?? "").trim();
          const content = String(v[4] ?? "").trim();
          const contact = String(v[5] ?? "").trim();

          if (!company && !material && !content) {
            return;
          }

          rows.push({ id: `${personIndex}-${rowNumber}`, company, cas, material, content, contact });
        });

        return { name: worksheet.name, rows };
      })
      .filter((person) => person.rows.length > 0);

    worklogDateLabel = parseWorklogDateFromFilename(file.name);
    worklogFilterInput.value = "";
    renderWorklog();
    updateWorklogDateDisplay();
    await saveWorklogToStorage();
  } catch (error) {
    console.error(error);
    worklogPeople = [];
    worklogResults.innerHTML = '<p class="empty-result">엑셀을 불러오지 못했습니다. 파일 형식을 확인하세요.</p>';
  }
}

function renderWorklog() {
  if (worklogPeople.length === 0) {
    worklogResults.innerHTML = '<p class="empty-result">엑셀 불러오기 버튼을 누르거나, 파일을 이 화면으로 끌어다 놓으세요.</p>';
    return;
  }

  const keyword = worklogFilterInput.value.trim().toLowerCase();

  const filtered = worklogPeople
    .map((person) => {
      if (!keyword) {
        return person;
      }

      const rows = person.rows.filter((row) => {
        return [person.name, row.company, row.material, row.content, row.contact]
          .join(" ")
          .toLowerCase()
          .includes(keyword);
      });

      return { name: person.name, rows };
    })
    .filter((person) => person.rows.length > 0);

  if (filtered.length === 0) {
    worklogResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  worklogResults.innerHTML = filtered.map((person) => `
    <details class="worklog-person">
      <summary>${escapeHtml(person.name)}<span class="worklog-count">${person.rows.length}건</span></summary>
      <div class="worklog-entries">
        ${person.rows.map((row) => `
          <article class="worklog-entry">
            <div class="worklog-entry-head">
              <span class="worklog-company">${escapeHtml(row.company)}</span>
              <span class="worklog-material">${escapeHtml(row.material)}</span>
              <span class="worklog-entry-meta">
                ${row.cas ? `<span>CAS ${escapeHtml(row.cas)}</span>` : ""}
                ${row.contact ? `<span>${escapeHtml(row.contact)}</span>` : ""}
              </span>
            </div>
            <div class="worklog-content">${escapeHtml(row.content)}</div>
          </article>
        `).join("")}
      </div>
    </details>
  `).join("");
}

wcSearchInput.addEventListener("input", renderWcResults);
excelExportButton.addEventListener("click", exportWcResults);
mfdsSearchInput.addEventListener("input", renderMfdsResults);
mfdsExcelExportButton.addEventListener("click", exportMfdsResults);
cnphSearchInput.addEventListener("input", renderCnphResults);
cnphExcelExportButton.addEventListener("click", exportCnphResults);
usdmfSearchInput.addEventListener("input", renderUsdmfResults);
usdmfExcelExportButton.addEventListener("click", exportUsdmfResults);
cnphCriteriaToggle.addEventListener("click", () => {
  cnphCriteria.hidden = !cnphCriteria.hidden;
});
globalExcelExportButton.addEventListener("click", exportGlobalResults);
poReceiveSubmit.addEventListener("click", submitPoReceive);
poReceiveInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    submitPoReceive();
  }
});
worklogClose.addEventListener("click", closeWorklog);
worklogClearButton.addEventListener("click", clearWorklog);
worklogLoadButton.addEventListener("click", () => worklogFileInput.click());
worklogFileInput.addEventListener("change", () => {
  const file = worklogFileInput.files[0];
  worklogFileInput.value = "";

  if (file) {
    loadWorklogFile(file);
  }
});
worklogFilterInput.addEventListener("input", renderWorklog);

let worklogDragDepth = 0;

worklogDropzone.addEventListener("dragenter", (event) => {
  event.preventDefault();
  worklogDragDepth++;
  worklogFullscreen.classList.add("is-dragover");
});

worklogDropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
});

worklogDropzone.addEventListener("dragleave", () => {
  worklogDragDepth = Math.max(0, worklogDragDepth - 1);

  if (worklogDragDepth === 0) {
    worklogFullscreen.classList.remove("is-dragover");
  }
});

worklogDropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  worklogDragDepth = 0;
  worklogFullscreen.classList.remove("is-dragover");

  const file = event.dataTransfer.files && event.dataTransfer.files[0];

  if (file) {
    loadWorklogFile(file);
  }
});

loadWorklogFromStorage();

if (location.hash === "#worklog") {
  const worklogCard = Array.from(cards).find((card) => card.querySelector("h2").textContent.includes("업무일지"));

  if (worklogCard) {
    openWorklog(worklogCard);
  }
}
