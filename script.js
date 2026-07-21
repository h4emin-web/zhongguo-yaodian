const cards = document.querySelectorAll(".tool-box");
const detailView = document.querySelector(".detail-view");
const detailPanel = document.querySelector(".detail-panel");
const closeButton = document.querySelector(".detail-close");
const topBarVideo = document.querySelector(".top-bar-video");
const marketIndexStrip = document.querySelector(".market-index-strip");
const marketIndexKospi = document.querySelector(".market-index-kospi strong");
const marketIndexKosdaq = document.querySelector(".market-index-kosdaq strong");
const detailKicker = document.querySelector(".detail-kicker");
const detailTitle = document.querySelector("#detail-title");
const detailDescription = document.querySelector(".detail-description");
const searchInput = document.querySelector("#workspace-search");
const workspaceToggle = document.querySelector(".workspace-toggle");
const workspaceContent = document.querySelector("#workspace-content");
const hanaRateCard = document.querySelector(".hana-rate-card");
const hanaRateBase = document.querySelector(".hana-rate-base");
const hanaRateSend = document.querySelector(".hana-rate-send");
const hanaRateUpdated = document.querySelector(".hana-rate-updated");
const hanaRateRefresh = document.querySelector(".hana-rate-refresh");
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
const indiawcSearchPanel = document.querySelector(".indiawc-search-panel");
const indiawcSearchInput = document.querySelector("#indiawc-search-input");
const indiawcResults = document.querySelector(".indiawc-results");
const indiawcExcelExportButton = document.querySelector(".indiawc-excel-export");
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
const marginCostInput = document.querySelector("#margin-cost-input");
const marginPriceInput = document.querySelector("#margin-price-input");
const marginResult = document.querySelector(".margin-calc-result");
const importCostItem = document.querySelector('.tools-item[data-tool="import-cost-calculator"]');
const importCertItem = document.querySelector('.tools-item[data-tool="import-cert-extractor"]');
const importCertPanel = document.querySelector(".importcert-panel");
const importCertDropzone = document.querySelector(".importcert-dropzone");
const importCertFileInput = document.querySelector(".importcert-file-input");
const importCertResult = document.querySelector(".importcert-result");
const autoSettlementItem = document.querySelector('.tools-item[data-tool="auto-settlement"]');
const pendingReceiptItem = document.querySelector('.tools-item[data-tool="pending-receipt"]');
const dailyNewsItem = document.querySelector('.tools-item[data-tool="daily-news"]');
const calendarItem = document.querySelector('.tools-item[data-tool="calendar"]');
const autoSettlementPanel = document.querySelector(".auto-settlement-panel");
const autoSettlementMode = document.querySelector("#auto-settlement-mode");
const autoSettlementManager = document.querySelector("#auto-settlement-manager");
const autoSettlementExchange = document.querySelector("#auto-settlement-exchange");
const autoSettlementQuantity = document.querySelector("#auto-settlement-quantity");
const autoSettlementBoarding = document.querySelector("#auto-settlement-boarding");
const autoSettlementInstock = document.querySelector("#auto-settlement-instock");
const autoSettlementSingleFields = document.querySelectorAll(".auto-single-field");
const autoSettlementBatchCountField = document.querySelector(".auto-batch-count");
const autoSettlementBatchCount = document.querySelector("#auto-settlement-batch-count");
const autoSettlementBatchRatioField = document.querySelector(".auto-batch-ratio");
const autoSettlementBatchRatio = document.querySelector("#auto-settlement-batch-ratio");
const autoSettlementBatchPanel = document.querySelector(".auto-settlement-batch");
const autoSettlementBatchRows = document.querySelector(".auto-settlement-batch-rows");
const autoSettlementFileInput = document.querySelector(".auto-settlement-file");
const autoSettlementUpload = document.querySelector(".auto-settlement-upload");
const autoSettlementSave = document.querySelector(".auto-settlement-save");
const autoSettlementDropzone = document.querySelector(".auto-settlement-dropzone");
const autoExchangeFetch = document.querySelector(".auto-exchange-fetch");
const autoExchangeResult = document.querySelector(".auto-exchange-result");
const autoSettlementResult = document.querySelector(".auto-settlement-result");
const pendingReceiptPanel = document.querySelector(".pending-receipt-panel");
const pendingReceiptMode = document.querySelector("#pending-receipt-mode");
const pendingReceiptInstock = document.querySelector("#pending-receipt-instock");
const pendingReceiptPo = document.querySelector("#pending-receipt-po");
const pendingReceiptSingleField = document.querySelector(".pending-receipt-single-field");
const pendingReceiptBatchCountField = document.querySelector(".pending-receipt-batch-count");
const pendingReceiptBatchCount = document.querySelector("#pending-receipt-batch-count");
const pendingReceiptBatchPanel = document.querySelector(".pending-receipt-batch");
const pendingReceiptBatchRows = document.querySelector(".pending-receipt-batch-rows");
const pendingReceiptRun = document.querySelector(".pending-receipt-run");
const pendingReceiptResult = document.querySelector(".pending-receipt-result");
const dailyNewsPanel = document.querySelector(".daily-news-panel");
const dailyNewsRefresh = document.querySelector(".daily-news-refresh");
const dailyNewsList = document.querySelector(".daily-news-list");
const calendarPanel = document.querySelector(".calendar-panel");
const calendarMonthLabel = document.querySelector(".calendar-month-label");
const calendarPrevButton = document.querySelector(".calendar-prev");
const calendarNextButton = document.querySelector(".calendar-next");
const calendarGrid = document.querySelector(".calendar-grid");
const calendarAddForm = document.querySelector(".calendar-add-form");
const calendarAddDate = document.querySelector("#calendar-add-date");
const calendarAddEndDate = document.querySelector("#calendar-add-end-date");
const calendarAddTitle = document.querySelector("#calendar-add-title");
const calendarColorSwatches = document.querySelector(".calendar-color-swatches");
const calendarAddMemo = document.querySelector("#calendar-add-memo");
const calendarAddSubmit = document.querySelector(".calendar-add-submit");
const calendarAddCancelEdit = document.querySelector(".calendar-add-cancel-edit");
const calendarEventList = document.querySelector(".calendar-event-list");
const calendarSelectedDateLabel = document.querySelector(".calendar-selected-date-label");
const ddayList = document.querySelector(".dday-list");
const calendarDday = document.querySelector(".calendar-dday");
const calendarDdayToggle = document.querySelector(".calendar-dday-toggle");
const calendarStockCard = document.querySelector(".calendar-stock-card");
const calendarStockRefresh = document.querySelector(".calendar-stock-refresh");
const calendarStockPrice = document.querySelector(".calendar-stock-price");
const calendarStockChange = document.querySelector(".calendar-stock-change");
const localLauncherButtons = document.querySelectorAll(".local-launcher-start");
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
const LOCAL_AUTOMATION_BASE = "http://127.0.0.1:4173";
const LOCAL_AUTOMATION_START_URL = "haemin-workspace://start";
const PROTECTED_TOOL_PASSWORD = "1515";
const STAR_BURST_COUNT = 12;
const STAR_BURST_HUE_STEP = 47;
const HANA_RATE_REFRESH_INTERVAL_MS = 60 * 60 * 1000;
const RZNOMICS_STOCK_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const MARKET_INDEX_REFRESH_INTERVAL_MS = 5 * 60 * 1000;
const CALENDAR_STORAGE_KEY = "haemin-workspace-calendar-events";
const EVENT_COLORS = [
  "#FF9EEB",
  "#FFB199",
  "#FFE699",
  "#C2FF99",
  "#99FFD6",
  "#99E6FF",
  "#B399FF",
  "#FF99B8"
];
const DEFAULT_EVENT_COLOR = EVENT_COLORS[0];
const WEEKDAY_LABELS = ["일", "월", "화", "수", "목", "금", "토"];

// Exact dates (including lunar-based holidays and 대체공휴일) for the years
// this calendar is most likely to be used in. Years outside this range fall
// back to FIXED_HOLIDAYS_MMDD below (fixed-date holidays only).
const KR_HOLIDAYS = {
  "2025-01-01": "신정",
  "2025-01-27": "임시공휴일",
  "2025-01-28": "설날 연휴",
  "2025-01-29": "설날",
  "2025-01-30": "설날 연휴",
  "2025-03-01": "삼일절",
  "2025-03-03": "대체공휴일",
  "2025-05-05": "어린이날·부처님오신날",
  "2025-05-06": "대체공휴일",
  "2025-06-06": "현충일",
  "2025-08-15": "광복절",
  "2025-10-03": "개천절",
  "2025-10-05": "추석 연휴",
  "2025-10-06": "추석",
  "2025-10-07": "추석 연휴",
  "2025-10-08": "대체공휴일",
  "2025-10-09": "한글날",
  "2025-12-25": "성탄절",
  "2026-01-01": "신정",
  "2026-02-16": "설날 연휴",
  "2026-02-17": "설날",
  "2026-02-18": "설날 연휴",
  "2026-03-01": "삼일절",
  "2026-03-02": "대체공휴일",
  "2026-05-05": "어린이날",
  "2026-05-24": "부처님오신날",
  "2026-05-25": "대체공휴일",
  "2026-06-06": "현충일",
  "2026-07-17": "제헌절",
  "2026-08-15": "광복절",
  "2026-08-17": "대체공휴일",
  "2026-09-24": "추석 연휴",
  "2026-09-25": "추석",
  "2026-09-26": "추석 연휴",
  "2026-10-03": "개천절",
  "2026-10-05": "대체공휴일",
  "2026-10-09": "한글날",
  "2026-12-25": "성탄절",
  "2027-01-01": "신정",
  "2027-02-06": "설날 연휴",
  "2027-02-07": "설날",
  "2027-02-08": "설날 연휴",
  "2027-02-09": "대체공휴일",
  "2027-03-01": "삼일절",
  "2027-05-05": "어린이날",
  "2027-05-13": "부처님오신날",
  "2027-06-06": "현충일",
  "2027-06-07": "대체공휴일",
  "2027-07-17": "제헌절",
  "2027-08-15": "광복절",
  "2027-08-16": "대체공휴일",
  "2027-09-14": "추석 연휴",
  "2027-09-15": "추석",
  "2027-09-16": "추석 연휴",
  "2027-10-03": "개천절",
  "2027-10-04": "대체공휴일",
  "2027-10-09": "한글날",
  "2027-10-11": "대체공휴일",
  "2027-12-25": "성탄절",
  "2027-12-27": "대체공휴일"
};

// Fallback for years outside KR_HOLIDAYS: fixed-date holidays only (no lunar
// dates, no substitute-holiday adjustments).
const FIXED_HOLIDAYS_MMDD = {
  "01-01": "신정",
  "03-01": "삼일절",
  "05-05": "어린이날",
  "06-06": "현충일",
  "08-15": "광복절",
  "10-03": "개천절",
  "10-09": "한글날",
  "12-25": "성탄절"
};

function getHolidayLabel(dateKey) {
  if (KR_HOLIDAYS[dateKey]) {
    return KR_HOLIDAYS[dateKey];
  }
  const year = Number(dateKey.slice(0, 4));
  if (year >= 2025 && year <= 2027) {
    return null;
  }
  return FIXED_HOLIDAYS_MMDD[dateKey.slice(5)] || null;
}

let lastFocusedCard = null;
let modalLocked = false;
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
let indiawcRows = [];
let indiawcLoaded = false;
let currentIndiawcResults = [];
let globalSearchToken = 0;
let currentGlobalKeyword = "";
let currentGlobalMatches = { wc: [], mfds: [], cnph: [], usdmf: [], indiawc: [] };
let worklogPeople = [];
let worklogDateLabel = "";
let starBurstHue = 0;
let hanaRateTimer = null;
let rznomicsStockTimer = null;
let marketIndexTimer = null;
let workspaceHeartTimer = null;
let dailyNewsLoaded = false;
let dailyNewsLoading = false;
let currentDailyNews = [];
let calendarEvents = loadCalendarEvents();
let calendarViewYear = new Date().getFullYear();
let calendarViewMonth = new Date().getMonth();
let calendarSelectedDate = formatDateKey(new Date());
let editingCalendarEventId = null;
let selectedCalendarColor = DEFAULT_EVENT_COLOR;
let autoSettlementState = {
  mode: "single",
  settlementFile: "",
  poNo: "",
  targetFile: "",
  boardingDate: "",
  instockDate: "",
  targetMonth: "",
  quantity: "",
  exchangeRate: "",
  productCode: "",
  unitPrice: "",
  purchaseUnitPrice: "",
  foreignAmount: "",
  krwAmount: "",
  batchItems: [],
  batchRatioBasis: "quantity"
};
let autoSettlementSelectedFile = null;
let pendingReceiptState = {
  mode: "single",
  poNo: "",
  instockDate: "",
  items: []
};

const AUTO_SETTLEMENT_TARGET_FILES = [
  "1.수입정산서C3-1(26년)-전서진,나지훈,보비.xlsm",
  "2.수입정산서C3-1(26년)-김수빈,이금매(4월~).xlsm",
  "3.수입정산서C3-1(26년)-박성윤,송하형,강해민.xlsm",
  "4.수입정산서C3-1(26년)-엄세희,이승엽,이금매(1~3월).xlsm",
  "5.수입정산서C3-1(26년)-주영웅,민승우,유일환.xlsm",
  "6.수입정산서C3-1(26년)-곽상희, 송예근.xlsm"
];

function parseCssSeconds(value) {
  const raw = String(value || "").split(",")[0].trim();

  if (raw.endsWith("ms")) {
    return Number.parseFloat(raw) / 1000;
  }

  if (raw.endsWith("s")) {
    return Number.parseFloat(raw);
  }

  return Number.parseFloat(raw);
}

function getHeaderDogRunSeconds() {
  if (!topBarVideo) {
    return 10;
  }

  const seconds = parseCssSeconds(window.getComputedStyle(topBarVideo).animationDuration);
  return Number.isFinite(seconds) && seconds > 0 ? seconds : 10;
}

function startHeaderDogVideo() {
  if (!topBarVideo) {
    return;
  }

  topBarVideo.loop = false;

  if (Number.isFinite(topBarVideo.duration) && topBarVideo.duration > 0) {
    const runSeconds = getHeaderDogRunSeconds();
    topBarVideo.defaultPlaybackRate = topBarVideo.duration / runSeconds;
    topBarVideo.playbackRate = topBarVideo.duration / runSeconds;
  }

  try {
    topBarVideo.currentTime = 0;
  } catch (error) {
    console.warn("Header video reset skipped", error);
  }

  topBarVideo.play().catch(() => {});
}

function initializeHeaderDogVideo() {
  if (!topBarVideo) {
    return;
  }

  if (topBarVideo.readyState >= 1) {
    startHeaderDogVideo();
  } else {
    topBarVideo.addEventListener("loadedmetadata", startHeaderDogVideo, { once: true });
  }

  topBarVideo.addEventListener("animationiteration", startHeaderDogVideo);
}

function formatHanaRateValue(value) {
  const numeric = Number(String(value ?? "").replace(/,/g, ""));

  if (!Number.isFinite(numeric)) {
    return "-";
  }

  return new Intl.NumberFormat("ko-KR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(numeric);
}

function formatHanaRateTimestamp(value) {
  const date = value ? new Date(value) : new Date();

  if (Number.isNaN(date.getTime())) {
    return `${String(value || "").replace(/\s+/g, " ").trim()} 기준`;
  }

  const parts = new Intl.DateTimeFormat("ko-KR", {
    timeZone: "Asia/Seoul",
    month: "long",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    hour12: false
  }).formatToParts(date);
  const part = (type) => parts.find((item) => item.type === type)?.value || "";
  const dayLabel = part("day").endsWith("일") ? part("day") : `${part("day")}일`;

  return `${part("month")} ${dayLabel} ${part("hour")}시 ${part("minute")}분 기준`;
}

function renderHanaExchangeRate(data) {
  if (!hanaRateCard || !hanaRateBase || !hanaRateSend || !hanaRateUpdated) {
    return;
  }

  hanaRateCard.classList.remove("is-error");
  hanaRateBase.textContent = formatHanaRateValue(data.baseRate);
  hanaRateSend.textContent = formatHanaRateValue(data.sendRate);
  hanaRateUpdated.textContent = formatHanaRateTimestamp(data.standardAt || data.fetchedAt);
}

function renderHanaExchangeRateError(message) {
  if (!hanaRateCard || !hanaRateBase || !hanaRateSend || !hanaRateUpdated) {
    return;
  }

  hanaRateCard.classList.add("is-error");
  hanaRateBase.textContent = "-";
  hanaRateSend.textContent = "-";
  hanaRateUpdated.textContent = message || "환율 조회 실패";
}

async function fetchHanaExchangeRate({ showLoading = false } = {}) {
  if (!hanaRateCard) {
    return;
  }

  if (showLoading && hanaRateUpdated) {
    hanaRateUpdated.textContent = "조회 중";
    hanaRateCard.classList.remove("is-error");
  }

  if (!supabaseClient) {
    renderHanaExchangeRateError("Supabase 연결 설정 필요");
    return;
  }

  if (hanaRateRefresh) {
    hanaRateRefresh.disabled = true;
  }

  try {
    const { data, error } = await supabaseClient.functions.invoke("hana-exchange-rate", {
      body: { currency: "USD" }
    });

    if (error) {
      throw error;
    }

    if (!data || !data.ok) {
      throw new Error((data && data.error) || "하나은행 환율 조회 실패");
    }

    renderHanaExchangeRate(data);
  } catch (error) {
    console.error(error);
    const message = error && typeof error === "object" && "message" in error
      ? error.message
      : String(error);
    renderHanaExchangeRateError(`환율 조회 실패: ${message}`);
  } finally {
    if (hanaRateRefresh) {
      hanaRateRefresh.disabled = false;
    }
  }
}

function initializeHanaExchangeRate() {
  if (!hanaRateCard) {
    return;
  }

  if (hanaRateRefresh) {
    hanaRateRefresh.addEventListener("click", () => fetchHanaExchangeRate({ showLoading: true }));
  }

  fetchHanaExchangeRate({ showLoading: true });
  hanaRateTimer = window.setInterval(() => {
    fetchHanaExchangeRate();
  }, HANA_RATE_REFRESH_INTERVAL_MS);
}

function formatMarketIndexRate(value) {
  const number = Number(value);

  if (!Number.isFinite(number)) {
    return "-";
  }

  if (number === 0) {
    return "0.00%";
  }

  const sign = number > 0 ? "+" : "-";
  return `${sign}${Math.abs(number).toFixed(2)}%`;
}

function formatMarketIndexValue(value) {
  const number = Number(value);

  if (!Number.isFinite(number) || number <= 0) {
    return "-";
  }

  return new Intl.NumberFormat("ko-KR", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(number);
}

function formatMarketIndexLabel(item) {
  return `${formatMarketIndexValue(item.value)} (${formatMarketIndexRate(item.changeRate)})`;
}

function setMarketIndexItem(element, item) {
  if (!element) {
    return;
  }

  const wrapper = element.closest(".market-index-item");
  wrapper?.classList.remove("is-up", "is-down", "is-flat", "is-error");
  wrapper?.classList.add(item.direction === "up" ? "is-up" : item.direction === "down" ? "is-down" : "is-flat");
  element.textContent = formatMarketIndexLabel(item);
}

function renderMarketIndexTrend(data) {
  if (!marketIndexStrip) {
    return;
  }

  const items = Array.isArray(data.indices) ? data.indices : [];
  const kospi = items.find((item) => item.code === "KOSPI");
  const kosdaq = items.find((item) => item.code === "KOSDAQ");

  setMarketIndexItem(marketIndexKospi, kospi || { changeRate: 0, direction: "flat", value: 0 });
  setMarketIndexItem(marketIndexKosdaq, kosdaq || { changeRate: 0, direction: "flat", value: 0 });
}

function renderMarketIndexError() {
  [marketIndexKospi, marketIndexKosdaq].forEach((element) => {
    if (!element) {
      return;
    }

    const wrapper = element.closest(".market-index-item");
    wrapper?.classList.remove("is-up", "is-down", "is-flat");
    wrapper?.classList.add("is-error");
    element.textContent = "-";
  });
}

async function fetchMarketIndexTrend() {
  if (!marketIndexStrip) {
    return;
  }

  if (!supabaseClient) {
    renderMarketIndexError();
    return;
  }

  try {
    const { data, error } = await supabaseClient.functions.invoke("market-index-trend", {
      body: {}
    });

    if (error) {
      throw error;
    }

    if (!data || !data.ok) {
      throw new Error((data && data.error) || "시장 지수 조회 실패");
    }

    renderMarketIndexTrend(data);
  } catch (error) {
    console.error(error);
    renderMarketIndexError();
  }
}

function initializeMarketIndexTrend() {
  if (!marketIndexStrip) {
    return;
  }

  fetchMarketIndexTrend();
  marketIndexTimer = window.setInterval(() => {
    fetchMarketIndexTrend();
  }, MARKET_INDEX_REFRESH_INTERVAL_MS);
}

function formatKrwStockValue(value) {
  const number = Number(value);

  if (!Number.isFinite(number) || number <= 0) {
    return "-";
  }

  return `${Math.round(number).toLocaleString()}원`;
}

function formatSignedStockChange(value, rate) {
  const change = Number(value);
  const changeRate = Number(rate);

  if (!Number.isFinite(change) || !Number.isFinite(changeRate)) {
    return "-";
  }

  if (change === 0 && changeRate === 0) {
    return "0원 0.00%";
  }

  const sign = change > 0 ? "+" : "-";
  return `${sign}${Math.abs(Math.round(change)).toLocaleString()}원 ${sign}${Math.abs(changeRate).toFixed(2)}%`;
}

function renderRznomicsStockPrice(data) {
  if (!calendarStockCard || !calendarStockPrice || !calendarStockChange) {
    return;
  }

  calendarStockCard.classList.remove("is-error", "is-up", "is-down", "is-flat");
  calendarStockCard.classList.add(data.direction === "up" ? "is-up" : data.direction === "down" ? "is-down" : "is-flat");
  calendarStockPrice.textContent = formatKrwStockValue(data.price);
  calendarStockChange.textContent = formatSignedStockChange(data.change, data.changeRate);
}

function renderRznomicsStockError(message) {
  if (!calendarStockCard || !calendarStockPrice || !calendarStockChange) {
    return;
  }

  calendarStockCard.classList.add("is-error");
  calendarStockCard.classList.remove("is-up", "is-down", "is-flat");
  calendarStockPrice.textContent = "-";
  calendarStockChange.textContent = message || "조회 실패";
}

async function fetchRznomicsStockPrice({ showLoading = false } = {}) {
  if (!calendarStockCard) {
    return;
  }

  if (showLoading) {
    calendarStockCard.classList.remove("is-error");
    calendarStockChange.textContent = "조회 중";
  }

  if (!supabaseClient) {
    renderRznomicsStockError("Supabase 연결 설정 필요");
    return;
  }

  if (calendarStockRefresh) {
    calendarStockRefresh.disabled = true;
  }

  try {
    const { data, error } = await supabaseClient.functions.invoke("rznomics-stock-price", {
      body: {}
    });

    if (error) {
      throw error;
    }

    if (!data || !data.ok) {
      throw new Error((data && data.error) || "알지노믹스 주가 조회 실패");
    }

    renderRznomicsStockPrice(data);
  } catch (error) {
    console.error(error);
    const message = error && typeof error === "object" && "message" in error
      ? error.message
      : String(error);
    renderRznomicsStockError(`주가 조회 실패: ${message}`);
  } finally {
    if (calendarStockRefresh) {
      calendarStockRefresh.disabled = false;
    }
  }
}

function startRznomicsStockUpdates() {
  if (!calendarStockCard) {
    return;
  }

  if (rznomicsStockTimer) {
    window.clearInterval(rznomicsStockTimer);
  }

  fetchRznomicsStockPrice({ showLoading: true });
  rznomicsStockTimer = window.setInterval(() => {
    fetchRznomicsStockPrice();
  }, RZNOMICS_STOCK_REFRESH_INTERVAL_MS);
}

function stopRznomicsStockUpdates() {
  if (!rznomicsStockTimer) {
    return;
  }

  window.clearInterval(rznomicsStockTimer);
  rznomicsStockTimer = null;
}

function setWorkspaceExpanded(isExpanded) {
  if (!workspaceToggle || !workspaceContent) {
    return;
  }

  workspaceToggle.setAttribute("aria-expanded", String(isExpanded));
  workspaceToggle.setAttribute("aria-label", isExpanded ? "작업 박스 접기" : "작업 박스 펼치기");
  workspaceContent.classList.toggle("is-expanded", isExpanded);
  workspaceContent.classList.toggle("is-collapsed", !isExpanded);
  workspaceContent.setAttribute("aria-hidden", String(!isExpanded));
  workspaceContent.inert = !isExpanded;
}

function toggleWorkspaceContent() {
  const isExpanded = workspaceToggle?.getAttribute("aria-expanded") === "true";
  setWorkspaceExpanded(!isExpanded);
  pulseWorkspaceHeart();
}

function pulseWorkspaceHeart() {
  if (!workspaceToggle) {
    return;
  }

  if (workspaceHeartTimer) {
    window.clearTimeout(workspaceHeartTimer);
  }

  workspaceToggle.classList.remove("is-heart-popping");
  void workspaceToggle.offsetWidth;
  workspaceToggle.classList.add("is-heart-popping");
  workspaceHeartTimer = window.setTimeout(() => {
    workspaceToggle.classList.remove("is-heart-popping");
    workspaceHeartTimer = null;
  }, 640);
}

function createStarBurst(event) {
  if (event.button !== undefined && event.button !== 0) {
    return;
  }

  const burst = document.createElement("div");
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const starCount = reducedMotion ? 5 : STAR_BURST_COUNT;
  const hueBase = starBurstHue;
  starBurstHue = (starBurstHue + STAR_BURST_HUE_STEP) % 360;

  burst.className = "star-burst";
  burst.style.left = `${event.clientX}px`;
  burst.style.top = `${event.clientY}px`;

  for (let index = 0; index < starCount; index += 1) {
    const angle = (Math.PI * 2 * index) / starCount + (Math.random() - 0.5) * 0.5;
    const distance = reducedMotion ? 34 + Math.random() * 28 : 56 + Math.random() * 76;
    const size = reducedMotion ? 18 + Math.random() * 8 : 24 + Math.random() * 16;
    const hue = (hueBase + index * (360 / starCount) + Math.random() * 18) % 360;
    const lightness = index % 3 === 0 ? 66 : 74;
    const star = document.createElement("span");

    star.className = "star-burst-star";
    star.style.setProperty("--x", `${Math.cos(angle) * distance}px`);
    star.style.setProperty("--y", `${Math.sin(angle) * distance}px`);
    star.style.setProperty("--size", `${size}px`);
    star.style.setProperty("--rotate", `${Math.round(Math.random() * 160 - 80)}deg`);
    star.style.setProperty("--spark-color", `hsl(${Math.round(hue)} 92% ${lightness}%)`);
    burst.appendChild(star);
  }

  document.body.appendChild(burst);
  burst.addEventListener("animationend", () => burst.remove(), { once: true });
}

function requestProtectedToolAccess(toolLabel, onSuccess) {
  const existingOverlay = document.querySelector(".tool-auth-overlay");
  if (existingOverlay) {
    existingOverlay.querySelector(".tool-auth-input")?.focus();
    return;
  }

  const overlay = document.createElement("div");
  overlay.className = "tool-auth-overlay";
  overlay.setAttribute("role", "dialog");
  overlay.setAttribute("aria-modal", "true");
  overlay.innerHTML = `
    <form class="tool-auth-dialog">
      <button class="tool-auth-close" type="button" aria-label="비밀번호 창 닫기">Close</button>
      <p class="tool-auth-kicker">Protected Tool</p>
      <h2>${toolLabel}</h2>
      <label for="tool-auth-password">비밀번호</label>
      <input id="tool-auth-password" class="tool-auth-input" type="password" inputmode="numeric" autocomplete="off" placeholder="비밀번호 입력">
      <p class="tool-auth-error" aria-live="polite"></p>
      <div class="tool-auth-actions">
        <button class="tool-auth-cancel" type="button">취소</button>
        <button class="tool-auth-submit" type="submit">확인</button>
      </div>
    </form>
  `;

  const form = overlay.querySelector(".tool-auth-dialog");
  const input = overlay.querySelector(".tool-auth-input");
  const error = overlay.querySelector(".tool-auth-error");
  const cancelButton = overlay.querySelector(".tool-auth-cancel");
  const closeAuthButton = overlay.querySelector(".tool-auth-close");
  let closed = false;

  function closeAuth() {
    if (closed) {
      return;
    }

    closed = true;
    document.removeEventListener("keydown", handleAuthKeydown);
    document.body.classList.remove("tool-auth-open");
    overlay.remove();
  }

  function handleAuthKeydown(event) {
    if (event.key === "Escape") {
      closeAuth();
    }
  }

  form.addEventListener("submit", (event) => {
    event.preventDefault();

    if (input.value === PROTECTED_TOOL_PASSWORD) {
      const accessCode = input.value;
      closeAuth();
      onSuccess(accessCode);
      return;
    }

    input.value = "";
    error.textContent = "비밀번호가 맞지 않습니다.";
    form.classList.remove("is-shaking");
    void form.offsetWidth;
    form.classList.add("is-shaking");
    input.focus();
  });

  cancelButton.addEventListener("click", closeAuth);
  closeAuthButton.addEventListener("click", closeAuth);
  overlay.addEventListener("click", (event) => {
    if (event.target === overlay) {
      closeAuth();
    }
  });
  document.addEventListener("keydown", handleAuthKeydown);
  document.body.appendChild(overlay);
  document.body.classList.add("tool-auth-open");
  requestAnimationFrame(() => input.focus());
}

initializeHeaderDogVideo();
document.addEventListener("pointerdown", createStarBurst, { passive: true });
setWorkspaceExpanded(false);

if (workspaceToggle) {
  workspaceToggle.addEventListener("click", toggleWorkspaceContent);
}

function organizeWorkspaceCatalog() {
  const catalog = document.querySelector(".workspace-catalog");

  if (!catalog) {
    return;
  }

  const categories = Array.from(catalog.querySelectorAll(".tool-category"));
  const findCategory = (title) => categories.find((category) => (
    category.querySelector(".tool-category-head h2")?.textContent.trim() === title
  ));
  const automationCategory = findCategory("업무 자동화");
  const infoCategory = findCategory("정보");
  const searchCategory = findCategory("자료 검색 · 기록");

  [automationCategory, infoCategory, searchCategory].forEach((category) => {
    if (category) {
      catalog.appendChild(category);
    }
  });

  if (!searchCategory) {
    return;
  }

  const searchGrid = searchCategory.querySelector(".workspace-grid");
  const automationList = automationCategory?.querySelector(".tool-category-list");
  const searchBoxes = Array.from(searchCategory.querySelectorAll(".tool-box"));
  const findBox = (title) => searchBoxes.find((box) => (
    box.querySelector("h2")?.textContent.trim().includes(title)
  ));
  const worklogBox = findBox("업무일지");

  if (automationList && worklogBox) {
    automationList.appendChild(worklogBox);
  }

  ["K-DMF 검색", "중국 약전", "중국 WC & COPP", "인도 WC", "미국 DMF"].forEach((title) => {
    const box = findBox(title);

    if (searchGrid && box) {
      searchGrid.appendChild(box);
    }
  });
}

organizeWorkspaceCatalog();

function openDetail(card) {
  const title = card.querySelector("h2").textContent;

  if (title.includes("업무일지")) {
    openWorklog(card);
    return;
  }

  const kicker = card.querySelector(".tool-kicker").textContent;
  const description = card.querySelector("p:last-child").textContent;

  lastFocusedCard = card;
  modalLocked = false;
  detailKicker.textContent = kicker;
  detailTitle.textContent = title;
  detailDescription.textContent = description;
  showWcSearch(title.includes("중국 WC & COPP"));
  showPoReceive(title.includes("PO 입고처리"));
  showMfdsSearch(title.includes("K-DMF 검색"));
  showCnphSearch(title.includes("중국 약전"));
  showUsdmfSearch(title.includes("미국 DMF"));
  showIndiawcSearch(title.includes("인도 WC"));
  showImportCostCalc(false);
  showImportCertPanel(false);
  showAutoSettlementPanel(false);
  showPendingReceiptPanel(false);
  showDailyNewsPanel(false);
  showCalendarPanel(false);
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
  } else if (title.includes("인도 WC")) {
    loadIndiawcData();
    indiawcSearchInput.focus();
  } else {
    closeButton.focus();
  }
}

function closeDetail() {
  modalLocked = false;
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
  if (!modalLocked && !detailPanel.contains(event.target)) {
    closeDetail();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !modalLocked && detailView.classList.contains("is-open")) {
    closeDetail();
  }

  if (event.key === "Escape" && worklogFullscreen.classList.contains("is-open")) {
    closeWorklog();
  }

});

marginCostInput.addEventListener("input", renderMarginResult);
marginPriceInput.addEventListener("input", renderMarginResult);

importCostItem.addEventListener("click", () => {
  openImportCostCalculator();
});

[importPriceInput, importPriceCurrency, importRateInput, importQtyInput, importQtyUnit, importShippingSelect, importDutyInput]
  .forEach((el) => {
    el.addEventListener("input", renderImportCostResult);
    el.addEventListener("change", renderImportCostResult);
  });

importCertItem.addEventListener("click", () => {
  openImportCertTool();
});

autoSettlementItem.addEventListener("click", () => {
  requestProtectedToolAccess("자동정산", openAutoSettlementTool);
});

pendingReceiptItem.addEventListener("click", () => {
  requestProtectedToolAccess("입고예정 처리", openPendingReceiptTool);
});

dailyNewsItem.addEventListener("click", openDailyNewsTool);
dailyNewsRefresh.addEventListener("click", () => loadDailyNews({ force: true }));

calendarItem.addEventListener("click", () => {
  requestProtectedToolAccess("캘린더", openCalendarTool);
});

calendarPrevButton.addEventListener("click", () => {
  calendarViewMonth -= 1;
  if (calendarViewMonth < 0) {
    calendarViewMonth = 11;
    calendarViewYear -= 1;
  }
  renderCalendarMonthLabel();
  renderCalendarGrid();
});

calendarNextButton.addEventListener("click", () => {
  calendarViewMonth += 1;
  if (calendarViewMonth > 11) {
    calendarViewMonth = 0;
    calendarViewYear += 1;
  }
  renderCalendarMonthLabel();
  renderCalendarGrid();
});

calendarGrid.addEventListener("click", (event) => {
  const dayButton = event.target.closest(".calendar-day");
  if (!dayButton) {
    return;
  }
  calendarSelectedDate = dayButton.dataset.date;
  resetCalendarAddForm();
  renderCalendarGrid();
  renderCalendarEventList();
});

calendarAddForm.addEventListener("submit", (event) => {
  event.preventDefault();
  let startDate = calendarAddDate.value;
  let endDate = calendarAddEndDate.value || startDate;
  const title = calendarAddTitle.value.trim();
  const color = selectedCalendarColor || DEFAULT_EVENT_COLOR;
  const memo = calendarAddMemo.value.trim();

  if (!startDate || !title) {
    return;
  }

  if (endDate < startDate) {
    [startDate, endDate] = [endDate, startDate];
  }

  calendarSelectedDate = startDate;

  if (editingCalendarEventId) {
    updateCalendarEvent(editingCalendarEventId, {
      date: startDate,
      endDate: endDate !== startDate ? endDate : "",
      title,
      color,
      memo
    });
  } else {
    addCalendarEvent(startDate, endDate, title, color, memo);
  }

  resetCalendarAddForm();
});

calendarAddCancelEdit.addEventListener("click", () => {
  resetCalendarAddForm();
});

calendarDdayToggle.addEventListener("click", () => {
  const isCollapsed = calendarDday.classList.toggle("is-collapsed");
  calendarDdayToggle.setAttribute("aria-expanded", String(!isCollapsed));
});

calendarStockRefresh?.addEventListener("click", () => fetchRznomicsStockPrice({ showLoading: true }));

renderCalendarColorSwatches();
setSelectedCalendarColor(DEFAULT_EVENT_COLOR);

calendarColorSwatches.addEventListener("click", (event) => {
  const swatch = event.target.closest(".calendar-color-swatch");
  if (!swatch) {
    return;
  }
  setSelectedCalendarColor(swatch.dataset.color);
});

function formatDateInputValue(rawValue) {
  const digits = rawValue.replace(/\D/g, "").slice(0, 8);
  const year = digits.slice(0, 4);
  const month = digits.slice(4, 6);
  const day = digits.slice(6, 8);
  return [year, month, day].filter(Boolean).join("-");
}

function attachDateAutoFormat(input) {
  input.addEventListener("focus", () => {
    input.select();
  });
  input.addEventListener("input", () => {
    const wasAtEnd = input.selectionStart === input.value.length;
    input.value = formatDateInputValue(input.value);
    if (wasAtEnd) {
      input.setSelectionRange(input.value.length, input.value.length);
    }
  });
}

attachDateAutoFormat(calendarAddDate);
attachDateAutoFormat(calendarAddEndDate);

calendarEventList.addEventListener("click", (event) => {
  const deleteButton = event.target.closest(".calendar-event-delete");
  if (deleteButton) {
    deleteCalendarEvent(deleteButton.dataset.id);
    return;
  }

  const editButton = event.target.closest(".calendar-event-edit");
  if (editButton) {
    startEditingCalendarEvent(editButton.dataset.id);
  }
});

pendingReceiptMode.addEventListener("change", updatePendingReceiptState);
pendingReceiptInstock.addEventListener("input", updatePendingReceiptState);
pendingReceiptPo.addEventListener("input", updatePendingReceiptState);
pendingReceiptBatchCount.addEventListener("change", () => {
  renderPendingReceiptRows();
  updatePendingReceiptState();
});
pendingReceiptBatchRows.addEventListener("input", updatePendingReceiptState);
pendingReceiptRun.addEventListener("click", runPendingReceiptProcess);
localLauncherButtons.forEach((button) => {
  button.addEventListener("click", startLocalAutomationLauncher);
});

importCertDropzone.addEventListener("click", () => importCertFileInput.click());
importCertDropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    importCertFileInput.click();
  }
});

importCertFileInput.addEventListener("change", () => {
  const file = importCertFileInput.files[0];
  importCertFileInput.value = "";

  if (file) {
    extractLastPdfPage(file);
  }
});

let importCertDragDepth = 0;

importCertDropzone.addEventListener("dragenter", (event) => {
  event.preventDefault();
  importCertDragDepth++;
  importCertDropzone.classList.add("is-dragover");
});

importCertDropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
});

importCertDropzone.addEventListener("dragleave", () => {
  importCertDragDepth = Math.max(0, importCertDragDepth - 1);

  if (importCertDragDepth === 0) {
    importCertDropzone.classList.remove("is-dragover");
  }
});

importCertDropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  importCertDragDepth = 0;
  importCertDropzone.classList.remove("is-dragover");

  const file = event.dataTransfer.files && event.dataTransfer.files[0];

  if (file) {
    extractLastPdfPage(file);
  }
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
  await Promise.all([loadWcData(), loadMfdsData(), loadCnphData(), loadUsdmfData(), loadIndiawcData()]);
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

const FLAG_US_SVG = `<svg class="flag-us" viewBox="0 0 30 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
  <rect y="0.00" width="30" height="2.86" fill="#B22234"/>
  <rect y="2.86" width="30" height="2.86" fill="#FFFFFF"/>
  <rect y="5.71" width="30" height="2.86" fill="#B22234"/>
  <rect y="8.57" width="30" height="2.86" fill="#FFFFFF"/>
  <rect y="11.43" width="30" height="2.86" fill="#B22234"/>
  <rect y="14.29" width="30" height="2.86" fill="#FFFFFF"/>
  <rect y="17.14" width="30" height="2.86" fill="#B22234"/>
  <rect width="12" height="11.43" fill="#3C3B6E"/>
  <circle cx="1.20" cy="1.00" r="0.35" fill="#FFFFFF"/>
  <circle cx="4.40" cy="1.00" r="0.35" fill="#FFFFFF"/>
  <circle cx="7.60" cy="1.00" r="0.35" fill="#FFFFFF"/>
  <circle cx="10.80" cy="1.00" r="0.35" fill="#FFFFFF"/>
  <circle cx="1.20" cy="5.71" r="0.35" fill="#FFFFFF"/>
  <circle cx="4.40" cy="5.71" r="0.35" fill="#FFFFFF"/>
  <circle cx="7.60" cy="5.71" r="0.35" fill="#FFFFFF"/>
  <circle cx="10.80" cy="5.71" r="0.35" fill="#FFFFFF"/>
  <circle cx="1.20" cy="10.43" r="0.35" fill="#FFFFFF"/>
  <circle cx="4.40" cy="10.43" r="0.35" fill="#FFFFFF"/>
  <circle cx="7.60" cy="10.43" r="0.35" fill="#FFFFFF"/>
  <circle cx="10.80" cy="10.43" r="0.35" fill="#FFFFFF"/>
</svg>`;

const FLAG_IN_SVG = `<svg class="flag-in" viewBox="0 0 30 20" xmlns="http://www.w3.org/2000/svg" aria-hidden="true" focusable="false">
  <rect width="30" height="20" fill="#FF9933"/>
  <rect y="6.67" width="30" height="6.67" fill="#FFFFFF"/>
  <rect y="13.33" width="30" height="6.67" fill="#138808"/>
  <circle cx="15" cy="10" r="2.2" fill="none" stroke="#000080" stroke-width="0.3"/>
  <circle cx="15" cy="10" r="0.35" fill="#000080"/>
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

function globalUsdmfItem(row) {
  return `
    <article class="wc-result-item">
      <p><strong>원료명</strong><span>${escapeHtml(row.english)}${row.korean ? ` / ${escapeHtml(row.korean)}` : ""}</span></p>
      <p><strong>제조사</strong><span>${escapeHtml(row.manufacturer)}</span></p>
      <p><strong>DMF번호</strong><span>${escapeHtml(row.dmfNumber)}</span></p>
      <p><strong>구분</strong><span>Type ${escapeHtml(row.type)} · ${escapeHtml(row.status)}</span></p>
      <p><strong>제출일</strong><span>${escapeHtml(row.regDate)}</span></p>
    </article>
  `;
}

function globalIndiawcItem(row) {
  return `
    <article class="wc-result-item">
      <p><strong>제조사</strong><span>${escapeHtml(row.manufacturer)}</span></p>
      <p><strong>품목</strong><span>${escapeHtml(row.english)}</span></p>
    </article>
  `;
}

const NAME_POOL_MIN_LENGTH = 2;

// 실제 데이터의 "영어명"란에는 원료명 대신 이런 제형/행정 용어만 덜렁 들어있는 경우가 많다
// (예: "tablet", "/tablet"). 이런 값은 원료를 특정하지 못하므로 이름 연결에서 제외한다.
const GENERIC_NAME_NOISE = new Set([
  "tablet", "tablets", "/tablet", "/tablets",
  "injection", "injections", "/injection", "/injections",
  "capsule", "capsules", "/capsule", "/capsules",
  "solution", "bulk drug", "api", "not applicable",
  "n/a", "na", "powder", "granule", "granules",
  "suspension", "syrup", "ointment", "cream", "gel",
  "patch", "drops", "spray"
]);

function isUsableName(value) {
  return Boolean(value) && value.length >= NAME_POOL_MIN_LENGTH && !GENERIC_NAME_NOISE.has(value);
}

function addToNamePool(list, seen, rawValue) {
  const value = rawValue.trim();

  if (!isUsableName(value) || seen.has(value)) {
    return;
  }

  seen.add(value);
  list.push(value);
}

function collectNamePools(rowGroups) {
  const chinese = [];
  const korean = [];
  const english = [];
  const seenChinese = new Set();
  const seenKorean = new Set();
  const seenEnglish = new Set();

  rowGroups.forEach((rows) => {
    rows.forEach((row) => {
      if (row.chinese) {
        addToNamePool(chinese, seenChinese, row.chinese);
      }

      if (row.korean) {
        addToNamePool(korean, seenKorean, row.korean);
      }

      if (row.english) {
        addToNamePool(english, seenEnglish, row.english.toLowerCase());
      }
    });
  });

  return { chinese, korean, english };
}

// 우파다시티닙 vs 유파다시티닙처럼, 외래어 표기 시 첫 음절 모음만 다르게 옮긴 같은
// 원료명을 잡기 위한 매칭. 길이가 같고 첫 글자를 뺀 나머지가 완전히 동일할 때만
// 인정한다 (중간이나 끝 글자가 다르면 실제로 다른 원료일 가능성이 높아 제외).
function isFuzzyNameMatch(a, b) {
  if (a.length !== b.length || a.length < 4 || a === b) {
    return false;
  }

  return a.slice(1) === b.slice(1);
}

function sharesName(value, pool) {
  if (!isUsableName(value)) {
    return false;
  }

  return pool.some((entry) => {
    return value.includes(entry) || entry.includes(value) || isFuzzyNameMatch(value, entry);
  });
}

function expandByNamePools(rows, pools) {
  return rows.filter((row) => {
    return sharesName(row.chinese, pools.chinese) ||
      sharesName(row.korean, pools.korean) ||
      (row.english && sharesName(row.english.toLowerCase().trim(), pools.english));
  });
}

function unionByRef(...groups) {
  return Array.from(new Set(groups.flat()));
}

async function renderGlobalSearch(keyword) {
  const token = ++globalSearchToken;
  currentGlobalKeyword = keyword;

  if (!keyword) {
    globalSearchHead.hidden = true;
    globalSearchResults.hidden = true;
    globalSearchResults.innerHTML = "";
    currentGlobalMatches = { wc: [], mfds: [], cnph: [], usdmf: [], indiawc: [] };
    return;
  }

  globalSearchHead.hidden = false;
  globalSearchResults.hidden = false;
  globalSearchResults.innerHTML = '<p class="empty-result">검색 중입니다.</p>';

  await ensureGlobalSearchData();

  if (token !== globalSearchToken) {
    return;
  }

  const wcDirect = wcRows.filter((row) => row._s.includes(keyword));
  const cnphDirect = cnphRows.filter((row) => row._s.includes(keyword));
  const usdmfDirect = usdmfRows.filter((row) => row._s.includes(keyword));
  const indiawcDirect = indiawcRows.filter((row) => row._s.includes(keyword));
  const mfdsDirect = mfdsRows.filter((row) => row._s.includes(keyword));

  const pools = collectNamePools([
    wcDirect,
    cnphDirect,
    usdmfDirect,
    indiawcDirect,
    mfdsDirect.map((row) => ({ korean: row.ingredient }))
  ]);

  const wcExpanded = dedupeWcRows(unionByRef(wcDirect, expandByNamePools(wcRows, pools)));
  const wcMatches = [
    ...latestByManufacturer(bySource(wcExpanded, "WC")),
    ...latestByManufacturer(bySource(wcExpanded, "COPP"))
  ];
  const cnphMatches = unionByRef(cnphDirect, expandByNamePools(cnphRows, pools));
  const usdmfMatches = unionByRef(
    usdmfDirect,
    usdmfRows.filter((row) => row.english && sharesName(row.english.toLowerCase().trim(), pools.english))
  );
  const indiawcMatches = unionByRef(
    indiawcDirect,
    indiawcRows.filter((row) => row.english && sharesName(row.english.toLowerCase().trim(), pools.english))
  );
  const mfdsMatches = unionByRef(
    mfdsDirect,
    mfdsRows.filter((row) => row.ingredient && sharesName(row.ingredient, pools.korean))
  );

  currentGlobalMatches = { wc: wcMatches, mfds: mfdsMatches, cnph: cnphMatches, usdmf: usdmfMatches, indiawc: indiawcMatches };

  globalSearchResults.innerHTML = [
    globalSearchGroup("K-DMF", mfdsMatches, globalMfdsItem, FLAG_KR_SVG),
    globalSearchGroup("중국 약전", cnphMatches, globalCnphItem, FLAG_CN_SVG),
    globalSearchGroup("WC", bySource(wcMatches, "WC"), globalWcItem, FLAG_CN_SVG),
    globalSearchGroup("COPP", bySource(wcMatches, "COPP"), globalWcItem, FLAG_CN_SVG),
    globalSearchGroup("미국 DMF", usdmfMatches, globalUsdmfItem, FLAG_US_SVG),
    globalSearchGroup("인도 WC", indiawcMatches, globalIndiawcItem, FLAG_IN_SVG)
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

  addExportSection(
    worksheet,
    "미국 DMF",
    ["영어명", "한국어명", "제조사", "DMF번호", "TYPE", "STATUS", "제출일"],
    currentGlobalMatches.usdmf,
    (row) => [row.english, row.korean, row.manufacturer, row.dmfNumber, row.type, row.status, row.regDate]
  );

  addExportSection(
    worksheet,
    "인도 WC",
    ["제조사", "품목", "WC번호", "발급일"],
    currentGlobalMatches.indiawc,
    (row) => [row.manufacturer, row.english, row.wcNumber, row.releaseDate]
  );

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
    marginCostInput.value = "";
    marginPriceInput.value = "";
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
    <div class="import-cost-columns">
      <div class="import-cost-breakdown">
        <p>원화단가<span>${Math.round(baseKrw).toLocaleString()}원/kg</span></p>
        <p>관세반영<span>${Math.round(dutyApplied).toLocaleString()}원/kg</span></p>
        <p>운송비<span>${Math.round(shippingPerKg).toLocaleString()}원/kg (총 ${Math.round(shippingTotal).toLocaleString()}원)</span></p>
      </div>
      <div class="import-cost-summary">
        <p>단가 : ${importPriceCurrency.value} ${price}/kg</p>
        <p>예상수입원가 : ${rounded.toLocaleString()}원/kg (@${rate.toLocaleString()})</p>
      </div>
    </div>
    <p class="import-cost-final">예상수입원가 ${rounded.toLocaleString()}원/kg</p>
  `;

  marginCostInput.value = rounded;
  renderMarginResult();
}

function openImportCostCalculator() {
  lastFocusedCard = importCostItem;
  modalLocked = true;
  detailKicker.textContent = "Tools";
  detailTitle.textContent = "예상수입원가 마진율";
  detailDescription.textContent = "단가, 환율, 수량, 운송방식, 관세를 입력해 예상수입원가를 확인하고, 납품가까지 입력하면 마진율도 볼 수 있습니다.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showIndiawcSearch(false);
  showImportCertPanel(false);
  showAutoSettlementPanel(false);
  showPendingReceiptPanel(false);
  showDailyNewsPanel(false);
  showCalendarPanel(false);
  showImportCostCalc(true);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  importPriceInput.focus();
}

function showImportCertPanel(isImportCert) {
  importCertPanel.hidden = !isImportCert;

  if (isImportCert) {
    importCertResult.innerHTML = "";
  }
}

function showAutoSettlementPanel(isAutoSettlement) {
  autoSettlementPanel.hidden = !isAutoSettlement;

  if (isAutoSettlement) {
    renderAutoSettlementResult();
  }
}

function showPendingReceiptPanel(isPendingReceipt) {
  pendingReceiptPanel.hidden = !isPendingReceipt;

  if (isPendingReceipt) {
    renderPendingReceiptRows();
    renderPendingReceiptResult();
  }
}

function showDailyNewsPanel(isDailyNews) {
  dailyNewsPanel.hidden = !isDailyNews;

  if (isDailyNews) {
    loadDailyNews();
  }
}

function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseDateKey(dateKey) {
  const [year, month, day] = dateKey.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function loadCalendarEvents() {
  try {
    const raw = localStorage.getItem(CALENDAR_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.error(error);
    return [];
  }
}

function saveCalendarEvents() {
  localStorage.setItem(CALENDAR_STORAGE_KEY, JSON.stringify(calendarEvents));
}

function renderCalendarMonthLabel() {
  calendarMonthLabel.textContent = `${calendarViewYear}년 ${calendarViewMonth + 1}월`;
}

function getEventEndDate(event) {
  return event.endDate || event.date;
}

function eventCoversDate(event, dateKey) {
  return dateKey >= event.date && dateKey <= getEventEndDate(event);
}

function formatDateLabel(dateKey) {
  const date = parseDateKey(dateKey);
  return `${date.getMonth() + 1}월 ${date.getDate()}일`;
}

function renderCalendarGrid() {
  const firstDayOfWeek = new Date(calendarViewYear, calendarViewMonth, 1).getDay();
  const daysInMonth = new Date(calendarViewYear, calendarViewMonth + 1, 0).getDate();
  const totalCells = Math.ceil((firstDayOfWeek + daysInMonth) / 7) * 7;
  const todayKey = formatDateKey(new Date());
  const eventsByDate = new Map();

  calendarEvents.forEach((event) => {
    const endDate = parseDateKey(getEventEndDate(event));
    const cursor = parseDateKey(event.date);

    while (cursor <= endDate) {
      const key = formatDateKey(cursor);
      if (!eventsByDate.has(key)) {
        eventsByDate.set(key, []);
      }
      eventsByDate.get(key).push(event);
      cursor.setDate(cursor.getDate() + 1);
    }
  });

  const MAX_VISIBLE_CHIPS = 2;
  let cellsHtml = "";

  for (let i = 0; i < totalCells; i += 1) {
    const cellDate = new Date(calendarViewYear, calendarViewMonth, i - firstDayOfWeek + 1);
    const dateKey = formatDateKey(cellDate);
    const isOtherMonth = cellDate.getMonth() !== calendarViewMonth;
    const dayOfWeek = cellDate.getDay();
    const holidayLabel = getHolidayLabel(dateKey);
    const classes = ["calendar-day"];

    if (isOtherMonth) {
      classes.push("is-other-month");
    }
    if (dateKey === todayKey) {
      classes.push("is-today");
    }
    if (dateKey === calendarSelectedDate) {
      classes.push("is-selected");
    }
    if (dayOfWeek === 0 || dayOfWeek === 6 || holidayLabel) {
      classes.push("is-red");
    }

    const dayEvents = eventsByDate.get(dateKey) || [];
    const visibleChips = dayEvents.slice(0, MAX_VISIBLE_CHIPS)
      .map((event) => `<span class="calendar-day-event-chip" style="background:${escapeHtml(event.color || DEFAULT_EVENT_COLOR)}">${escapeHtml(event.title)}</span>`)
      .join("");
    const moreLabel = dayEvents.length > MAX_VISIBLE_CHIPS
      ? `<span class="calendar-day-event-more">+${dayEvents.length - MAX_VISIBLE_CHIPS}</span>`
      : "";

    cellsHtml += `
      <button type="button" class="${classes.join(" ")}" data-date="${dateKey}">
        <span class="calendar-day-number">${cellDate.getDate()}</span>
        ${holidayLabel ? `<span class="calendar-day-holiday">${escapeHtml(holidayLabel)}</span>` : ""}
        <div class="calendar-day-events">${visibleChips}${moreLabel}</div>
      </button>
    `;
  }

  calendarGrid.innerHTML = cellsHtml;
}

function renderCalendarEventList() {
  const date = parseDateKey(calendarSelectedDate);
  const label = `${date.getMonth() + 1}월 ${date.getDate()}일 (${WEEKDAY_LABELS[date.getDay()]}) 일정`;
  calendarSelectedDateLabel.textContent = label;

  const eventsForDay = calendarEvents
    .filter((event) => eventCoversDate(event, calendarSelectedDate))
    .sort((a, b) => a.id.localeCompare(b.id));

  if (eventsForDay.length === 0) {
    calendarEventList.innerHTML = '<p class="empty-result">이 날짜에 등록된 일정이 없습니다.</p>';
    return;
  }

  calendarEventList.innerHTML = eventsForDay.map((event) => {
    const endDate = getEventEndDate(event);
    const isRange = endDate !== event.date;

    return `
      <div class="calendar-event-item">
        <div class="calendar-event-item-info">
          <p class="calendar-event-title">
            <span class="calendar-event-color-dot" style="background:${escapeHtml(event.color || DEFAULT_EVENT_COLOR)}"></span>${escapeHtml(event.title)}
          </p>
          ${isRange ? `<p class="calendar-event-range">${escapeHtml(formatDateLabel(event.date))} ~ ${escapeHtml(formatDateLabel(endDate))}</p>` : ""}
          ${event.memo ? `<p class="calendar-event-memo">${escapeHtml(event.memo)}</p>` : ""}
        </div>
        <div class="calendar-event-item-actions">
          <button class="calendar-event-edit" type="button" data-id="${event.id}">수정</button>
          <button class="calendar-event-delete" type="button" data-id="${event.id}">삭제</button>
        </div>
      </div>
    `;
  }).join("");
}

function renderDdayWidget() {
  if (!ddayList) {
    return;
  }

  const todayKey = formatDateKey(new Date());
  const todayDate = parseDateKey(todayKey);
  const upcomingEvents = calendarEvents
    .filter((event) => getEventEndDate(event) >= todayKey)
    .sort((a, b) => a.date.localeCompare(b.date));

  if (upcomingEvents.length === 0) {
    ddayList.innerHTML = '<li class="dday-empty">등록된 일정이 없습니다.</li>';
    return;
  }

  ddayList.innerHTML = upcomingEvents.map((event) => {
    const eventStartDate = parseDateKey(event.date);
    const diffDays = Math.round((eventStartDate - todayDate) / 86400000);
    let badgeLabel;
    let badgeClass = "dday-item-badge";

    if (diffDays === 0) {
      badgeLabel = "D-DAY";
      badgeClass += " is-today";
    } else if (diffDays > 0) {
      badgeLabel = `D-${diffDays}`;
    } else {
      badgeLabel = "진행중";
      badgeClass += " is-today";
    }

    return `
      <li class="dday-item">
        <span class="dday-item-title">${escapeHtml(event.title)}</span>
        <span class="${badgeClass}">${escapeHtml(badgeLabel)}</span>
      </li>
    `;
  }).join("");
}

function addCalendarEvent(dateKey, endDateKey, title, color, memo) {
  calendarEvents.push({
    id: `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    date: dateKey,
    endDate: endDateKey && endDateKey !== dateKey ? endDateKey : "",
    title,
    color: color || DEFAULT_EVENT_COLOR,
    memo: memo || ""
  });
  saveCalendarEvents();
  renderCalendarGrid();
  renderCalendarEventList();
  renderDdayWidget();
}

function updateCalendarEvent(id, updates) {
  const target = calendarEvents.find((event) => event.id === id);
  if (!target) {
    return;
  }
  Object.assign(target, updates);
  saveCalendarEvents();
  renderCalendarGrid();
  renderCalendarEventList();
  renderDdayWidget();
}

function deleteCalendarEvent(id) {
  calendarEvents = calendarEvents.filter((event) => event.id !== id);
  saveCalendarEvents();
  renderCalendarGrid();
  renderCalendarEventList();
  renderDdayWidget();

  if (editingCalendarEventId === id) {
    resetCalendarAddForm();
  }
}

function renderCalendarColorSwatches() {
  calendarColorSwatches.innerHTML = EVENT_COLORS.map((color) => `
    <button type="button" class="calendar-color-swatch" data-color="${color}" style="background:${color}" aria-label="색상 ${color}"></button>
  `).join("");
}

function setSelectedCalendarColor(color) {
  selectedCalendarColor = color;
  calendarColorSwatches.querySelectorAll(".calendar-color-swatch").forEach((swatch) => {
    swatch.classList.toggle("is-selected", swatch.dataset.color === color);
  });
}

function resetCalendarAddForm() {
  editingCalendarEventId = null;
  calendarAddForm.reset();
  calendarAddDate.value = calendarSelectedDate;
  calendarAddEndDate.value = "";
  setSelectedCalendarColor(DEFAULT_EVENT_COLOR);
  calendarAddMemo.value = "";
  calendarAddSubmit.textContent = "일정 추가";
  calendarAddCancelEdit.hidden = true;
}

function startEditingCalendarEvent(id) {
  const event = calendarEvents.find((item) => item.id === id);
  if (!event) {
    return;
  }

  editingCalendarEventId = id;
  calendarAddDate.value = event.date;
  calendarAddEndDate.value = getEventEndDate(event) !== event.date ? getEventEndDate(event) : "";
  calendarAddTitle.value = event.title;
  setSelectedCalendarColor(event.color || DEFAULT_EVENT_COLOR);
  calendarAddMemo.value = event.memo || "";
  calendarAddSubmit.textContent = "일정 수정";
  calendarAddCancelEdit.hidden = false;
  calendarAddTitle.focus();
}

function showCalendarPanel(isCalendar) {
  calendarPanel.hidden = !isCalendar;

  if (isCalendar) {
    const today = new Date();
    calendarViewYear = today.getFullYear();
    calendarViewMonth = today.getMonth();
    calendarSelectedDate = formatDateKey(today);
    resetCalendarAddForm();
    renderCalendarMonthLabel();
    renderCalendarGrid();
    renderCalendarEventList();
    renderDdayWidget();
  }
}

// "Z26-00575 Mecobalamin 23kg 선적서류.pdf" -> "Z26-00575 Mecobalamin 23kg 수입신고필증.pdf"
// "Z26-00381 Methyl salicylate 16000kg 정산서.pdf" -> "...16000kg 수입신고필증.pdf"
const IMPORTCERT_SUFFIX_PATTERN = /\s*(?:선적서류|선적서|정산서|계산서|명세서)\s*$/;

function deriveImportCertFileName(originalName) {
  const withoutExt = originalName.replace(/\.pdf$/i, "");
  const basePrefix = withoutExt.replace(IMPORTCERT_SUFFIX_PATTERN, "").trim();

  return `${basePrefix} 수입신고필증.pdf`;
}

function downloadPdfBytes(bytes, fileName) {
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function extractLastPdfPage(file) {
  if (!window.PDFLib) {
    alert("PDF 처리 도구를 불러오지 못했습니다.");
    return;
  }

  importCertResult.innerHTML = '<p class="empty-result">처리 중입니다.</p>';

  const fileName = deriveImportCertFileName(file.name);

  try {
    const buffer = await file.arrayBuffer();
    const { PDFDocument } = window.PDFLib;
    const srcDoc = await PDFDocument.load(buffer);
    const pageCount = srcDoc.getPageCount();

    if (pageCount === 0) {
      throw new Error("빈 PDF입니다.");
    }

    const newDoc = await PDFDocument.create();
    const [copiedPage] = await newDoc.copyPages(srcDoc, [pageCount - 1]);
    newDoc.addPage(copiedPage);

    const bytes = await newDoc.save();

    downloadPdfBytes(bytes, fileName);
    importCertResult.innerHTML = `<p class="empty-result">다운로드 완료: "${escapeHtml(fileName)}" (전체 ${pageCount}장 중 마지막 장 추출)</p>`;
  } catch (error) {
    console.error(error);
    importCertResult.innerHTML = '<p class="empty-result">PDF 처리 중 오류가 발생했습니다. 파일을 확인해주세요.</p>';
  }
}

function openImportCertTool() {
  lastFocusedCard = importCertItem;
  modalLocked = false;
  detailKicker.textContent = "Tools";
  detailTitle.textContent = "수입신고필증";
  detailDescription.textContent = "PDF를 올리면 마지막 장만 추출해서 새 PDF로 만들어 드립니다.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showIndiawcSearch(false);
  showImportCostCalc(false);
  showAutoSettlementPanel(false);
  showPendingReceiptPanel(false);
  showDailyNewsPanel(false);
  showCalendarPanel(false);
  showImportCertPanel(true);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  importCertDropzone.focus();
}

function openAutoSettlementTool() {
  lastFocusedCard = autoSettlementItem;
  modalLocked = true;
  detailKicker.textContent = "Tools";
  detailTitle.textContent = "자동정산";
  detailDescription.textContent = "정산서 파일명과 오퍼발행내역의 Boarding/Instock 날짜, 환율을 기준으로 수입정산서와 ERP 구매입력을 처리합니다.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showIndiawcSearch(false);
  showImportCostCalc(false);
  showImportCertPanel(false);
  showPendingReceiptPanel(false);
  showDailyNewsPanel(false);
  showCalendarPanel(false);
  showAutoSettlementPanel(true);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  autoSettlementManager.focus();
}

function openPendingReceiptTool() {
  lastFocusedCard = pendingReceiptItem;
  modalLocked = true;
  detailKicker.textContent = "Tools";
  detailTitle.textContent = "입고예정 처리";
  detailDescription.textContent = "PO 번호와 향남입고일자를 기준으로 오퍼발행내역, 입출고지시서, ERP 구매관리, COA 복사를 처리합니다.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showIndiawcSearch(false);
  showImportCostCalc(false);
  showImportCertPanel(false);
  showAutoSettlementPanel(false);
  showPendingReceiptPanel(true);
  showDailyNewsPanel(false);
  showCalendarPanel(false);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  pendingReceiptPo.focus();
}

function openDailyNewsTool() {
  lastFocusedCard = dailyNewsItem;
  modalLocked = false;
  detailKicker.textContent = "News";
  detailTitle.textContent = "오늘의 뉴스";
  detailDescription.textContent = "데일리팜 메인에 올라온 오늘의 TOP 10 뉴스입니다.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showIndiawcSearch(false);
  showImportCostCalc(false);
  showImportCertPanel(false);
  showAutoSettlementPanel(false);
  showPendingReceiptPanel(false);
  showDailyNewsPanel(true);
  showCalendarPanel(false);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  dailyNewsRefresh.focus();
}

function openCalendarTool() {
  lastFocusedCard = calendarItem;
  modalLocked = true;
  detailKicker.textContent = "Tools";
  detailTitle.textContent = "캘린더";
  detailDescription.textContent = "이번 달 일정을 확인하고 새 일정을 추가할 수 있습니다.";
  showWcSearch(false);
  showPoReceive(false);
  showMfdsSearch(false);
  showCnphSearch(false);
  showUsdmfSearch(false);
  showIndiawcSearch(false);
  showImportCostCalc(false);
  showImportCertPanel(false);
  showAutoSettlementPanel(false);
  showPendingReceiptPanel(false);
  showDailyNewsPanel(false);
  showCalendarPanel(true);
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
}

function renderDailyNewsError(message) {
  dailyNewsList.innerHTML = `<p class="status-error">${escapeHtml(message)}</p>`;
}

function renderDailyNews(data) {
  const news = Array.isArray(data.news) ? data.news : [];
  currentDailyNews = news;

  if (news.length === 0) {
    dailyNewsList.innerHTML = '<p class="empty-result">뉴스를 찾지 못했습니다.</p>';
    return;
  }

  const fetchedAt = data.fetchedAt ? formatHanaRateTimestamp(data.fetchedAt) : "방금 기준";

  dailyNewsList.innerHTML = `
    <p class="daily-news-meta">데일리팜 메인 · ${escapeHtml(fetchedAt)}</p>
    <ol class="daily-news-items">
      ${news.map((item, index) => {
        const rank = Number(item.rank) || index + 1;
        return `
          <li class="daily-news-item">
            <a href="${escapeHtml(item.url)}" target="_blank" rel="noopener noreferrer">
              <span class="daily-news-rank">${rank}</span>
              <span class="daily-news-title">${escapeHtml(item.title)}</span>
            </a>
          </li>
        `;
      }).join("")}
    </ol>
  `;
}

async function loadDailyNews({ force = false } = {}) {
  if (dailyNewsLoading || (!force && dailyNewsLoaded && currentDailyNews.length > 0)) {
    return;
  }

  if (!supabaseClient) {
    renderDailyNewsError("Supabase 연결 설정이 없어 뉴스를 불러올 수 없습니다.");
    return;
  }

  dailyNewsLoading = true;
  dailyNewsRefresh.disabled = true;
  dailyNewsList.innerHTML = '<p class="empty-result">데일리팜 TOP 10 뉴스를 불러오는 중입니다.</p>';

  try {
    const { data, error } = await supabaseClient.functions.invoke("dailypharm-top-news", {
      body: {}
    });

    if (error) {
      throw error;
    }

    if (!data || !data.ok) {
      throw new Error((data && data.error) || "뉴스 조회 실패");
    }

    dailyNewsLoaded = true;
    renderDailyNews(data);
  } catch (error) {
    console.error(error);
    dailyNewsLoaded = false;
    const message = error && typeof error === "object" && "message" in error
      ? error.message
      : String(error);
    renderDailyNewsError(`뉴스 조회 실패: ${message}`);
  } finally {
    dailyNewsLoading = false;
    dailyNewsRefresh.disabled = false;
  }
}

function extractPoNo(text) {
  const match = String(text || "").match(/[A-Z]\d{2}-\d{5}(?:\(\d+\))?/i);
  return match ? match[0].toUpperCase() : "";
}

function extractKgQuantity(text) {
  const match = String(text || "").match(/([\d,]+(?:\.\d+)?)\s*kg/i);
  return match ? match[1].replace(/,/g, "") : "";
}

function normalizeExcelValue(value) {
  if (value == null) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "object") {
    if (value.text) {
      return String(value.text);
    }

    if (value.result != null) {
      return normalizeExcelValue(value.result);
    }

    if (Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text || "").join("");
    }
  }

  return String(value).trim();
}

function getMonthLabelFromDate(value) {
  const text = normalizeExcelValue(value);
  const match = text.match(/(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);

  if (!match) {
    return "";
  }

  return `${Number(match[2])}월`;
}

function findAutoSettlementTargetFile(managerName) {
  const keyword = String(managerName || "").trim();

  if (!keyword) {
    return "";
  }

  return AUTO_SETTLEMENT_TARGET_FILES.find((fileName) => fileName.includes(keyword)) || "";
}

function collectAutoSettlementBatchItems() {
  if (!autoSettlementBatchRows) {
    return [];
  }

  return Array.from(autoSettlementBatchRows.querySelectorAll(".auto-settlement-batch-row")).map((row) => ({
    poNo: row.querySelector('[data-field="poNo"]').value.trim(),
    quantity: row.querySelector('[data-field="quantity"]').value.trim(),
    vat: row.querySelector('[data-field="vat"]').value.trim(),
    duty: row.querySelector('[data-field="duty"]').value.trim()
  }));
}

function resolveAutoSettlementBatchRatioBasis(batchItems = collectAutoSettlementBatchItems()) {
  const selected = autoSettlementBatchRatio?.value || "quantity";

  if (selected === "quantity") {
    return "quantity";
  }

  const hasMissingVat = batchItems.some((item) => !item.vat);
  const allHaveQuantity = batchItems.length > 0 && batchItems.every((item) => item.quantity);

  return hasMissingVat && allHaveQuantity ? "quantity" : "tax";
}

function renderAutoSettlementBatchRows() {
  if (!autoSettlementBatchRows || !autoSettlementBatchCount) {
    return;
  }

  const previous = collectAutoSettlementBatchItems();
  const count = Math.max(2, Math.min(12, Number(autoSettlementBatchCount.value) || 2));
  autoSettlementBatchCount.value = String(count);
  autoSettlementBatchRows.innerHTML = "";

  for (let index = 0; index < count; index += 1) {
    const item = previous[index] || autoSettlementState.batchItems[index] || {};
    const row = document.createElement("div");
    row.className = "auto-settlement-batch-row";
    row.innerHTML = `
      <input data-field="poNo" type="text" placeholder="Z26-00000" autocomplete="off" value="${escapeHtml(item.poNo || "")}">
      <input data-field="quantity" type="number" inputmode="decimal" placeholder="수량" value="${escapeHtml(item.quantity || "")}">
      <input data-field="vat" type="number" inputmode="decimal" placeholder="부가세" value="${escapeHtml(item.vat || "")}">
      <input data-field="duty" type="number" inputmode="decimal" placeholder="관세 없으면 빈칸" value="${escapeHtml(item.duty || "")}">
    `;
    autoSettlementBatchRows.appendChild(row);
  }
}

function syncAutoSettlementMode() {
  const mode = autoSettlementMode?.value || "single";
  autoSettlementState.mode = mode;
  const isBatch = mode === "batch";

  autoSettlementSingleFields.forEach((field) => {
    field.hidden = isBatch;
  });

  if (autoSettlementBatchCountField) {
    autoSettlementBatchCountField.hidden = !isBatch;
  }

  if (autoSettlementBatchRatioField) {
    autoSettlementBatchRatioField.hidden = !isBatch;
  }

  if (autoSettlementBatchPanel) {
    autoSettlementBatchPanel.hidden = !isBatch;
  }

  if (isBatch) {
    const desiredCount = Math.max(2, Math.min(12, Number(autoSettlementBatchCount?.value) || 2));
    const currentCount = autoSettlementBatchRows?.querySelectorAll(".auto-settlement-batch-row").length || 0;

    if (currentCount !== desiredCount) {
      renderAutoSettlementBatchRows();
    }
  }
}

function collectPendingReceiptItems() {
  if (!pendingReceiptBatchRows) {
    return [];
  }

  return Array.from(pendingReceiptBatchRows.querySelectorAll(".pending-receipt-batch-row")).map((row) => ({
    poNo: row.querySelector('[data-field="poNo"]').value.trim()
  })).filter((item) => item.poNo);
}

function renderPendingReceiptRows() {
  if (!pendingReceiptBatchRows || !pendingReceiptBatchCount) {
    return;
  }

  const previous = collectPendingReceiptItems();
  const count = Math.max(2, Math.min(12, Number(pendingReceiptBatchCount.value) || 2));
  pendingReceiptBatchCount.value = String(count);
  pendingReceiptBatchRows.innerHTML = "";

  for (let index = 0; index < count; index += 1) {
    const item = previous[index] || pendingReceiptState.items[index] || {};
    const row = document.createElement("div");
    row.className = "pending-receipt-batch-row";
    row.innerHTML = `<input data-field="poNo" type="text" placeholder="Z26-00000" autocomplete="off" value="${escapeHtml(item.poNo || "")}">`;
    pendingReceiptBatchRows.appendChild(row);
  }
}

function syncPendingReceiptMode() {
  const mode = pendingReceiptMode?.value || "single";
  pendingReceiptState.mode = mode;
  const isBatch = mode === "batch";

  if (pendingReceiptSingleField) {
    pendingReceiptSingleField.hidden = isBatch;
  }

  if (pendingReceiptBatchCountField) {
    pendingReceiptBatchCountField.hidden = !isBatch;
  }

  if (pendingReceiptBatchPanel) {
    pendingReceiptBatchPanel.hidden = !isBatch;
  }

  if (isBatch) {
    const desiredCount = Math.max(2, Math.min(12, Number(pendingReceiptBatchCount?.value) || 2));
    const currentCount = pendingReceiptBatchRows?.querySelectorAll(".pending-receipt-batch-row").length || 0;

    if (currentCount !== desiredCount) {
      renderPendingReceiptRows();
    }
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function isLocalAutomationReady() {
  const response = await fetch(`${LOCAL_AUTOMATION_BASE}/api/health`, {
    method: "GET",
    cache: "no-store"
  });
  return response.ok;
}

function getLocalLauncherStatusElement(button) {
  const panel = button.closest(".wc-search-panel");
  return panel?.querySelector(".auto-settlement-result, .pending-receipt-result") || null;
}

function renderLocalLauncherStatus(statusElement, message, status = "error") {
  if (!statusElement) {
    return;
  }

  statusElement.innerHTML = `<p class="status-${status === "ok" ? "ok" : "error"}">${escapeHtml(message)}</p>`;
}

async function startLocalAutomationLauncher(event) {
  const button = event.currentTarget;
  const statusElement = getLocalLauncherStatusElement(button);
  const originalText = button.textContent;

  button.disabled = true;
  button.textContent = "실행기 확인 중";
  renderLocalLauncherStatus(statusElement, "로컬 자동화 실행기 상태를 확인 중입니다.", "ok");

  try {
    const alreadyReady = await isLocalAutomationReady().catch(() => false);

    if (alreadyReady) {
      renderLocalLauncherStatus(statusElement, "로컬 자동화 실행기가 이미 켜져 있습니다.", "ok");
      return;
    }

    button.textContent = "실행기 여는 중";
    renderLocalLauncherStatus(statusElement, "Windows에서 로컬 자동화 실행기를 여는 중입니다. 권한 확인창이 뜨면 허용해주세요.", "ok");
    window.location.href = LOCAL_AUTOMATION_START_URL;

    for (let attempt = 0; attempt < 15; attempt += 1) {
      await sleep(1000);
      const ready = await isLocalAutomationReady().catch(() => false);

      if (ready) {
        renderLocalLauncherStatus(statusElement, "로컬 자동화 실행기가 켜졌습니다. 이제 다시 실행 버튼을 누르면 됩니다.", "ok");
        return;
      }
    }

    renderLocalLauncherStatus(statusElement, "실행기가 아직 연결되지 않았습니다. 처음 한 번 scripts/install-local-launcher-protocol.ps1을 실행한 뒤 다시 눌러주세요.");
  } finally {
    button.disabled = false;
    button.textContent = originalText;
  }
}

function updatePendingReceiptState() {
  syncPendingReceiptMode();
  pendingReceiptState.instockDate = pendingReceiptInstock.value.trim();

  if (pendingReceiptState.mode === "batch") {
    pendingReceiptState.poNo = "";
    pendingReceiptState.items = collectPendingReceiptItems();
  } else {
    pendingReceiptState.poNo = pendingReceiptPo.value.trim();
    pendingReceiptState.items = pendingReceiptState.poNo ? [{ poNo: pendingReceiptState.poNo }] : [];
  }

  renderPendingReceiptResult();
}

function renderPendingReceiptResult(message = "", status = "error") {
  if (!pendingReceiptResult) {
    return;
  }

  syncPendingReceiptMode();
  const items = pendingReceiptState.mode === "batch" ? collectPendingReceiptItems() : pendingReceiptState.items;
  const rows = [
    ["처리 방식", pendingReceiptState.mode === "batch" ? "일괄" : "단일선적"],
    ["향남입고일자", pendingReceiptState.instockDate || "입력 필요"],
    ["PO 건수", String(items.length || 0)],
    ["처리 범위", "오퍼발행내역 / 입출고지시서 / ERP 구매관리 / COA 복사"]
  ];

  pendingReceiptResult.innerHTML = `
    ${message ? `<p class="status-${status === "ok" ? "ok" : "error"}">${escapeHtml(message)}</p>` : ""}
    <div class="auto-settlement-summary">
      ${rows.map(([label, value]) => `
        <p><strong>${escapeHtml(label)}</strong><span>${escapeHtml(value)}</span></p>
      `).join("")}
    </div>
  `;
}

async function runPendingReceiptProcess() {
  updatePendingReceiptState();

  const isLocal = location.hostname === "localhost" || location.hostname === "127.0.0.1";
  const automationUrl = isLocal
    ? "/api/pending-receipt/process"
    : "http://127.0.0.1:4173/api/pending-receipt/process";
  const items = pendingReceiptState.mode === "batch" ? collectPendingReceiptItems() : pendingReceiptState.items;
  const missing = [
    ["향남입고일자", pendingReceiptState.instockDate]
  ];

  if (!items.length) {
    missing.push(["PO 번호", ""]);
  }

  items.forEach((item, index) => {
    if (!item.poNo) {
      missing.push([`${index + 1}번 PO 번호`, ""]);
    }
  });

  if (missing.filter(([, value]) => !value).length) {
    renderPendingReceiptResult(`${missing.filter(([, value]) => !value).map(([label]) => label).join(", ")} 입력이 필요합니다.`);
    return;
  }

  pendingReceiptRun.disabled = true;
  pendingReceiptResult.innerHTML = '<p class="empty-result">입고예정 처리를 진행 중입니다. Excel과 ERP 자동화가 잠시 걸릴 수 있습니다.</p>';

  try {
    const response = await fetch(automationUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        mode: pendingReceiptState.mode,
        instockDate: pendingReceiptState.instockDate,
        items
      })
    });
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "입고예정 처리에 실패했습니다.");
    }

    const count = Array.isArray(data.items) ? data.items.length : 0;
    renderPendingReceiptResult(`입고예정 처리 완료: ${count}건`, "ok");
    pendingReceiptResult.insertAdjacentHTML("beforeend", `
      <div class="auto-settlement-summary">
        ${(data.items || []).map((item) => `
          <p><strong>${escapeHtml(item.poNo || "")}</strong><span>${escapeHtml(item.productCode || "")} / ${escapeHtml(item.productName || "")} / ${escapeHtml(item.quantity || "")}${escapeHtml(item.unit || "")}</span></p>
        `).join("")}
      </div>
      ${data.warning ? `<p class="status-error">${escapeHtml(data.warning)}</p>` : ""}
    `);
  } catch (error) {
    console.error(error);
    const message = error instanceof TypeError && !isLocal
      ? "로컬 자동화 실행기가 켜져 있어야 합니다. 실행 후 haemin.space에서 다시 시도하세요."
      : error instanceof Error
        ? error.message
        : String(error);
    renderPendingReceiptResult(message);
  } finally {
    pendingReceiptRun.disabled = false;
  }
}

function updateAutoSettlementCalculations() {
  syncAutoSettlementMode();
  const managerName = autoSettlementManager.value.trim();

  autoSettlementState.targetFile = findAutoSettlementTargetFile(managerName);
  autoSettlementState.exchangeRate = autoSettlementExchange.value.trim();
  autoSettlementState.batchItems = collectAutoSettlementBatchItems();
  autoSettlementState.batchRatioBasis = resolveAutoSettlementBatchRatioBasis(autoSettlementState.batchItems);

  if (autoSettlementState.mode === "batch") {
    autoSettlementState.quantity = "";
    autoSettlementState.boardingDate = "";
    autoSettlementState.instockDate = "";
    autoSettlementState.targetMonth = `${autoSettlementState.batchItems.length || 0}건`;
  } else {
    autoSettlementState.quantity = autoSettlementQuantity.value.trim() || autoSettlementState.quantity;
    autoSettlementState.boardingDate = autoSettlementBoarding.value;
    autoSettlementState.instockDate = autoSettlementInstock.value;
    autoSettlementState.targetMonth = getMonthLabelFromDate(autoSettlementInstock.value);
  }

  renderAutoSettlementResult();
}

async function readAutoSettlementFile(file) {
  autoSettlementSelectedFile = file;
  autoSettlementState.settlementFile = file.name;
  autoSettlementState.poNo = extractPoNo(file.name);
  const quantityFromName = extractKgQuantity(file.name);

  if (quantityFromName && !autoSettlementQuantity.value) {
    autoSettlementQuantity.value = quantityFromName;
    autoSettlementState.quantity = quantityFromName;
  }

  updateAutoSettlementCalculations();
}

function handleAutoSettlementFile(file) {
  if (!file) {
    return;
  }

  if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
    renderAutoSettlementResult("정산서 엑셀 파일(.xlsx, .xlsm)을 올려주세요.");
    return;
  }

  readAutoSettlementFile(file);
}

async function saveAutoSettlementToWorkbook() {
  updateAutoSettlementCalculations();

  const isLocal = location.hostname === "localhost" || location.hostname === "127.0.0.1";
  const automationUrl = isLocal
    ? "/api/auto-settlement/save"
    : "http://127.0.0.1:4173/api/auto-settlement/save";

  if (!autoSettlementSelectedFile) {
    renderAutoSettlementResult("정산서 엑셀 파일을 먼저 올려주세요.");
    return;
  }

  const isBatch = autoSettlementState.mode === "batch";
  const batchItems = isBatch ? collectAutoSettlementBatchItems() : [];
  autoSettlementState.batchRatioBasis = resolveAutoSettlementBatchRatioBasis(batchItems);
  const payload = {
    settlementMode: autoSettlementState.mode,
    managerName: autoSettlementManager.value.trim(),
    poNo: autoSettlementState.poNo,
    boardingDate: autoSettlementState.boardingDate,
    instockDate: autoSettlementState.instockDate,
    exchangeRate: autoSettlementState.exchangeRate,
    quantity: autoSettlementState.quantity,
    batchRatioBasis: autoSettlementState.batchRatioBasis,
    batchItems
  };
  const missing = [
    ["담당자 이름", payload.managerName],
    ["ERP 환율", payload.exchangeRate]
  ].filter(([, value]) => !value);

  if (isBatch) {
    if (!batchItems.length) {
      missing.push(["일괄 항목", ""]);
    }

    batchItems.forEach((item, index) => {
      if (!item.poNo) {
        missing.push([`${index + 1}번 PO 번호`, ""]);
      }
      if (!item.quantity) {
        missing.push([`${index + 1}번 수량`, ""]);
      }
      if (payload.batchRatioBasis !== "quantity" && !item.vat) {
        missing.push([`${index + 1}번 부가세`, ""]);
      }
    });
  } else if (!payload.poNo) {
    missing.push(["PO 번호", ""]);
  }

  if (missing.length) {
    renderAutoSettlementResult(`${missing.map(([label]) => label).join(", ")} 입력이 필요합니다.`);
    return;
  }

  const formData = new FormData();
  formData.append("settlementFile", autoSettlementSelectedFile, autoSettlementSelectedFile.name);
  Object.entries(payload).forEach(([key, value]) => {
    formData.append(key, Array.isArray(value) ? JSON.stringify(value) : value);
  });

  autoSettlementSave.disabled = true;
  autoSettlementResult.innerHTML = '<p class="empty-result">수입정산서 저장 후 입출고 지시서와 ERP 구매 저장을 진행 중입니다. ERP 저장은 잠시 걸릴 수 있습니다.</p>';

  try {
    const response = await fetch(automationUrl, {
      method: "POST",
      body: formData
    });
    const data = await response.json();

    if (!response.ok || !data.ok) {
      throw new Error(data.error || "수입정산서 저장에 실패했습니다.");
    }

    const firstItem = Array.isArray(data.items) ? data.items[0] : data;
    autoSettlementState.targetFile = data.targetFile || firstItem.targetFile || autoSettlementState.targetFile;
    autoSettlementState.targetMonth = data.sheetName || firstItem.sheetName || autoSettlementState.targetMonth;
    autoSettlementState.boardingDate = firstItem.boardingDate || autoSettlementState.boardingDate;
    autoSettlementState.instockDate = firstItem.instockDate || autoSettlementState.instockDate;
    autoSettlementBoarding.value = autoSettlementState.boardingDate;
    autoSettlementInstock.value = autoSettlementState.instockDate;
    autoSettlementState.productCode = firstItem.productCode || autoSettlementState.productCode;
    autoSettlementState.quantity = firstItem.quantity || autoSettlementState.quantity;
    autoSettlementState.exchangeRate = data.exchangeRate || autoSettlementState.exchangeRate;
    autoSettlementState.unitPrice = firstItem.unitPrice || autoSettlementState.unitPrice;
    autoSettlementState.purchaseUnitPrice = firstItem.purchaseUnitPrice || autoSettlementState.purchaseUnitPrice;
    autoSettlementState.foreignAmount = firstItem.foreignAmount || autoSettlementState.foreignAmount;
    autoSettlementState.krwAmount = firstItem.krwAmount || autoSettlementState.krwAmount;
    renderAutoSettlementResult();
    const ecountList = Array.isArray(data.ecount) ? data.ecount : (data.ecount ? [data.ecount] : []);
    const inoutList = Array.isArray(data.inout) ? data.inout : (data.inout ? [data.inout] : []);
    const ecountMessage = ecountList.length
      ? ` / ERP 구매 ${ecountList.length}건 처리 완료`
      : "";
    const inoutMessage = inoutList.length
      ? ` / 입출고 지시서 ${inoutList.length}건 반영 완료`
      : "";
    const savedLabel = Array.isArray(data.items)
      ? `${escapeHtml(data.targetFile)} / ${escapeHtml(String(data.items.length))}건`
      : `${escapeHtml(data.targetFile)} / ${escapeHtml(data.sheetName)} ${escapeHtml(String(data.startRow))}행`;
    autoSettlementResult.insertAdjacentHTML("afterbegin", `
      <p class="status-ok">원본 저장 완료: ${savedLabel}${ecountMessage}${inoutMessage}</p>
      ${data.warning ? `<p class="status-error">${escapeHtml(data.warning)}</p>` : ""}
    `);
  } catch (error) {
    console.error(error);
    const message = error instanceof TypeError && !isLocal
      ? "로컬 자동정산 실행기가 켜져 있어야 합니다. 실행 후 haemin.space에서 다시 업로드하세요."
      : error instanceof Error
        ? error.message
        : String(error);
    renderAutoSettlementResult(message);
  } finally {
    autoSettlementSave.disabled = false;
  }
}

function renderAutoSettlementResult(message = "") {
  if (!autoSettlementResult) {
    return;
  }

  const isBatch = autoSettlementState.mode === "batch";
  const batchItems = isBatch ? collectAutoSettlementBatchItems() : [];
  const rows = isBatch ? [
    ["담당자 파일", autoSettlementState.targetFile || "담당자 이름 입력 후 확인"],
    ["처리 방식", "일괄"],
    ["비율 기준", resolveAutoSettlementBatchRatioBasis(batchItems) === "quantity" ? "수량" : "부가세/관세"],
    ["일괄 건수", String(batchItems.length || 0)],
    ["ERP 환율", autoSettlementState.exchangeRate || "수동 입력 필요"]
  ] : [
    ["담당자 파일", autoSettlementState.targetFile || "담당자 이름 입력 후 확인"],
    ["PO 번호", autoSettlementState.poNo || "정산서 파일명에서 추출 예정"],
    ["선적일자", autoSettlementState.boardingDate || "오퍼발행내역 자동 조회"],
    ["입고일자", autoSettlementState.instockDate || "오퍼발행내역 자동 조회"],
    ["작성 시트", autoSettlementState.targetMonth || "입고일자 월 기준"],
    ["ERP 환율", autoSettlementState.exchangeRate || "수동 입력 필요"],
    ["수량", autoSettlementState.quantity || "정산서 파일명 또는 수동 입력"],
    ["품목코드", autoSettlementState.productCode || "업로드 후 확인"],
    ["ERP 단가", autoSettlementState.unitPrice || "업로드 후 확인"],
    ["외화 단가", autoSettlementState.purchaseUnitPrice || "업로드 후 확인"],
    ["외화금액", autoSettlementState.foreignAmount || "업로드 후 확인"],
    ["원화금액", autoSettlementState.krwAmount || "업로드 후 확인"]
  ];

  autoSettlementResult.innerHTML = `
    ${message ? `<p class="status-error">${escapeHtml(message)}</p>` : ""}
    <div class="auto-settlement-summary">
      ${rows.map(([label, value]) => `
        <p><strong>${escapeHtml(label)}</strong><span>${escapeHtml(value)}</span></p>
      `).join("")}
    </div>
    <p class="wc-search-meta">선적일자와 입고일자는 오퍼발행내역에서 PO 번호로 자동 조회합니다. 일괄은 PO별 부가세와 관세 입력값으로 비율을 계산합니다.</p>
  `;
}

async function exportAutoSettlementSummary() {
  if (!window.ExcelJS) {
    alert("엑셀 처리 도구를 불러오지 못했습니다.");
    return;
  }

  updateAutoSettlementCalculations();

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WorkSpace";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet("자동정산 입력값");

  worksheet.columns = [
    { header: "항목", key: "label", width: 24 },
    { header: "값", key: "value", width: 56 }
  ];

  [
    ["담당자", autoSettlementManager.value.trim()],
    ["담당자 파일", autoSettlementState.targetFile],
    ["정산서 파일", autoSettlementState.settlementFile],
    ["PO 번호", autoSettlementState.poNo],
    ["선적일자", autoSettlementState.boardingDate],
    ["입고일자", autoSettlementState.instockDate],
    ["작성 시트", autoSettlementState.targetMonth],
    ["ERP 환율", autoSettlementState.exchangeRate],
    ["수량", autoSettlementState.quantity]
  ].forEach(([label, value]) => worksheet.addRow({ label, value }));

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = { name: EXPORT_FONT_NAME, size: rowNumber === 1 ? 11 : 10, bold: rowNumber === 1 };
      cell.alignment = { vertical: "middle", wrapText: true };
      cell.border = {
        top: { style: "thin", color: { argb: "FFD8E0EA" } },
        left: { style: "thin", color: { argb: "FFD8E0EA" } },
        bottom: { style: "thin", color: { argb: "FFD8E0EA" } },
        right: { style: "thin", color: { argb: "FFD8E0EA" } }
      };
    });
  });

  worksheet.getRow(1).fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFE9F1FB" }
  };

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = `자동정산_${autoSettlementState.poNo || "입력값"}.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function fetchAutoSettlementExchangeRate() {
  const poNo = autoSettlementState.poNo || extractPoNo(autoSettlementState.settlementFile);

  if (!poNo) {
    autoExchangeResult.innerHTML = '<p class="status-error">먼저 정산서 파일을 올려 PO 번호를 확인해주세요.</p>';
    return;
  }

  autoExchangeFetch.disabled = true;
  autoExchangeResult.innerHTML = '<p class="empty-result">ERP에서 환율을 조회하는 중입니다.</p>';

  try {
    let data;
    let error;
    const isLocal = location.hostname === "localhost" || location.hostname === "127.0.0.1";

    if (isLocal) {
      const response = await fetch("/api/ecount-rate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ poNo })
      });

      data = await response.json();
      error = response.ok ? null : data;
    } else {
      if (!supabaseClient) {
        autoExchangeResult.innerHTML = '<p class="status-error">Supabase 연결 설정이 없어 ERP 조회를 할 수 없습니다.</p>';
        return;
      }

      const result = await supabaseClient.functions.invoke("ecount-exchange-rate", {
        body: { poNo }
      });

      data = result.data;
      error = result.error;
    }

    if (error) {
      throw error;
    }

    if (!data || !data.ok) {
      autoExchangeResult.innerHTML = `<p class="status-error">${escapeHtml((data && data.error) || "ERP 환율 조회에 실패했습니다.")}</p>`;
      return;
    }

    autoSettlementExchange.value = data.exchangeRate;
    updateAutoSettlementCalculations();
    autoExchangeResult.innerHTML = `<p class="status-ok">ERP 환율 ${escapeHtml(data.exchangeRate)} 적용 완료</p>`;
  } catch (error) {
    console.error(error);
    const message = error && typeof error === "object" && "error" in error
      ? error.error
      : error instanceof Error
        ? error.message
        : String(error);
    autoExchangeResult.innerHTML = `<p class="status-error">ERP 환율 조회 실패: ${escapeHtml(message)}</p>`;
  } finally {
    autoExchangeFetch.disabled = false;
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
      row._s = row.english.toLowerCase();
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
      <p><strong>원료명</strong><span>${escapeHtml(row.english)}${row.korean ? ` / ${escapeHtml(row.korean)}` : ""}</span></p>
      <p><strong>제조사</strong><span>${escapeHtml(row.manufacturer)}</span></p>
      <p><strong>DMF번호</strong><span>${escapeHtml(row.dmfNumber)}</span></p>
      <p><strong>구분</strong><span>Type ${escapeHtml(row.type)} · ${escapeHtml(row.status)}</span></p>
      <p><strong>제출일</strong><span>${escapeHtml(row.regDate)}</span></p>
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
    { header: "영어명", key: "english", width: 38 },
    { header: "한국어명", key: "korean", width: 30 },
    { header: "제조사", key: "manufacturer", width: 34 },
    { header: "DMF번호", key: "dmfNumber", width: 18 },
    { header: "TYPE", key: "type", width: 10 },
    { header: "STATUS", key: "status", width: 10 },
    { header: "제출일", key: "regDate", width: 16 }
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

function showIndiawcSearch(isIndiawcSearch) {
  indiawcSearchPanel.hidden = !isIndiawcSearch;

  if (isIndiawcSearch) {
    indiawcSearchInput.value = "";
    indiawcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  }
}

async function loadIndiawcData() {
  if (indiawcLoaded) {
    return;
  }

  indiawcResults.innerHTML = '<p class="empty-result">데이터를 불러오는 중입니다.</p>';

  try {
    const dataUrl = supabaseConfig.indiaWcDataUrl || "data/india-wc.json";
    const response = await fetch(dataUrl);

    if (!response.ok) {
      throw new Error(`Failed to load India WC data: ${response.status}`);
    }

    indiawcRows = await response.json();
    indiawcRows.forEach((row) => {
      row._s = row.english.toLowerCase();
    });
    indiawcLoaded = true;
    indiawcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  } catch (error) {
    console.error(error);
    indiawcResults.innerHTML = '<p class="empty-result">데이터를 불러오지 못했습니다.</p>';
  }
}

function renderIndiawcResults() {
  const keyword = indiawcSearchInput.value.trim().toLowerCase();

  if (!keyword) {
    indiawcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
    return;
  }

  const results = indiawcRows
    .filter((row) => row._s.includes(keyword))
    .slice(0, 80);

  currentIndiawcResults = results;

  if (results.length === 0) {
    indiawcResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  indiawcResults.innerHTML = results.map((row) => `
    <article class="wc-result-item">
      <p><strong>제조사</strong><span>${escapeHtml(row.manufacturer)}</span></p>
      <p><strong>품목</strong><span>${escapeHtml(row.english)}</span></p>
    </article>
  `).join("");
}

async function exportIndiawcResults() {
  if (!window.ExcelJS) {
    alert("엑셀 추출 도구를 불러오지 못했습니다.");
    return;
  }

  if (currentIndiawcResults.length === 0) {
    alert("추출할 검색 결과가 없습니다.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "WorkSpace";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("인도 WC 검색결과");
  worksheet.columns = [
    { header: "제조사", key: "manufacturer", width: 40 },
    { header: "품목", key: "english", width: 42 },
    { header: "WC번호", key: "wcNumber", width: 16 },
    { header: "발급일", key: "releaseDate", width: 16 }
  ];

  currentIndiawcResults.forEach((row) => worksheet.addRow(row));

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
  link.download = `인도WC_검색결과_${dateText}.xlsx`;
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

function parseValidityDate(value) {
  if (!value) {
    return null;
  }

  const cleaned = String(value).replace(/(\d+)(st|nd|rd|th)/gi, "$1");
  const parsed = new Date(cleaned);

  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function normalizeForKey(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9一-鿿가-힣]/g, "");
}

function latestByManufacturer(rows) {
  const bestByKey = new Map();

  rows.forEach((row) => {
    const key = [row.chinese, row.korean, row.english, row.manufacturer].map(normalizeForKey).join("|");
    const existing = bestByKey.get(key);

    if (!existing) {
      bestByKey.set(key, row);
      return;
    }

    const existingDate = parseValidityDate(existing.validity);
    const rowDate = parseValidityDate(row.validity);

    if (rowDate && (!existingDate || rowDate > existingDate)) {
      bestByKey.set(key, row);
    }
  });

  return Array.from(bestByKey.values());
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
  const wcOnly = latestByManufacturer(bySource(matches, "WC"));
  const coppOnly = latestByManufacturer(bySource(matches, "COPP"));
  currentWcResults = [...wcOnly, ...coppOnly];

  if (currentWcResults.length === 0) {
    wcResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  wcResults.innerHTML = [
    renderWcSourceGroup("WC", wcOnly),
    renderWcSourceGroup("COPP", coppOnly)
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
indiawcSearchInput.addEventListener("input", renderIndiawcResults);
indiawcExcelExportButton.addEventListener("click", exportIndiawcResults);
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
autoSettlementManager.addEventListener("input", updateAutoSettlementCalculations);
autoSettlementMode.addEventListener("change", updateAutoSettlementCalculations);
autoSettlementExchange.addEventListener("input", updateAutoSettlementCalculations);
autoSettlementQuantity.addEventListener("input", updateAutoSettlementCalculations);
autoSettlementBoarding.addEventListener("input", updateAutoSettlementCalculations);
autoSettlementInstock.addEventListener("input", updateAutoSettlementCalculations);
autoSettlementBatchRatio.addEventListener("change", updateAutoSettlementCalculations);
function handleAutoSettlementBatchCountChange() {
  renderAutoSettlementBatchRows();
  updateAutoSettlementCalculations();
}

autoSettlementBatchCount.addEventListener("input", handleAutoSettlementBatchCountChange);
autoSettlementBatchCount.addEventListener("change", handleAutoSettlementBatchCountChange);
autoSettlementBatchRows.addEventListener("input", updateAutoSettlementCalculations);
autoSettlementUpload.addEventListener("click", () => autoSettlementFileInput.click());
autoExchangeFetch.addEventListener("click", fetchAutoSettlementExchangeRate);
autoSettlementSave.addEventListener("click", saveAutoSettlementToWorkbook);
autoSettlementDropzone.addEventListener("click", () => autoSettlementFileInput.click());
autoSettlementDropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    autoSettlementFileInput.click();
  }
});
autoSettlementDropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  autoSettlementDropzone.classList.add("is-dragover");
});
autoSettlementDropzone.addEventListener("dragleave", () => {
  autoSettlementDropzone.classList.remove("is-dragover");
});
autoSettlementDropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  autoSettlementDropzone.classList.remove("is-dragover");
  handleAutoSettlementFile(event.dataTransfer.files && event.dataTransfer.files[0]);
});
autoSettlementFileInput.addEventListener("change", () => {
  const file = autoSettlementFileInput.files[0];
  autoSettlementFileInput.value = "";

  handleAutoSettlementFile(file);
});
syncAutoSettlementMode();
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

initializeHanaExchangeRate();
initializeMarketIndexTrend();
startRznomicsStockUpdates();
loadWorklogFromStorage();

if (location.hash === "#worklog") {
  const worklogCard = Array.from(cards).find((card) => card.querySelector("h2").textContent.includes("업무일지"));

  if (worklogCard) {
    openWorklog(worklogCard);
  }
}
