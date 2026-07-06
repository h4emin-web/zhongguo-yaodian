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
const excelExportButton = document.querySelector(".excel-export");
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

function openDetail(card) {
  const kicker = card.querySelector(".tool-kicker").textContent;
  const title = card.querySelector("h2").textContent;
  const description = card.querySelector("p:last-child").textContent;

  lastFocusedCard = card;
  detailKicker.textContent = kicker;
  detailTitle.textContent = title;
  detailDescription.textContent = description;
  showWcSearch(title.includes("WC & COPP"));
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");

  if (title.includes("WC & COPP")) {
    loadWcData();
    wcSearchInput.focus();
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
});

searchInput.addEventListener("input", () => {
  const keyword = searchInput.value.trim().toLowerCase();

  cards.forEach((card) => {
    const cardText = card.textContent.toLowerCase();
    card.hidden = keyword.length > 0 && !cardText.includes(keyword);
  });
});

function showWcSearch(isWcSearch) {
  wcSearchPanel.hidden = !isWcSearch;

  if (isWcSearch) {
    wcSearchInput.value = "";
    wcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
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
    wcLoaded = true;
    wcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
  } catch (error) {
    console.error(error);
    wcResults.innerHTML = '<p class="empty-result">데이터를 불러오지 못했습니다.</p>';
  }
}

function renderWcResults() {
  const keyword = wcSearchInput.value.trim().toLowerCase();

  if (!keyword) {
    wcResults.innerHTML = '<p class="empty-result">검색어를 입력하세요.</p>';
    return;
  }

  const results = wcRows
    .filter((row) => {
      return [row.chinese, row.korean, row.english]
        .join(" ")
        .toLowerCase()
        .includes(keyword);
    })
    .slice(0, 80);

  currentWcResults = results;

  if (results.length === 0) {
    wcResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  wcResults.innerHTML = results.map((row) => `
    <article class="wc-result-item">
      <p><strong>원료명</strong><span>${escapeHtml(row.chinese)} / ${escapeHtml(row.korean)} / ${escapeHtml(row.english)}</span></p>
      <p><strong>제조사</strong><span>${escapeHtml(row.manufacturer)}</span></p>
      <p><strong>유효기간</strong><span>${escapeHtml(row.validity)}</span></p>
      <p><strong>이메일</strong><span>${escapeHtml(row.email)}</span></p>
      <p><strong>연락처</strong><span>${escapeHtml(row.phone)}</span></p>
    </article>
  `).join("");
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
  workbook.creator = "Haemin WorkSpace";
  workbook.created = new Date();

  const worksheet = workbook.addWorksheet("WC COPP 검색결과");
  worksheet.columns = [
    { header: "원료명(중국어)", key: "chinese", width: 28 },
    { header: "원료명(한국어)", key: "korean", width: 34 },
    { header: "원료명(영어)", key: "english", width: 38 },
    { header: "제조사", key: "manufacturer", width: 36 },
    { header: "유효기간", key: "validity", width: 18 },
    { header: "이메일", key: "email", width: 30 },
    { header: "연락처", key: "phone", width: 22 }
  ];

  currentWcResults.forEach((row) => worksheet.addRow(row));

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = {
        name: "Nanum Gothic",
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

wcSearchInput.addEventListener("input", renderWcResults);
excelExportButton.addEventListener("click", exportWcResults);
