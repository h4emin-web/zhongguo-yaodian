const cards = document.querySelectorAll(".tool-box");
const detailView = document.querySelector(".detail-view");
const detailPanel = document.querySelector(".detail-panel");
const closeButton = document.querySelector(".detail-close");
const detailKicker = document.querySelector(".detail-kicker");
const detailTitle = document.querySelector("#detail-title");
const detailDescription = document.querySelector(".detail-description");
const searchInput = document.querySelector("#workspace-search");
const connectionStatus = document.querySelector(".connection-status");
const detailPlaceholder = document.querySelector(".detail-placeholder");
const wcSearchPanel = document.querySelector(".wc-search-panel");
const wcSearchInput = document.querySelector("#wc-search-input");
const wcResults = document.querySelector(".wc-results");
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
  detailPlaceholder.hidden = isWcSearch;

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

  if (results.length === 0) {
    wcResults.innerHTML = '<p class="empty-result">검색 결과가 없습니다.</p>';
    return;
  }

  wcResults.innerHTML = results.map((row) => `
    <article class="wc-result-item">
      <p><strong>중국어</strong><span>${escapeHtml(row.chinese)}</span></p>
      <p><strong>한국어</strong><span>${escapeHtml(row.korean)}</span></p>
      <p><strong>영어</strong><span>${escapeHtml(row.english)}</span></p>
    </article>
  `).join("");
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

async function checkSupabaseConnection() {
  if (!connectionStatus) {
    return;
  }

  if (!supabaseClient) {
    connectionStatus.textContent = "Supabase anon key를 넣으면 연결됩니다";
    connectionStatus.dataset.status = "waiting";
    return;
  }

  try {
    const { error } = await supabaseClient.auth.getSession();

    if (error) {
      throw error;
    }

    connectionStatus.textContent = "Supabase 연결됨";
    connectionStatus.dataset.status = "connected";
  } catch (error) {
    connectionStatus.textContent = "Supabase 연결 확인 실패";
    connectionStatus.dataset.status = "error";
    console.error("Supabase connection check failed:", error);
  }
}

checkSupabaseConnection();
