const cards = document.querySelectorAll(".tool-box");
const detailView = document.querySelector(".detail-view");
const detailPanel = document.querySelector(".detail-panel");
const closeButton = document.querySelector(".detail-close");
const detailKicker = document.querySelector(".detail-kicker");
const detailTitle = document.querySelector("#detail-title");
const detailDescription = document.querySelector(".detail-description");
const searchInput = document.querySelector("#workspace-search");
const connectionStatus = document.querySelector(".connection-status");
const supabaseConfig = window.HAEMIN_SUPABASE || {};
const hasSupabaseKey = Boolean(
  supabaseConfig.anonKey && supabaseConfig.anonKey !== "YOUR_SUPABASE_ANON_KEY"
);
const supabaseClient = hasSupabaseKey && window.supabase
  ? window.supabase.createClient(supabaseConfig.url, supabaseConfig.anonKey)
  : null;

let lastFocusedCard = null;

function openDetail(card) {
  const kicker = card.querySelector(".tool-kicker").textContent;
  const title = card.querySelector("h2").textContent;
  const description = card.querySelector("p:last-child").textContent;

  lastFocusedCard = card;
  detailKicker.textContent = kicker;
  detailTitle.textContent = title;
  detailDescription.textContent = description;
  detailView.classList.add("is-open");
  detailView.setAttribute("aria-hidden", "false");
  document.body.classList.add("detail-open");
  closeButton.focus();
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
