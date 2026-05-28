/* ============================================================
   InterviewAI — Frontend Application Logic
   3-page SPA: Home → Form → Results
   Backend: Flask at http://localhost:5000
   ============================================================ */

const API_BASE = "http://localhost:5000";

// ── State ─────────────────────────────────────────────────────────────────
let currentData = null;
let currentJobRole = "";
let activeTab = "technical";
let interviewType = "mixed";
let activeHistoryId = null;

// ── DOM: Pages & Nav ───────────────────────────────────────────────────────
const pageHome = document.getElementById("page-home");
const pageForm = document.getElementById("page-form");
const pageResults = document.getElementById("page-results");
const navActions = document.getElementById("nav-actions");
const navBrand = document.getElementById("nav-brand-link");

// ── DOM: Form ──────────────────────────────────────────────────────────────
const form = document.getElementById("generator-form");
const jobRoleInput = document.getElementById("job-role");
const skillsInput = document.getElementById("skills");
const expSelect = document.getElementById("experience-level");
const typeHidden = document.getElementById("interview-type");
const numRange = document.getElementById("num-questions");
const numDisplay = document.getElementById("num-display");
const rangeBadge = document.getElementById("range-badge");
const contextArea = document.getElementById("company-context");
const generateBtn = document.getElementById("generate-btn");
const btnText = document.getElementById("btn-text");
const btnIcon = document.getElementById("btn-icon");

// ── DOM: Results ───────────────────────────────────────────────────────────
const loadingState = document.getElementById("loading-state");
const errorBanner = document.getElementById("error-banner");
const errorMessage = document.getElementById("error-message");
const resultsContent = document.getElementById("results-content");
const resultsEmptyState = document.getElementById("results-empty-state");

const resultsTitle = document.getElementById("results-title");
const resultsMeta = document.getElementById("results-meta");
const statTotal = document.getElementById("stat-total");
const statTech = document.getElementById("stat-tech");
const statBehav = document.getElementById("stat-behav");
const statSit = document.getElementById("stat-sit");
const tcTech = document.getElementById("tc-tech");
const tcBehav = document.getElementById("tc-behav");
const tcSit = document.getElementById("tc-sit");

const panelTech = document.getElementById("panel-technical");
const panelBehav = document.getElementById("panel-behavioral");
const panelSit = document.getElementById("panel-situational");
const panelEval = document.getElementById("panel-evaluation");

const btnDocx = document.getElementById("btn-download-docx");
const btnPrint = document.getElementById("btn-print");
const btnCopyAll = document.getElementById("btn-copy-all");
const toast = document.getElementById("toast");

// ── DOM: History ───────────────────────────────────────────────────────────
const historyList = document.getElementById("history-list");
const historyCount = document.getElementById("history-count");
const btnClearHistory = document.getElementById("btn-clear-history");
const homeHistoryGrid = document.getElementById("home-history-grid");

// ── Page Navigation ────────────────────────────────────────────────────────
function showPage(pageId) {
  [pageHome, pageForm, pageResults].forEach(p => p.classList.add("hidden"));
  document.getElementById(pageId).classList.remove("hidden");
  window.scrollTo({ top: 0, behavior: "smooth" });
  setNavActions(pageId);

  if (pageId === "page-home") {
    document.querySelectorAll(".details-section").forEach(el => {
      el.classList.add("initially-hidden");
    });
  }
}

function setNavActions(pageId) {
  if (pageId === "page-home") {
    navActions.innerHTML = `
      <button class="nav-btn-primary" id="nav-start-btn">✨ Let's Start</button>`;
    document.getElementById("nav-start-btn")
      .addEventListener("click", () => showPage("page-form"));

  } else if (pageId === "page-form") {
    navActions.innerHTML = `
      <button class="nav-btn-secondary" id="nav-home-btn">← Home</button>`;
    document.getElementById("nav-home-btn")
      .addEventListener("click", () => showPage("page-home"));

  } else if (pageId === "page-results") {
    navActions.innerHTML = `
      <button class="nav-btn-secondary" id="nav-new-btn">✨ New Interview</button>
      <button class="nav-btn-ghost"     id="nav-home-r-btn">🏠 Home</button>`;
    document.getElementById("nav-new-btn")
      .addEventListener("click", () => showPage("page-form"));
    document.getElementById("nav-home-r-btn")
      .addEventListener("click", () => showPage("page-home"));
  }
}

// Clicking the brand logo → Home
navBrand.addEventListener("click", () => showPage("page-home"));
navBrand.addEventListener("keydown", e => { if (e.key === "Enter") showPage("page-home"); });

// ── Results State Helpers ──────────────────────────────────────────────────
function showResultsState(which) {
  // which: "loading" | "results" | "empty"
  loadingState.classList.toggle("hidden", which !== "loading");
  resultsContent.classList.toggle("hidden", which !== "results");
  resultsEmptyState.classList.toggle("hidden", which !== "empty");
}

function hideError() { errorBanner.classList.add("hidden"); }

function showError(msg) {
  errorMessage.textContent = msg;
  errorBanner.classList.remove("hidden");
  showResultsState("empty");
}

// ── Toast ──────────────────────────────────────────────────────────────────
let toastTimer;
function showToast(msg, type = "success") {
  clearTimeout(toastTimer);
  toast.textContent = (type === "success" ? "✅ " : "❌ ") + msg;
  toast.className = `toast show ${type}`;
  toastTimer = setTimeout(() => toast.classList.remove("show"), 3000);
}

// ── Range Slider ───────────────────────────────────────────────────────────
function updateRange() {
  const min = +numRange.min, max = +numRange.max, val = +numRange.value;
  const pct = ((val - min) / (max - min)) * 100;
  numRange.style.setProperty("--pct", pct + "%");
  numDisplay.textContent = val;
  rangeBadge.textContent = val;
  numRange.setAttribute("aria-valuenow", val);
}
numRange.addEventListener("input", updateRange);
updateRange();

// ── Interview Type Toggle ──────────────────────────────────────────────────
document.querySelectorAll(".type-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".type-btn").forEach(b => {
      b.classList.remove("active");
      b.setAttribute("aria-pressed", "false");
    });
    btn.classList.add("active");
    btn.setAttribute("aria-pressed", "true");
    interviewType = btn.dataset.type;
    typeHidden.value = interviewType;
  });
});

// ── Tab Navigation ─────────────────────────────────────────────────────────
document.querySelectorAll(".tab-btn").forEach(btn => {
  btn.addEventListener("click", () => {
    document.querySelectorAll(".tab-btn").forEach(b => {
      b.classList.remove("active");
      b.setAttribute("aria-selected", "false");
    });
    btn.classList.add("active");
    btn.setAttribute("aria-selected", "true");
    activeTab = btn.dataset.tab;
    [panelTech, panelBehav, panelSit, panelEval].forEach(p => p.classList.add("hidden"));
    ({ technical: panelTech, behavioral: panelBehav, situational: panelSit, evaluation: panelEval })[activeTab]
      ?.classList.remove("hidden");
  });
});

// ── Badge Helpers ──────────────────────────────────────────────────────────
function diffClass(diff = "") {
  const d = diff.toLowerCase();
  if (d === "easy") return "badge-easy";
  if (d === "medium") return "badge-medium";
  if (d === "hard") return "badge-hard";
  return "badge-easy";
}
function catBadge(cat = "") {
  const c = cat.toLowerCase();
  if (c === "technical") return { cls: "badge-tech", label: "⚙️ Technical" };
  if (c === "behavioral") return { cls: "badge-behav", label: "🤝 Behavioral" };
  if (c === "situational") return { cls: "badge-sit", label: "💡 Situational" };
  return { cls: "badge-tech", label: cat };
}

// ── Render Question Card ───────────────────────────────────────────────────
function renderCard(q, index, cardType) {
  const typeClassMap = { technical: "tech", behavioral: "behav", situational: "sit" };
  const typeClass = typeClassMap[cardType] ?? "tech";
  const cat = catBadge(q.category || cardType);
  const dc = diffClass(q.difficulty);

  return `
    <article class="question-card ${typeClass}" style="animation-delay:${index * 0.06}s">
      <div class="card-header">
        <div class="card-badges">
          <span class="badge ${cat.cls}">${cat.label}</span>
          <span class="badge ${dc}">${q.difficulty || "Medium"}</span>
        </div>
        <div class="card-actions">
          <button class="btn-icon" title="Copy question" onclick="copyQuestion(this, ${JSON.stringify(q.question).replace(/'/g, "&#39;")})">📋</button>
        </div>
      </div>
      <div class="question-num">Question ${index + 1}</div>
      <p class="question-text">${escHtml(q.question)}</p>
      <div class="card-details">
        ${q.follow_up ? `
          <div class="detail-row">
            <span class="detail-label">↩ Follow-up:</span>
            <span class="detail-text">${escHtml(q.follow_up)}</span>
          </div>` : ""}
        ${q.evaluation_tip ? `
          <div class="detail-row">
            <span class="detail-label">💡 Eval Tip:</span>
            <span class="detail-text">${escHtml(q.evaluation_tip)}</span>
          </div>` : ""}
      </div>
    </article>`;
}

// ── Render Evaluation Panel ────────────────────────────────────────────────
function renderEvaluation(data) {
  const ec = data.evaluation_criteria || {};
  const tips = data.interview_tips || [];

  const competencies = (ec.key_competencies || [])
    .map(c => `<span class="comp-tag">${escHtml(c)}</span>`).join("");

  const greenFlags = (ec.green_flags || [])
    .map(f => `<li><span class="bullet" style="color:var(--diff-easy-c)">✓</span>${escHtml(f)}</li>`).join("");

  const redFlags = (ec.red_flags || [])
    .map(f => `<li><span class="bullet" style="color:var(--diff-hard-c)">✗</span>${escHtml(f)}</li>`).join("");

  const tipsHtml = tips
    .map(t => `<div class="tip-item"><span class="tip-icon">💡</span>${escHtml(t)}</div>`).join("");

  panelEval.innerHTML = `
    <div class="question-card eval" style="padding:1.25rem">
      <div class="eval-card-title" style="font-size:0.85rem;font-weight:700;color:var(--text-secondary);margin-bottom:0.75rem;">📊 Scoring Scale</div>
      <p style="font-size:0.88rem;color:var(--text-secondary);">${escHtml(ec.scoring_scale || "1–5 scale")}</p>
      ${competencies ? `
        <div style="margin-top:0.75rem;">
          <div style="font-size:0.75rem;font-weight:700;color:var(--text-muted);text-transform:uppercase;letter-spacing:.07em;margin-bottom:.5rem;">Key Competencies</div>
          <div class="competency-tags">${competencies}</div>
        </div>` : ""}
    </div>
    <div class="eval-grid">
      ${greenFlags ? `<div class="eval-card"><div class="eval-card-title" style="color:var(--diff-easy-c)">✓ Green Flags</div><ul class="eval-list">${greenFlags}</ul></div>` : ""}
      ${redFlags ? `<div class="eval-card"><div class="eval-card-title" style="color:var(--diff-hard-c)">✗ Red Flags</div><ul class="eval-list">${redFlags}</ul></div>` : ""}
    </div>
    ${tipsHtml ? `<div class="tips-section"><div class="tips-title">💬 Interviewer Tips</div><div class="tips-list">${tipsHtml}</div></div>` : ""}`;
}

// ── Render All Results ─────────────────────────────────────────────────────
function renderResults(data) {
  currentData = data;

  const tech = data.technical_questions || [];
  const behav = data.behavioral_questions || [];
  const sit = data.situational_questions || [];
  const total = tech.length + behav.length + sit.length;

  statTotal.textContent = total;
  statTech.textContent = tech.length;
  statBehav.textContent = behav.length;
  statSit.textContent = sit.length;
  tcTech.textContent = tech.length;
  tcBehav.textContent = behav.length;
  tcSit.textContent = sit.length;

  resultsTitle.textContent = `${data.job_role} — Interview Questions`;
  resultsMeta.innerHTML = `
    <span class="meta-chip">🎓 <strong>${capitalize(data.experience_level)}</strong></span>
    <span class="meta-chip">🔀 <strong>${capitalize(data.interview_type)}</strong></span>
    <span class="meta-chip">📝 <strong>${total}</strong> Questions</span>`;

  panelTech.innerHTML = tech.length ? tech.map((q, i) => renderCard(q, i, "technical")).join("") : emptyPanel("No technical questions generated.");
  panelBehav.innerHTML = behav.length ? behav.map((q, i) => renderCard(q, i, "behavioral")).join("") : emptyPanel("No behavioral questions generated.");
  panelSit.innerHTML = sit.length ? sit.map((q, i) => renderCard(q, i, "situational")).join("") : emptyPanel("No situational questions generated.");

  renderEvaluation(data);

  // Reset tabs
  document.querySelectorAll(".tab-btn").forEach(b => {
    b.classList.remove("active");
    b.setAttribute("aria-selected", "false");
  });
  document.getElementById("tab-technical").classList.add("active");
  document.getElementById("tab-technical").setAttribute("aria-selected", "true");
  [panelTech, panelBehav, panelSit, panelEval].forEach(p => p.classList.add("hidden"));
  panelTech.classList.remove("hidden");
  activeTab = "technical";

  showResultsState("results");
  hideError();
}

function emptyPanel(msg) {
  return `<div style="text-align:center;padding:3rem;color:var(--text-muted);font-size:0.9rem;">${msg}</div>`;
}

// ── Form Submission ────────────────────────────────────────────────────────
form.addEventListener("submit", async (e) => {
  e.preventDefault();
  hideError();

  const jobRole = jobRoleInput.value.trim();
  if (!jobRole) {
    jobRoleInput.focus();
    jobRoleInput.style.borderColor = "#f87171";
    setTimeout(() => (jobRoleInput.style.borderColor = ""), 2000);
    showToast("Please enter a job role.", "error");
    return;
  }

  currentJobRole = jobRole;

  // Navigate to results page and show loading
  showPage("page-results");
  showResultsState("loading");

  generateBtn.disabled = true;
  btnText.textContent = "Generating…";
  btnIcon.textContent = "⏳";

  try {
    const payload = {
      job_role: jobRole,
      skills: skillsInput.value.trim(),
      experience_level: expSelect.value,
      interview_type: interviewType,
      num_questions: parseInt(numRange.value, 10),
      company_context: contextArea.value.trim(),
    };

    const res = await fetch(`${API_BASE}/api/generate`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const json = await res.json();

    if (!res.ok) throw new Error(json.error || `Server error ${res.status}`);

    renderResults(json);
    showToast("Questions generated successfully!");
    loadHistory();

  } catch (err) {
    console.error(err);
    showError(err.message || "Failed to connect to the API. Make sure the Flask server is running.");
  } finally {
    generateBtn.disabled = false;
    btnText.textContent = "Generate Questions";
    btnIcon.textContent = "✨";
  }
});

// ── New interview from empty results state ─────────────────────────────────
document.getElementById("btn-new-from-empty")
  .addEventListener("click", () => showPage("page-form"));

// ── Copy Single Question ────────────────────────────────────────────────────
function copyQuestion(btn, questionText) {
  navigator.clipboard.writeText(questionText).then(() => {
    const orig = btn.textContent;
    btn.textContent = "✅";
    showToast("Question copied!");
    setTimeout(() => (btn.textContent = orig), 1500);
  });
}

// ── Copy All ────────────────────────────────────────────────────────────────
btnCopyAll.addEventListener("click", () => {
  if (!currentData) return;
  const tech = currentData.technical_questions || [];
  const behav = currentData.behavioral_questions || [];
  const sit = currentData.situational_questions || [];

  let text = `INTERVIEW QUESTIONS — ${currentData.job_role.toUpperCase()}\n`;
  text += `Experience: ${capitalize(currentData.experience_level)} | Type: ${capitalize(currentData.interview_type)}\n`;
  text += "=".repeat(60) + "\n\n";

  if (tech.length) {
    text += "TECHNICAL QUESTIONS\n" + "-".repeat(30) + "\n";
    tech.forEach((q, i) => { text += `Q${i + 1}. ${q.question}\n`; if (q.follow_up) text += `   ↩ ${q.follow_up}\n`; text += "\n"; });
  }
  if (behav.length) {
    text += "BEHAVIORAL QUESTIONS\n" + "-".repeat(30) + "\n";
    behav.forEach((q, i) => { text += `Q${i + 1}. ${q.question}\n`; if (q.follow_up) text += `   ↩ ${q.follow_up}\n`; text += "\n"; });
  }
  if (sit.length) {
    text += "SITUATIONAL QUESTIONS\n" + "-".repeat(30) + "\n";
    sit.forEach((q, i) => { text += `Q${i + 1}. ${q.question}\n`; if (q.follow_up) text += `   ↩ ${q.follow_up}\n`; text += "\n"; });
  }

  navigator.clipboard.writeText(text).then(() => showToast("All questions copied to clipboard!"));
});

// ── Print ──────────────────────────────────────────────────────────────────
btnPrint.addEventListener("click", () => window.print());

// ── Download DOCX ──────────────────────────────────────────────────────────
btnDocx.addEventListener("click", async () => {
  if (!currentData) return;
  const orig = btnDocx.textContent;
  btnDocx.textContent = "⏳ Downloading…";
  btnDocx.disabled = true;
  try {
    const res = await fetch(`${API_BASE}/api/export`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ questions_data: currentData, job_role: currentJobRole }),
    });
    if (!res.ok) { const err = await res.json(); throw new Error(err.error || "Export failed"); }

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${currentJobRole.replace(/\s+/g, "_")}_Interview_Questions.docx`;
    a.click();
    URL.revokeObjectURL(url);
    showToast("DOCX downloaded!");
  } catch (err) {
    showToast(err.message, "error");
  } finally {
    btnDocx.textContent = orig;
    btnDocx.disabled = false;
  }
});

// ── Utilities ──────────────────────────────────────────────────────────────
function escHtml(str = "") {
  return str
    .replace(/&/g, "&amp;").replace(/</g, "&lt;")
    .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
function capitalize(str = "") { return str.charAt(0).toUpperCase() + str.slice(1); }

// Expose for inline onclick
window.copyQuestion = copyQuestion;

// ══════════════════════════════════════════════════════════════════════════
//  HISTORY MANAGEMENT
// ══════════════════════════════════════════════════════════════════════════

function formatRelativeTime(iso) {
  const diff = Math.floor((new Date() - new Date(iso)) / 1000);
  if (diff < 60) return "Just now";
  if (diff < 3600) { const m = Math.floor(diff / 60); return `${m} min${m > 1 ? "s" : ""} ago`; }
  if (diff < 86400) { const h = Math.floor(diff / 3600); return `${h} hour${h > 1 ? "s" : ""} ago`; }
  const d = Math.floor(diff / 86400);
  if (d < 7) return `${d} day${d > 1 ? "s" : ""} ago`;
  return new Date(iso).toLocaleDateString(undefined, { month: "short", day: "numeric" });
}

// ── Load & refresh ─────────────────────────────────────────────────────────
async function loadHistory() {
  try {
    const res = await fetch(`${API_BASE}/api/history`);
    const items = await res.json();
    const arr = Array.isArray(items) ? items : [];
    renderHistoryList(arr);
    renderHomeHistoryGrid(arr);
  } catch (err) {
    console.warn("History fetch failed:", err);
    renderHistoryList([]);
    renderHomeHistoryGrid([]);
  }
}

// ── Sidebar history (Results page) ─────────────────────────────────────────
function renderHistoryList(items) {
  historyCount.textContent = items.length;
  if (items.length === 0) {
    historyList.innerHTML = `<div class="history-empty">No saved interviews yet.</div>`;
    btnClearHistory.classList.add("hidden");
    return;
  }
  btnClearHistory.classList.remove("hidden");

  const typeLabel = { technical: "⚙️ Tech", behavioral: "🤝 Behav", mixed: "🔀 Mixed", situational: "💡 Sit" };
  const expLabel = { beginner: "🌱 Beginner", intermediate: "🔧 Mid", advanced: "⚡ Adv", senior: "🏆 Senior" };

  historyList.innerHTML = items.map(item => {
    const isActive = item.id === activeHistoryId ? " active" : "";
    return `
      <div class="history-card${isActive}"
           id="hcard-${item.id}"
           onclick="loadHistoryItem('${item.id}')"
           role="button" tabindex="0"
           aria-label="Load: ${escHtml(item.job_role)}">
        <div class="history-info">
          <div class="history-card-title" title="${escHtml(item.job_role)}">${escHtml(item.job_role)}</div>
          <div class="history-meta">
            <span>${expLabel[item.experience_level] || capitalize(item.experience_level)}</span>
            <span>${typeLabel[item.interview_type] || capitalize(item.interview_type)}</span>
            <span>${item.num_questions} Qs</span>
            <span>${formatRelativeTime(item.timestamp)}</span>
          </div>
        </div>
        <button class="btn-history-delete"
          onclick="deleteHistoryItem(event,'${item.id}')"
          title="Delete" aria-label="Delete ${escHtml(item.job_role)}">✕</button>
      </div>`;
  }).join("");
}

// ── Home page history grid ──────────────────────────────────────────────────
function renderHomeHistoryGrid(items) {
  if (!homeHistoryGrid) return;

  if (items.length === 0) {
    homeHistoryGrid.innerHTML = `
      <div class="home-history-empty">
        <div class="home-history-empty-icon">📭</div>
        <p>No sessions yet. Generate your first interview question set!</p>
        <button class="btn-cta-secondary" onclick="showPage('page-form')">Start Generating →</button>
      </div>`;
    return;
  }

  const typeLabel = { technical: "⚙️ Technical", behavioral: "🤝 Behavioral", mixed: "🔀 Mixed", situational: "💡 Situational" };
  const expLabel = { beginner: "🌱 Beginner", intermediate: "🔧 Intermediate", advanced: "⚡ Advanced", senior: "🏆 Senior" };
  const recent = items.slice(0, 6);

  homeHistoryGrid.innerHTML = recent.map(item => `
    <div class="home-history-card"
         onclick="loadHistoryItem('${item.id}')"
         role="button" tabindex="0"
         aria-label="View: ${escHtml(item.job_role)}">
      <div class="hhc-icon">📋</div>
      <div class="hhc-content">
        <div class="hhc-title">${escHtml(item.job_role)}</div>
        <div class="hhc-meta">
          <span class="hhc-badge">${expLabel[item.experience_level] || capitalize(item.experience_level)}</span>
          <span class="hhc-badge">${typeLabel[item.interview_type] || capitalize(item.interview_type)}</span>
          <span class="hhc-badge">${item.num_questions} Questions</span>
        </div>
        <div class="hhc-time">${formatRelativeTime(item.timestamp)}</div>
      </div>
      <div class="hhc-arrow">→</div>
    </div>`).join("") +
    (items.length > 6 ? `<div class="home-history-more">+${items.length - 6} more sessions in history</div>` : "");
}

// ── Load history item ─────────────────────────────────────────────────────
async function loadHistoryItem(id) {
  activeHistoryId = id;

  // Navigate to results page
  showPage("page-results");
  showResultsState("loading");
  hideError();

  // Highlight active card in sidebar
  document.querySelectorAll(".history-card").forEach(c => c.classList.remove("active"));
  const card = document.getElementById(`hcard-${id}`);
  if (card) card.classList.add("active");

  try {
    const res = await fetch(`${API_BASE}/api/history/${id}`);
    if (!res.ok) throw new Error("History item not found");
    const item = await res.json();

    const payload = {
      ...item.data,
      job_role: item.job_role,
      experience_level: item.experience_level,
      interview_type: item.interview_type,
      id: item.id,
    };
    currentJobRole = item.job_role;
    renderResults(payload);
    showToast(`Loaded: ${item.job_role}`);

  } catch (err) {
    showError(err.message || "Failed to load history item.");
    activeHistoryId = null;
  }
}

// ── Delete single history item ─────────────────────────────────────────────
async function deleteHistoryItem(event, id) {
  event.stopPropagation();
  try {
    const res = await fetch(`${API_BASE}/api/history/${id}`, { method: "DELETE" });
    if (!res.ok) throw new Error("Delete failed");

    if (activeHistoryId === id) {
      activeHistoryId = null;
      showResultsState("empty");
    }
    showToast("History entry deleted.");
    loadHistory();
  } catch (err) {
    showToast(err.message || "Failed to delete.", "error");
  }
}

// ── Clear all history ──────────────────────────────────────────────────────
btnClearHistory.addEventListener("click", async () => {
  if (!confirm("Clear all saved history? This cannot be undone.")) return;
  try {
    const res = await fetch(`${API_BASE}/api/history`);
    const items = await res.json();
    await Promise.all(items.map(item =>
      fetch(`${API_BASE}/api/history/${item.id}`, { method: "DELETE" })
    ));
    activeHistoryId = null;
    showResultsState("empty");
    showToast("All history cleared.");
    loadHistory();
  } catch (err) {
    showToast("Failed to clear history.", "error");
  }
});

// ── Keyboard support for history cards ─────────────────────────────────────
document.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && e.target.classList.contains("history-card")) e.target.click();
});

// ── Home page button wiring ────────────────────────────────────────────────
document.getElementById("btn-get-started")
  .addEventListener("click", () => showPage("page-form"));
document.getElementById("btn-get-started-2")
  .addEventListener("click", () => showPage("page-form"));
const btnStartEmpty = document.getElementById("btn-start-empty");
if (btnStartEmpty) btnStartEmpty.addEventListener("click", () => showPage("page-form"));

// Explore More click handler
const btnExploreMore = document.getElementById("btn-explore-more");
if (btnExploreMore) {
  btnExploreMore.addEventListener("click", () => {
    document.querySelectorAll(".details-section").forEach(el => {
      el.classList.remove("initially-hidden");
    });
    const howSection = document.querySelector(".how-section");
    if (howSection) {
      howSection.scrollIntoView({ behavior: "smooth" });
    }
  });
}

// Expose globals for inline onclick
window.loadHistoryItem = loadHistoryItem;
window.deleteHistoryItem = deleteHistoryItem;
window.showPage = showPage;

// ── Initialise ─────────────────────────────────────────────────────────────
setNavActions("page-home");
loadHistory();
