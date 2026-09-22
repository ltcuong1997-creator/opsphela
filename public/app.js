// ------------------------------------------------------------------
// Phê La · Báo cáo — dashboard "khung", đọc dữ liệu tổng quát từ Firestore.
//
// Ý tưởng cốt lõi: mọi loại báo cáo (doanh thu, bán hàng theo cửa hàng,
// v.v.) đều đổ về cùng một cấu trúc dữ liệu chung trong Firestore:
//
//   reportTypes/{reportTypeId}          -> mô tả 1 loại báo cáo (tên,
//                                          loại biểu đồ, các chỉ số)
//   reports/{reportTypeId}/entries/{id} -> từng dòng dữ liệu của báo cáo đó
//
// Muốn thêm 1 dashboard mới: chỉ cần thêm 1 document vào "reportTypes"
// và đổ dữ liệu đúng cấu trúc vào "reports/{reportTypeId}/entries" —
// không cần sửa file này. Xem README.md phần "Thêm báo cáo mới".
// ------------------------------------------------------------------

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.13.0/firebase-app.js";
import {
  getAuth,
  onAuthStateChanged,
  signInWithEmailAndPassword,
  signOut,
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-auth.js";
import {
  getFirestore,
  collection,
  getDocs,
  query,
  where,
  orderBy,
} from "https://www.gstatic.com/firebasejs/10.13.0/firebase-firestore.js";

// Config Web App của project "opsphela" (Firebase Console -> Project settings).
// apiKey này là khoá công khai của web app, không phải bí mật - quyền truy cập
// dữ liệu do Firebase Auth + firestore.rules quyết định.
const firebaseConfig = {
  apiKey: "AIzaSyDou4bM03zukkzgDpVUWZJOd2owM5rnd6E",
  authDomain: "opsphela.firebaseapp.com",
  projectId: "opsphela",
  storageBucket: "opsphela.firebasestorage.app",
  messagingSenderId: "39705637257",
  appId: "1:39705637257:web:205d8fdad1cb0fbd3a3158",
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// ---------------- Màu theo thứ tự cố định (không bao giờ đổi thứ tự) ----------------
const SERIES_COLORS = [
  "var(--series-1)", "var(--series-2)", "var(--series-3)", "var(--series-4)",
  "var(--series-5)", "var(--series-6)", "var(--series-7)", "var(--series-8)",
];
function resolveColor(varName) {
  return getComputedStyle(document.documentElement).getPropertyValue(varName.replace(/var\((.+)\)/, "$1")).trim();
}

// ---------------- State ----------------
let reportTypes = [];
let activeReportType = null;
let activeRange = "7d";
let chartInstance = null;

// ---------------- Auth ----------------
const loginScreen = document.getElementById("login-screen");
const appShell = document.getElementById("app-shell");
const loginForm = document.getElementById("login-form");
const loginError = document.getElementById("login-error");

loginForm.addEventListener("submit", async (e) => {
  e.preventDefault();
  loginError.hidden = true;
  const email = document.getElementById("login-email").value.trim();
  const password = document.getElementById("login-password").value;
  try {
    await signInWithEmailAndPassword(auth, email, password);
  } catch (err) {
    loginError.textContent = "Đăng nhập thất bại: " + humanizeAuthError(err.code);
    loginError.hidden = false;
  }
});

document.getElementById("logout-btn").addEventListener("click", () => signOut(auth));

function humanizeAuthError(code) {
  if (code === "auth/invalid-credential" || code === "auth/wrong-password") return "sai email hoặc mật khẩu";
  if (code === "auth/user-not-found") return "tài khoản không tồn tại";
  if (code === "auth/too-many-requests") return "thử lại sau ít phút";
  return code || "lỗi không xác định";
}

onAuthStateChanged(auth, async (user) => {
  if (user) {
    loginScreen.hidden = true;
    appShell.hidden = false;
    await loadReportTypes();
  } else {
    loginScreen.hidden = false;
    appShell.hidden = true;
  }
});

// ---------------- Report types (sidebar) ----------------
async function loadReportTypes() {
  const nav = document.getElementById("report-nav");
  nav.innerHTML = "<p class='muted small'>Đang tải...</p>";

  const snap = await getDocs(query(collection(db, "reportTypes"), orderBy("order", "asc")));
  reportTypes = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

  nav.innerHTML = "";
  if (reportTypes.length === 0) {
    nav.innerHTML = "<p class='muted small'>Chưa có báo cáo nào. Thêm document vào collection \"reportTypes\".</p>";
    return;
  }

  reportTypes.forEach((rt) => {
    const btn = document.createElement("button");
    btn.textContent = rt.label || rt.id;
    btn.dataset.id = rt.id;
    btn.addEventListener("click", () => selectReportType(rt.id));
    nav.appendChild(btn);
  });

  selectReportType(reportTypes[0].id);
}

function selectReportType(id) {
  activeReportType = reportTypes.find((r) => r.id === id);
  document.querySelectorAll("#report-nav button").forEach((b) => b.classList.toggle("active", b.dataset.id === id));
  document.getElementById("report-title").textContent = activeReportType.label || id;
  loadAndRender();
}

// ---------------- Date range ----------------
document.querySelectorAll("#date-range button").forEach((btn) => {
  btn.addEventListener("click", () => {
    activeRange = btn.dataset.range;
    document.querySelectorAll("#date-range button").forEach((b) => b.classList.toggle("active", b === btn));
    loadAndRender();
  });
});

function currentDateRange() {
  const end = new Date();
  const start = new Date();
  if (activeRange === "today") {
    // start = end = hôm nay
  } else if (activeRange === "7d") {
    start.setDate(start.getDate() - 6);
  } else if (activeRange === "30d") {
    start.setDate(start.getDate() - 29);
  } else if (activeRange === "mtd") {
    start.setDate(1);
  }
  const fmt = (d) => d.toISOString().slice(0, 10);
  return { start: fmt(start), end: fmt(end) };
}

// ---------------- Load + render ----------------
async function loadAndRender() {
  if (!activeReportType) return;
  const { start, end } = currentDateRange();

  const entriesRef = collection(db, "reports", activeReportType.id, "entries");
  const q = query(entriesRef, where("date", ">=", start), where("date", "<=", end), orderBy("date", "asc"));
  const snap = await getDocs(q);
  const entries = snap.docs.map((d) => d.data());

  renderKpis(entries);
  renderChart(entries);
  renderTable(entries);
}

function renderKpis(entries) {
  const row = document.getElementById("kpi-row");
  row.innerHTML = "";
  const metricKeys = activeReportType.metricKeys || [];
  metricKeys.forEach((key) => {
    const total = entries.reduce((sum, e) => sum + (Number(e.metrics?.[key]) || 0), 0);
    const unit = (activeReportType.units && activeReportType.units[key]) || "";
    const tile = document.createElement("div");
    tile.className = "kpi-tile";
    tile.innerHTML = `
      <div class="label">${activeReportType.metricLabels?.[key] || key}</div>
      <div class="value">${total.toLocaleString("vi-VN")}${unit ? " " + unit : ""}</div>
    `;
    row.appendChild(tile);
  });
}

function renderChart(entries) {
  const card = document.querySelector(".chart-card");
  const canvas = document.getElementById("main-chart");
  const empty = document.getElementById("chart-empty");

  if (activeReportType.chartType === "table") {
    card.hidden = true;
    return;
  }
  card.hidden = false;

  if (entries.length === 0) {
    canvas.hidden = true;
    empty.hidden = false;
    if (chartInstance) chartInstance.destroy();
    return;
  }
  canvas.hidden = false;
  empty.hidden = true;

  const dimensionKey = activeReportType.dimensionKey || "date";
  const metricKeys = activeReportType.metricKeys || [];

  // Gộp theo dimensionKey (ví dụ theo ngày, hoặc theo cửa hàng)
  const groups = new Map();
  entries.forEach((e) => {
    const key = e.dimensions?.[dimensionKey] ?? e[dimensionKey] ?? "?";
    if (!groups.has(key)) groups.set(key, {});
    const bucket = groups.get(key);
    metricKeys.forEach((m) => {
      bucket[m] = (bucket[m] || 0) + (Number(e.metrics?.[m]) || 0);
    });
  });

  const labels = [...groups.keys()];
  const datasets = metricKeys.map((m, i) => ({
    label: activeReportType.metricLabels?.[m] || m,
    data: labels.map((l) => groups.get(l)[m] || 0),
    borderColor: resolveColor(SERIES_COLORS[i % SERIES_COLORS.length]),
    backgroundColor: resolveColor(SERIES_COLORS[i % SERIES_COLORS.length]),
    borderWidth: 2,
    borderRadius: 4,
    tension: 0.25,
    pointRadius: 3,
  }));

  if (chartInstance) chartInstance.destroy();
  chartInstance = new Chart(canvas, {
    type: activeReportType.chartType === "bar" ? "bar" : "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: datasets.length > 1, labels: { color: resolveColor("var(--text-secondary)") } },
        tooltip: { enabled: true },
      },
      scales: {
        x: { grid: { color: resolveColor("var(--grid)") }, ticks: { color: resolveColor("var(--text-muted)") } },
        y: { grid: { color: resolveColor("var(--grid)") }, ticks: { color: resolveColor("var(--text-muted)") }, beginAtZero: true },
      },
    },
  });
}

function renderTable(entries) {
  const head = document.getElementById("table-head");
  const body = document.getElementById("table-body");
  head.innerHTML = "";
  body.innerHTML = "";
  if (entries.length === 0) return;

  const dimensionCols = Object.keys(entries[0].dimensions || {});
  const metricCols = activeReportType.metricKeys || Object.keys(entries[0].metrics || {});
  const cols = ["date", ...dimensionCols, ...metricCols];

  cols.forEach((c) => {
    const th = document.createElement("th");
    th.textContent = activeReportType.metricLabels?.[c] || c;
    head.appendChild(th);
  });

  entries.slice(0, 200).forEach((e) => { // giới hạn 200 dòng hiển thị cho nhẹ trang
    const tr = document.createElement("tr");
    cols.forEach((c) => {
      const td = document.createElement("td");
      let val;
      if (c === "date") val = e.date;
      else if (dimensionCols.includes(c)) val = e.dimensions?.[c];
      else val = e.metrics?.[c];
      td.textContent = typeof val === "number" ? val.toLocaleString("vi-VN") : (val ?? "");
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
}
