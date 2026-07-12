const loginView = document.querySelector("[data-login-view]");
const dashboard = document.querySelector("[data-dashboard]");
const loginForm = document.querySelector("[data-login-form]");
const loginError = document.querySelector("[data-login-error]");
const daysSelect = document.querySelector("[data-days]");

async function request(url, options) {
  const response = await fetch(url, { ...options, headers: { "Content-Type": "application/json", ...(options?.headers || {}) } });
  if (!response.ok) throw new Error((await response.json().catch(() => ({}))).error || `请求失败: ${response.status}`);
  return response.json();
}

function showDashboard() {
  loginView.hidden = true;
  dashboard.hidden = false;
}

function showLogin() {
  dashboard.hidden = true;
  loginView.hidden = false;
}

function renderDaily(rows) {
  const chart = document.querySelector("[data-daily-chart]");
  if (rows.length === 0) {
    const empty = document.createElement("p");
    empty.className = "empty-data";
    empty.textContent = "暂无下载数据";
    chart.replaceChildren(empty);
    return;
  }
  const maximum = Math.max(1, ...rows.map((row) => row.downloads));
  chart.replaceChildren(...rows.map((row, index) => {
    const item = document.createElement("div");
    item.className = "bar-item";
    item.title = `${row.day}: ${row.downloads}`;
    const bar = document.createElement("i");
    bar.style.height = `${Math.max(2, row.downloads / maximum * 100)}%`;
    item.append(bar);
    if (index === 0 || index === rows.length - 1 || index % 7 === 0) {
      const label = document.createElement("span");
      label.textContent = row.day.slice(5);
      item.append(label);
    }
    return item;
  }));
}

function renderVersions(rows) {
  const list = document.querySelector("[data-version-list]");
  if (rows.length === 0) {
    const empty = document.createElement("p");
    empty.className = "empty-data";
    empty.textContent = "暂无版本数据";
    list.replaceChildren(empty);
    return;
  }
  const maximum = Math.max(1, ...rows.map((row) => row.downloads));
  list.replaceChildren(...rows.map((row) => {
    const line = document.createElement("div");
    line.className = "version-row";
    const version = document.createElement("span");
    version.textContent = `v${row.version}`;
    const track = document.createElement("span");
    track.className = "version-track";
    const fill = document.createElement("i");
    fill.style.width = `${row.downloads / maximum * 100}%`;
    track.append(fill);
    const count = document.createElement("strong");
    count.textContent = row.downloads;
    line.append(version, track, count);
    return line;
  }));
}

function renderRecent(rows) {
  const body = document.querySelector("[data-recent]");
  if (rows.length === 0) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.colSpan = 5;
    cell.className = "empty-table";
    cell.textContent = "暂无下载记录";
    row.append(cell);
    body.replaceChildren(row);
    return;
  }
  body.replaceChildren(...rows.map((row) => {
    const tr = document.createElement("tr");
    const values = [
      new Date(row.downloadedAt).toLocaleString("zh-CN"),
      `v${row.version}`,
      row.visitor,
      row.referer || "直接访问",
      row.userAgent || "未知",
    ];
    for (const value of values) {
      const td = document.createElement("td");
      td.textContent = value;
      td.title = value;
      tr.append(td);
    }
    return tr;
  }));
}

async function loadDashboard() {
  const [stats, release] = await Promise.all([
    request(`/api/admin/stats?days=${daysSelect.value}`),
    request("/api/admin/release").catch(() => null),
  ]);
  document.querySelector("[data-total]").textContent = stats.summary.total || 0;
  document.querySelector("[data-unique]").textContent = stats.summary.uniqueDownloads || 0;
  document.querySelector("[data-today]").textContent = stats.summary.today || 0;
  document.querySelector("[data-version]").textContent = release ? `v${release.version}` : "-";
  document.querySelector("[data-current-release]").textContent = release
    ? `当前发布版本 v${release.version}`
    : "尚未发布安装包";
  renderDaily(stats.daily);
  renderVersions(stats.versions);
  renderRecent(stats.recent);
}

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginError.textContent = "";
  try {
    const form = new FormData(loginForm);
    await request("/api/admin/login", { method: "POST", body: JSON.stringify({ password: form.get("password") }) });
    showDashboard();
    await loadDashboard();
  } catch (error) {
    loginError.textContent = error.message;
  }
});

document.querySelector("[data-logout]").addEventListener("click", async () => {
  await request("/api/admin/logout", { method: "POST" }).catch(() => undefined);
  showLogin();
});
daysSelect.addEventListener("change", () => void loadDashboard());

request("/api/admin/session")
  .then(async () => {
    showDashboard();
    await loadDashboard();
  })
  .catch(showLogin);
