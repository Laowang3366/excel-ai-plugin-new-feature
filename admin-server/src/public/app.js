/**
 * 文格 AI 助手 — 激活管理后台 前端逻辑
 *
 * 无框架依赖，原生 JS SPA。
 */

// ============================================================
// 状态管理
// ============================================================

const STATE = {
  token: localStorage.getItem("admin_token") || null,
  username: localStorage.getItem("admin_username") || null,
  currentPage: "dashboard",
};

// API 基础路径
const API = "";

// ============================================================
// 工具函数
// ============================================================

function $(sel) {
  return document.querySelector(sel);
}

function $$(sel) {
  return document.querySelectorAll(sel);
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** 显示 Toast 消息 */
function showToast(message, type = "success") {
  let toast = $("#toast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "toast";
    toast.className = "toast";
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.className = `toast ${type}`;
  // Trigger reflow
  void toast.offsetWidth;
  toast.classList.add("show");
  clearTimeout(toast._hideTimer);
  toast._hideTimer = setTimeout(() => toast.classList.remove("show"), 3000);
}

function showConfirmDialog({ title = "确认操作", message, confirmText = "确定", danger = false } = {}) {
  let overlay = $("#confirm-modal");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "confirm-modal";
    overlay.className = "modal-overlay";
    overlay.innerHTML = `
      <div class="modal modal-compact" role="dialog" aria-modal="true">
        <div class="modal-header">
          <h2 id="confirm-title"></h2>
          <button class="modal-close" id="confirm-close" aria-label="关闭">×</button>
        </div>
        <div class="modal-body">
          <p class="confirm-message" id="confirm-message"></p>
        </div>
        <div class="modal-actions modal-footer">
          <button class="btn btn-outline" id="confirm-cancel">取消</button>
          <button class="btn" id="confirm-ok">确定</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
  }

  $("#confirm-title").textContent = title;
  $("#confirm-message").textContent = message || "";
  const okBtn = $("#confirm-ok");
  okBtn.textContent = confirmText;
  okBtn.className = danger ? "btn btn-danger" : "btn btn-primary";

  return new Promise((resolve) => {
    const close = (result) => {
      overlay.classList.remove("open");
      $("#confirm-close").removeEventListener("click", onCancel);
      $("#confirm-cancel").removeEventListener("click", onCancel);
      okBtn.removeEventListener("click", onConfirm);
      overlay.removeEventListener("click", onOverlay);
      resolve(result);
    };
    const onCancel = () => close(false);
    const onConfirm = () => close(true);
    const onOverlay = (event) => {
      if (event.target === overlay) close(false);
    };

    $("#confirm-close").addEventListener("click", onCancel);
    $("#confirm-cancel").addEventListener("click", onCancel);
    okBtn.addEventListener("click", onConfirm);
    overlay.addEventListener("click", onOverlay);
    overlay.classList.add("open");
  });
}

/** API 请求（自动注入 JWT） */
async function apiRequest(endpoint, options = {}) {
  const headers = { "Content-Type": "application/json", ...options.headers };
  if (STATE.token) {
    headers["Authorization"] = `Bearer ${STATE.token}`;
  }

  const response = await fetch(`${API}${endpoint}`, {
    ...options,
    headers,
  });

  const data = await response.json();

  if (!response.ok) {
    if (response.status === 401) {
      // Token 过期，跳转登录
      STATE.token = null;
      STATE.username = null;
      localStorage.removeItem("admin_token");
      localStorage.removeItem("admin_username");
      renderApp();
    }
    throw new Error(data.error || `请求失败 (${response.status})`);
  }

  return data;
}

// ============================================================
// 页面渲染函数
// ============================================================

/** 登录页面 */
function renderLogin() {
  const app = $("#app");
  app.innerHTML = `
    <div class="login-page">
      <div class="login-card">
        <h1>🔐 管理后台</h1>
        <p class="subtitle">文格 AI 助手 — 激活管理系统</p>
        <div id="login-error" class="error-message"></div>
        <form id="login-form">
          <div class="form-group">
            <label for="username">用户名</label>
            <input type="text" id="username" placeholder="请输入管理员用户名" autofocus>
          </div>
          <div class="form-group">
            <label for="password">密码</label>
            <input type="password" id="password" placeholder="请输入密码">
          </div>
          <button type="submit" class="btn btn-primary btn-block">登 录</button>
        </form>
      </div>
    </div>
  `;

  $("#login-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const username = $("#username").value.trim();
    const password = $("#password").value;
    const errorEl = $("#login-error");

    if (!username || !password) {
      errorEl.textContent = "请输入用户名和密码";
      errorEl.classList.add("visible");
      return;
    }

    try {
      const result = await apiRequest("/api/admin/login", {
        method: "POST",
        body: JSON.stringify({ username, password }),
        headers: {}, // 登录不需要 auth header
      });

      STATE.token = result.token;
      STATE.username = result.username;
      localStorage.setItem("admin_token", result.token);
      localStorage.setItem("admin_username", result.username);
      renderApp();
    } catch (err) {
      errorEl.textContent = err.message;
      errorEl.classList.add("visible");
    }
  });
}

/** 管理后台布局（侧边栏 + 内容区） */
function renderAdminLayout() {
  const app = $("#app");
  const navItems = [
    { id: "dashboard", icon: "📊", label: "仪表盘" },
    { id: "keys", icon: "🔑", label: "卡密管理" },
    { id: "monitor", icon: "📡", label: "在线监控" },
  ];

  app.innerHTML = `
    <div class="admin-layout">
      <aside class="sidebar">
        <div class="sidebar-header">
          <h2>文格 AI</h2>
          <div class="version">激活管理后台 v1.0</div>
        </div>
        <nav class="sidebar-nav" id="sidebar-nav">
          ${navItems
            .map(
              (item) => `
            <a href="#${item.id}" data-page="${item.id}" class="${STATE.currentPage === item.id ? "active" : ""}">
              <span class="nav-icon">${item.icon}</span>
              <span>${item.label}</span>
            </a>
          `
            )
            .join("")}
        </nav>
        <div class="sidebar-footer">
          <div class="admin-info">👤 ${escapeHtml(STATE.username)}</div>
          <button class="logout-btn" id="logout-btn">🚪 退出登录</button>
        </div>
      </aside>
      <main class="main-content" id="main-content">
        <!-- 由子页面填充 -->
      </main>
    </div>
  `;

  // 导航点击
  $("#sidebar-nav").addEventListener("click", (e) => {
    const link = e.target.closest("a[data-page]");
    if (link) {
      e.preventDefault();
      const page = link.dataset.page;
      navigateTo(page);
    }
  });

  // 退出登录
  $("#logout-btn").addEventListener("click", () => {
    STATE.token = null;
    STATE.username = null;
    localStorage.removeItem("admin_token");
    localStorage.removeItem("admin_username");
    renderApp();
  });

  // 监听 hash 变化
  window.addEventListener("hashchange", handleHashChange);

  // 渲染当前页面
  renderPage(STATE.currentPage);
}

/** 根据 hash 更新当前页面 */
function handleHashChange() {
  const hash = location.hash.slice(1) || "dashboard";
  if (hash !== STATE.currentPage && ["dashboard", "keys", "monitor"].includes(hash)) {
    STATE.currentPage = hash;
    // 更新导航高亮
    $$("#sidebar-nav a").forEach((a) => {
      a.classList.toggle("active", a.dataset.page === hash);
    });
    renderPage(hash);
  }
}

function navigateTo(page) {
  STATE.currentPage = page;
  location.hash = page;
  $$("#sidebar-nav a").forEach((a) => {
    a.classList.toggle("active", a.dataset.page === page);
  });
  renderPage(page);
}

/** 页面路由 */
function renderPage(page) {
  switch (page) {
    case "dashboard":
      renderDashboard();
      break;
    case "keys":
      renderKeys();
      break;
    case "monitor":
      renderMonitor();
      break;
  }
}

// ============================================================
// 仪表盘
// ============================================================

async function renderDashboard() {
  const container = $("#main-content");
  container.innerHTML = `<div class="loading-state"><div class="spinner"></div><p>加载中...</p></div>`;

  try {
    const data = await apiRequest("/api/admin/dashboard");

    container.innerHTML = `
      <div class="page-header">
        <h1>📊 仪表盘</h1>
      </div>

      <div class="stats-grid">
        <div class="stat-card blue">
          <div class="stat-label">总卡密数</div>
          <div class="stat-value">${data.totalKeys}</div>
          <div class="stat-sub">今日新增 ${data.todayNewKeys}</div>
        </div>
        <div class="stat-card green">
          <div class="stat-label">有效卡密</div>
          <div class="stat-value">${data.activeKeys}</div>
          <div class="stat-sub">已使用 ${data.usedKeys}</div>
        </div>
        <div class="stat-card orange">
          <div class="stat-label">今日激活</div>
          <div class="stat-value">${data.todayActivations}</div>
          <div class="stat-sub">累计 ${data.totalMachines} 台设备</div>
        </div>
        <div class="stat-card green">
          <div class="stat-label">当前在线</div>
          <div class="stat-value">${data.onlineNow}</div>
          <div class="stat-sub">${data.totalMachines} 台已激活</div>
        </div>
        <div class="stat-card red">
          <div class="stat-label">已禁用</div>
          <div class="stat-value">${data.disabledKeys}</div>
        </div>
      </div>

      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">
        <div class="detail-section">
          <h3>📈 近 7 天激活趋势</h3>
          ${renderWeeklyTrend(data.weeklyTrend)}
        </div>
        <div class="detail-section">
          <h3>🏆 在线时长排行 (Top 10)</h3>
          ${renderTopOnline(data.topOnline)}
        </div>
      </div>
    `;
  } catch (err) {
    container.innerHTML = `<div class="empty-state"><p>❌ 加载失败: ${escapeHtml(err.message)}</p></div>`;
  }
}

function renderWeeklyTrend(trend) {
  if (!trend || trend.length === 0) {
    return `<div class="empty-state"><p>暂无数据</p></div>`;
  }
  const maxCount = Math.max(...trend.map((d) => d.count), 1);
  return `
    <div style="display:flex;align-items:flex-end;gap:8px;height:120px;padding:8px 0;">
      ${trend
        .map((d) => {
          const height = (d.count / maxCount) * 100;
          return `
            <div style="flex:1;display:flex;flex-direction:column;align-items:center;gap:4px;">
              <div style="font-size:11px;color:var(--gray-500);">${d.count}</div>
              <div style="width:100%;background:var(--primary-light);border-radius:4px 4px 0 0;height:${height}%;min-height:4px;"></div>
              <div style="font-size:10px;color:var(--gray-400);">${d.day.slice(5)}</div>
            </div>
          `;
        })
        .join("")}
    </div>
  `;
}

function renderTopOnline(list) {
  if (!list || list.length === 0) {
    return `<div class="empty-state"><p>暂无数据</p></div>`;
  }
  return `
    <table>
      <thead>
        <tr>
          <th>#</th>
          <th>设备</th>
          <th>卡密</th>
          <th>在线时长</th>
        </tr>
      </thead>
      <tbody>
        ${list
          .map(
            (item, i) => `
          <tr>
            <td>${i + 1}</td>
            <td>${escapeHtml(item.machine_name || item.machine_id)}</td>
            <td style="font-family:monospace;">${escapeHtml(item.key_code)}</td>
            <td>${formatDuration(item.total_online_seconds)}</td>
          </tr>
        `
          )
          .join("")}
      </tbody>
    </table>
  `;
}

// ============================================================
// 卡密管理
// ============================================================

let keysPageState = {
  page: 1,
  pageSize: 20,
  status: "",
  search: "",
};
const selectedKeyIds = new Set();

async function renderKeys() {
  const container = $("#main-content");
  container.innerHTML = `
    <div class="page-header">
      <h1>🔑 卡密管理</h1>
      <div class="header-actions">
        <button class="btn btn-outline" id="btn-export-keys">导出</button>
        <button class="btn btn-primary" id="btn-generate-key">+ 生成卡密</button>
      </div>
    </div>
    <div class="table-container">
      <div class="table-toolbar">
        <div class="left-actions">
          <input type="text" class="search-input" id="key-search" placeholder="搜索卡密或备注..." value="${escapeHtml(keysPageState.search)}">
          <select class="filter-select" id="key-status-filter">
            <option value="">全部状态</option>
            <option value="active" ${keysPageState.status === "active" ? "selected" : ""}>有效</option>
            <option value="disabled" ${keysPageState.status === "disabled" ? "selected" : ""}>已禁用</option>
            <option value="used" ${keysPageState.status === "used" ? "selected" : ""}>已使用</option>
            <option value="expired" ${keysPageState.status === "expired" ? "selected" : ""}>已过期</option>
          </select>
          <div class="bulk-actions" id="bulk-actions" hidden>
            <button class="btn btn-sm btn-danger" id="btn-bulk-delete">删除所选</button>
            <span class="selected-count" id="selected-count">已选 0 项</span>
          </div>
        </div>
      </div>
      <div id="keys-table-container">
        <div class="loading-state"><div class="spinner"></div><p>加载中...</p></div>
      </div>
    </div>

    <!-- 生成卡密弹窗 -->
    <div class="modal-overlay" id="generate-modal">
      <div class="modal">
        <div class="modal-header">
          <h2>生成卡密</h2>
          <button class="modal-close" data-close-modal="generate-modal" aria-label="关闭">×</button>
        </div>
        <div class="modal-body">
          <div class="form-group">
            <label for="gen-count">生成数量（1-100）</label>
            <input type="number" id="gen-count" value="1" min="1" max="100">
          </div>
          <div class="form-group">
            <label for="gen-duration">有效天数（留空 = 永久有效）</label>
            <input type="number" id="gen-duration" placeholder="例如：365" min="1">
          </div>
          <div class="form-group">
            <label for="gen-machines">允许绑定设备数</label>
            <input type="number" id="gen-machines" value="1" min="1" max="999">
          </div>
          <div class="form-group">
            <label for="gen-note">备注</label>
            <input type="text" id="gen-note" placeholder="用于标识卡密用途">
          </div>
          <div id="gen-result" style="display:none;">
            <div class="detail-section">
              <h3>✅ 生成成功！</h3>
              <div id="gen-keys-list" style="font-family:monospace;font-size:13px;word-break:break-all;"></div>
            </div>
          </div>
        </div>
        <div class="modal-actions modal-footer">
          <button class="btn btn-outline" id="btn-gen-cancel">取消</button>
          <button class="btn btn-primary" id="btn-gen-confirm">确认生成</button>
        </div>
      </div>
    </div>

    <!-- 卡密详情弹窗 -->
    <div class="modal-overlay" id="detail-modal">
      <div class="modal">
        <div class="modal-header">
          <h2>卡密详情</h2>
          <button class="modal-close" data-close-modal="detail-modal" aria-label="关闭">×</button>
        </div>
        <div class="modal-body" id="detail-content"><div class="loading-state"><div class="spinner"></div></div></div>
      </div>
    </div>

    <!-- 编辑卡密弹窗 -->
    <div class="modal-overlay" id="edit-modal">
      <div class="modal">
        <div class="modal-header">
          <h2>编辑卡密</h2>
          <button class="modal-close" data-close-modal="edit-modal" aria-label="关闭">×</button>
        </div>
        <div class="modal-body">
          <div class="key-code-display" id="edit-key-code"></div>
          <div class="form-group">
            <label for="edit-status">状态</label>
            <select id="edit-status">
              <option value="active">有效</option>
              <option value="disabled">已禁用</option>
              <option value="expired">已过期</option>
            </select>
          </div>
          <div class="form-group">
            <label for="edit-duration">有效天数（留空 = 永久有效）</label>
            <input type="number" id="edit-duration" min="1">
          </div>
          <div class="form-group">
            <label for="edit-machines">允许绑定设备数</label>
            <input type="number" id="edit-machines" min="1" max="999">
          </div>
          <div class="form-group">
            <label for="edit-note">备注</label>
            <textarea id="edit-note" rows="4" placeholder="用于标识卡密用途"></textarea>
          </div>
        </div>
        <div class="modal-actions modal-footer">
          <button class="btn btn-outline" data-close-modal="edit-modal">取消</button>
          <button class="btn btn-primary" id="btn-edit-confirm">保存</button>
        </div>
      </div>
    </div>

    <!-- 导出卡密弹窗 -->
    <div class="modal-overlay" id="export-modal">
      <div class="modal modal-compact">
        <div class="modal-header">
          <h2>导出卡密</h2>
          <button class="modal-close" data-close-modal="export-modal" aria-label="关闭">×</button>
        </div>
        <div class="modal-body">
          <div class="form-group">
            <label for="export-filter">筛选条件</label>
            <select id="export-filter">
              <option value="unused">未使用卡密</option>
              <option value="active">有效卡密</option>
            </select>
          </div>
        </div>
        <div class="modal-actions modal-footer">
          <button class="btn btn-outline" data-close-modal="export-modal">取消</button>
          <button class="btn btn-primary" id="btn-export-confirm">导出</button>
        </div>
      </div>
    </div>
  `;

  // 加载卡密列表
  await loadKeysTable();

  // 搜索事件（防抖）
  let searchTimer;
  $("#key-search").addEventListener("input", () => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      keysPageState.search = $("#key-search").value;
      keysPageState.page = 1;
      selectedKeyIds.clear();
      loadKeysTable();
    }, 400);
  });

  // 状态筛选
  $("#key-status-filter").addEventListener("change", () => {
    keysPageState.status = $("#key-status-filter").value;
    keysPageState.page = 1;
    selectedKeyIds.clear();
    loadKeysTable();
  });

  // 生成卡密弹窗
  $("#btn-generate-key").addEventListener("click", () => {
    $("#generate-modal").classList.add("open");
    $("#gen-result").style.display = "none";
    $("#gen-count").value = 1;
    $("#gen-duration").value = "";
    $("#gen-machines").value = 1;
    $("#gen-note").value = "";
  });

  $("#btn-gen-cancel").addEventListener("click", () => {
    $("#generate-modal").classList.remove("open");
  });

  $("#btn-gen-confirm").addEventListener("click", handleGenerateKeys);
  $("#btn-export-keys").addEventListener("click", () => {
    $("#export-modal").classList.add("open");
  });
  $("#btn-export-confirm").addEventListener("click", handleExportKeys);
  $("#btn-edit-confirm").addEventListener("click", handleEditKey);
  $("#btn-bulk-delete").addEventListener("click", handleBulkDelete);

  $$("[data-close-modal]").forEach((btn) => {
    btn.addEventListener("click", () => {
      $(`#${btn.dataset.closeModal}`).classList.remove("open");
    });
  });

  // 点击遮罩关闭弹窗
  $$(".modal-overlay").forEach((el) => {
    el.addEventListener("click", (e) => {
      if (e.target === el) el.classList.remove("open");
    });
  });
}

async function loadKeysTable() {
  const container = $("#keys-table-container");
  if (!container) return;

  try {
    const params = new URLSearchParams({
      page: keysPageState.page,
      pageSize: keysPageState.pageSize,
      sortBy: "created_at",
      sortOrder: "desc",
    });
    if (keysPageState.status) params.set("status", keysPageState.status);
    if (keysPageState.search) params.set("search", keysPageState.search);

    const data = await apiRequest(`/api/admin/keys?${params}`);

    if (data.keys.length === 0) {
      container.innerHTML = `<div class="empty-state"><p>暂无卡密数据</p></div>`;
      selectedKeyIds.clear();
      updateSelectedKeysUi();
      return;
    }

    container.innerHTML = `
      <div class="keys-table-scroll">
      <table class="keys-table">
        <thead>
          <tr>
            <th class="col-select sticky-select"><input type="checkbox" id="select-all-keys" aria-label="全选当前页卡密"></th>
            <th class="col-key sticky-key">卡密</th>
            <th class="col-status">状态</th>
            <th class="col-duration">有效期</th>
            <th class="col-machines">已绑定/上限</th>
            <th class="col-online">在线</th>
            <th class="col-note">备注</th>
            <th class="col-created">创建时间</th>
            <th class="col-actions sticky-actions">操作</th>
          </tr>
        </thead>
        <tbody>
          ${data.keys
            .map(
              (key) => `
            <tr>
              <td class="col-select sticky-select"><input type="checkbox" class="key-select" value="${key.id}" ${selectedKeyIds.has(key.id) ? "checked" : ""} aria-label="选择卡密 ${escapeHtml(key.key_code)}"></td>
              <td class="col-key sticky-key key-code-cell">${escapeHtml(key.key_code)}</td>
              <td class="col-status"><span class="badge badge-${key.status}">${STATUS_LABELS[key.status] || key.status}</span></td>
              <td class="col-duration">${key.duration_days ? `${key.duration_days} 天` : "永久"}</td>
              <td class="col-machines">${key.machine_count} / ${key.max_machines}</td>
              <td class="col-online">${key.online_count > 0 ? `<span class="status-ok">● ${key.online_count}</span>` : "—"}</td>
              <td class="col-note ellipsis-cell">${escapeHtml(key.note) || "—"}</td>
              <td class="col-created date-cell">${key.created_at}</td>
              <td class="col-actions sticky-actions">
                <button class="btn btn-sm btn-outline" data-action="detail" data-id="${key.id}">详情</button>
                <button class="btn btn-sm btn-outline" data-action="edit" data-id="${key.id}">编辑</button>
                ${key.status === "active" ? `<button class="btn btn-sm btn-danger" data-action="disable" data-id="${key.id}">禁用</button>` : ""}
                ${key.status === "disabled" ? `<button class="btn btn-sm btn-primary" data-action="enable" data-id="${key.id}">启用</button>` : ""}
                <button class="btn btn-sm btn-danger" data-action="delete" data-id="${key.id}">删除</button>
              </td>
            </tr>
          `
            )
            .join("")}
        </tbody>
      </table>
      </div>
      <div class="pagination">
        <button ${data.pagination.page <= 1 ? "disabled" : ""} data-page="${data.pagination.page - 1}" class="page-btn">上一页</button>
        <span class="page-info">第 ${data.pagination.page} / ${data.pagination.totalPages} 页（共 ${data.pagination.total} 条）</span>
        <button ${data.pagination.page >= data.pagination.totalPages ? "disabled" : ""} data-page="${data.pagination.page + 1}" class="page-btn">下一页</button>
      </div>
    `;

    // 分页事件
    $$(".page-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const page = parseInt(btn.dataset.page, 10);
        if (page > 0) {
          keysPageState.page = page;
          selectedKeyIds.clear();
          loadKeysTable();
        }
      });
    });

    const selectAll = $("#select-all-keys");
    const checkboxes = $$(".key-select");
    if (selectAll) {
      selectAll.addEventListener("change", () => {
        checkboxes.forEach((checkbox) => {
          checkbox.checked = selectAll.checked;
          const id = parseInt(checkbox.value, 10);
          if (selectAll.checked) selectedKeyIds.add(id);
          else selectedKeyIds.delete(id);
        });
        updateSelectedKeysUi();
      });
    }
    checkboxes.forEach((checkbox) => {
      checkbox.addEventListener("change", () => {
        const id = parseInt(checkbox.value, 10);
        if (checkbox.checked) selectedKeyIds.add(id);
        else selectedKeyIds.delete(id);
        updateSelectedKeysUi();
        syncSelectAllState();
      });
    });
    syncSelectAllState();
    updateSelectedKeysUi();

    // 操作按钮事件（委托）
    container.querySelector("table").addEventListener("click", async (e) => {
      const btn = e.target.closest("button[data-action]");
      if (!btn) return;
      const action = btn.dataset.action;
      const id = btn.dataset.id;

      if (action === "detail") {
        await showKeyDetail(id);
      } else if (action === "edit") {
        await showEditKey(id);
      } else if (action === "delete") {
        await handleDeleteKey(id);
      } else if (action === "disable" || action === "enable") {
        const newStatus = action === "disable" ? "disabled" : "active";
        try {
          await apiRequest(`/api/admin/keys/${id}`, {
            method: "PUT",
            body: JSON.stringify({ status: newStatus }),
          });
          showToast(action === "disable" ? "已禁用该卡密" : "已启用该卡密");
          selectedKeyIds.clear();
          await loadKeysTable();
        } catch (err) {
          showToast(err.message, "error");
        }
      }
    });
  } catch (err) {
    container.innerHTML = `<div class="empty-state"><p>❌ 加载失败: ${escapeHtml(err.message)}</p></div>`;
  }
}

function updateSelectedKeysUi() {
  const bulkActions = $("#bulk-actions");
  const bulkBtn = $("#btn-bulk-delete");
  const countEl = $("#selected-count");
  const hasSelection = selectedKeyIds.size > 0;
  if (bulkActions) bulkActions.hidden = !hasSelection;
  if (bulkBtn) bulkBtn.disabled = !hasSelection;
  if (countEl) countEl.textContent = `已选 ${selectedKeyIds.size} 项`;
}

function syncSelectAllState() {
  const selectAll = $("#select-all-keys");
  const checkboxes = Array.from($$(".key-select"));
  if (!selectAll || checkboxes.length === 0) return;
  const checkedCount = checkboxes.filter((checkbox) => checkbox.checked).length;
  selectAll.checked = checkedCount === checkboxes.length;
  selectAll.indeterminate = checkedCount > 0 && checkedCount < checkboxes.length;
}

const STATUS_LABELS = {
  active: "有效",
  disabled: "已禁用",
  used: "已使用",
  expired: "已过期",
};

async function showKeyDetail(id) {
  const modal = $("#detail-modal");
  const content = $("#detail-content");
  modal.classList.add("open");

  try {
    const key = await apiRequest(`/api/admin/keys/${id}`);

    content.innerHTML = `
      <div class="detail-section">
        <div class="key-code-display">${escapeHtml(key.key_code)}</div>
        <div class="detail-grid">
          <div class="detail-item">
            <span class="label">状态</span>
            <span class="value"><span class="badge badge-${key.status}">${STATUS_LABELS[key.status] || key.status}</span></span>
          </div>
          <div class="detail-item">
            <span class="label">有效天数</span>
            <span class="value">${key.duration_days ? `${key.duration_days} 天` : "永久"}</span>
          </div>
          <div class="detail-item">
            <span class="label">最大设备数</span>
            <span class="value">${key.max_machines}</span>
          </div>
          <div class="detail-item">
            <span class="label">已激活次数</span>
            <span class="value">${key.used_count}</span>
          </div>
          <div class="detail-item">
            <span class="label">创建时间</span>
            <span class="value" style="font-size:13px;">${key.created_at}</span>
          </div>
          <div class="detail-item">
            <span class="label">创建者</span>
            <span class="value">${escapeHtml(key.created_by || "—")}</span>
          </div>
          <div class="detail-item" style="grid-column:1/-1;">
            <span class="label">备注</span>
            <span class="value">${escapeHtml(key.note) || "无"}</span>
          </div>
        </div>
      </div>

      <div class="detail-section">
        <h3>📱 已绑定设备（${key.machines.length} 台）</h3>
        ${key.machines.length === 0
          ? '<div class="empty-state"><p>暂无绑定设备</p></div>'
          : `
          <table>
            <thead>
              <tr>
                <th>设备名</th>
                <th>设备 ID</th>
                <th>状态</th>
                <th>在线时长</th>
                <th>最后心跳</th>
                <th>激活时间</th>
              </tr>
            </thead>
            <tbody>
              ${key.machines
                .map(
                  (m) => `
                <tr>
                  <td>${escapeHtml(m.machine_name || "—")}</td>
                  <td style="font-size:12px;font-family:monospace;">${escapeHtml(m.machine_id)}</td>
                  <td>
                    <div class="online-indicator">
                      <span class="online-dot ${m.is_online ? "online" : "offline"}"></span>
                      ${m.is_online ? "在线" : "离线"}
                    </div>
                  </td>
                  <td>${m.online_duration_formatted}</td>
                  <td style="font-size:13px;color:var(--gray-500);">${m.last_heartbeat || "—"}</td>
                  <td style="font-size:13px;color:var(--gray-500);">${m.activated_at}</td>
                </tr>
              `
                )
                .join("")}
            </tbody>
          </table>
        `}
      </div>
    `;
  } catch (err) {
    content.innerHTML = `<div class="empty-state"><p>❌ ${escapeHtml(err.message)}</p></div>`;
  }
}

async function showEditKey(id) {
  const modal = $("#edit-modal");
  modal.dataset.keyId = id;
  modal.classList.add("open");

  try {
    const key = await apiRequest(`/api/admin/keys/${id}`);
    $("#edit-key-code").textContent = key.key_code;
    $("#edit-status").value = key.status === "used" ? "active" : key.status;
    $("#edit-duration").value = key.duration_days || "";
    $("#edit-machines").value = key.max_machines || 1;
    $("#edit-note").value = key.note || "";
  } catch (err) {
    modal.classList.remove("open");
    showToast(err.message, "error");
  }
}

async function handleEditKey() {
  const modal = $("#edit-modal");
  const id = modal.dataset.keyId;
  if (!id) return;

  const duration = $("#edit-duration").value.trim();
  const maxMachines = parseInt($("#edit-machines").value, 10) || 1;
  const payload = {
    status: $("#edit-status").value,
    duration_days: duration ? parseInt(duration, 10) : null,
    max_machines: maxMachines,
    note: $("#edit-note").value.trim(),
  };

  try {
    await apiRequest(`/api/admin/keys/${id}`, {
      method: "PUT",
      body: JSON.stringify(payload),
    });
    modal.classList.remove("open");
    showToast("卡密已更新");
    selectedKeyIds.clear();
    await loadKeysTable();
  } catch (err) {
    showToast(err.message, "error");
  }
}

async function handleDeleteKey(id) {
  const confirmed = await showConfirmDialog({
    title: "删除卡密",
    message: "确定要删除这张卡密吗？关联设备记录也会一并删除。",
    confirmText: "删除",
    danger: true,
  });
  if (!confirmed) return;

  try {
    await apiRequest(`/api/admin/keys/${id}`, { method: "DELETE" });
    selectedKeyIds.delete(parseInt(id, 10));
    showToast("卡密已删除");
    await loadKeysTable();
  } catch (err) {
    showToast(err.message, "error");
  }
}

async function handleBulkDelete() {
  const ids = Array.from(selectedKeyIds);
  if (ids.length === 0) return;
  const confirmed = await showConfirmDialog({
    title: "批量删除",
    message: `确定要删除选中的 ${ids.length} 张卡密吗？关联设备记录也会一并删除。`,
    confirmText: "删除",
    danger: true,
  });
  if (!confirmed) return;

  try {
    const result = await apiRequest("/api/admin/keys/bulk", {
      method: "DELETE",
      body: JSON.stringify({ ids }),
    });
    selectedKeyIds.clear();
    showToast(`已删除 ${result.deleted} 张卡密`);
    await loadKeysTable();
  } catch (err) {
    showToast(err.message, "error");
  }
}

async function handleExportKeys() {
  const filter = $("#export-filter").value;
  const headers = {};
  if (STATE.token) headers["Authorization"] = `Bearer ${STATE.token}`;

  try {
    const response = await fetch(`${API}/api/admin/keys/export?filter=${encodeURIComponent(filter)}`, { headers });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.error || `导出失败 (${response.status})`);
    }

    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const filenameMatch = /filename="?([^"]+)"?/i.exec(disposition);
    const filename = filenameMatch ? filenameMatch[1] : `license-keys-${filter}.txt`;
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    $("#export-modal").classList.remove("open");
    showToast("卡密导出已开始");
  } catch (err) {
    showToast(err.message, "error");
  }
}

async function handleGenerateKeys() {
  const count = parseInt($("#gen-count").value, 10) || 1;
  const durationDays = parseInt($("#gen-duration").value, 10) || null;
  const maxMachines = parseInt($("#gen-machines").value, 10) || 1;
  const note = $("#gen-note").value.trim();
  const resultDiv = $("#gen-result");
  const keysList = $("#gen-keys-list");

  try {
    const data = await apiRequest("/api/admin/keys", {
      method: "POST",
      body: JSON.stringify({ count, duration_days: durationDays, max_machines: maxMachines, note }),
    });

    resultDiv.style.display = "block";
    keysList.innerHTML = `<br>${data.keys.map((k) => `🔑 ${k}`).join("<br>")}<br><br><small style="color:var(--gray-400);">共生成 ${data.count} 个卡密</small>`;
    showToast(`成功生成 ${data.count} 个卡密`);
    await loadKeysTable();
  } catch (err) {
    showToast(err.message, "error");
  }
}

// ============================================================
// 在线监控
// ============================================================

let monitorInterval = null;

async function renderMonitor() {
  // 清除旧定时器
  if (monitorInterval) {
    clearInterval(monitorInterval);
    monitorInterval = null;
  }

  const container = $("#main-content");
  container.innerHTML = `
    <div class="page-header">
      <h1>📡 在线监控</h1>
      <span style="font-size:13px;color:var(--gray-500);">自动刷新中 <span id="refresh-countdown">30</span>s</span>
    </div>
    <div class="table-container">
      <div class="auto-refresh-bar">
        <span id="online-summary">加载中...</span>
        <button class="btn btn-sm btn-outline" id="btn-refresh-now">立即刷新</button>
      </div>
      <div id="monitor-table-container">
        <div class="loading-state"><div class="spinner"></div><p>加载中...</p></div>
      </div>
    </div>
  `;

  await loadMonitorData();

  // 立即刷新
  $("#btn-refresh-now").addEventListener("click", async () => {
    await loadMonitorData();
    countdown = 30;
  });

  // 自动刷新（30秒）
  let countdown = 30;
  monitorInterval = setInterval(async () => {
    countdown--;
    const el = $("#refresh-countdown");
    if (el) el.textContent = countdown;
    if (countdown <= 0) {
      await loadMonitorData();
      countdown = 30;
    }
  }, 1000);
}

async function loadMonitorData() {
  const container = $("#monitor-table-container");
  const summary = $("#online-summary");
  if (!container) return;

  try {
    // 同时获取在线设备和全部设备
    const [onlineData, allData] = await Promise.all([
      apiRequest("/api/admin/machines/online"),
      apiRequest("/api/admin/machines?page=1&pageSize=100"),
    ]);

    if (summary) {
      summary.textContent = `🟢 ${onlineData.count} 台在线 · 共 ${allData.pagination.total} 台已激活设备`;
    }

    // 在线设备列表
    const onlineMachines = onlineData.machines || [];
    const allMachines = allData.machines || [];

    if (allMachines.length === 0) {
      container.innerHTML = `<div class="empty-state"><p>暂无已激活设备</p></div>`;
      return;
    }

    container.innerHTML = `
      <table>
        <thead>
          <tr>
            <th>状态</th>
            <th>卡密</th>
            <th>设备名称</th>
            <th>设备 ID</th>
            <th>累计在线时长</th>
            <th>最后心跳</th>
            <th>激活时间</th>
          </tr>
        </thead>
        <tbody>
          ${allMachines
            .map(
              (m) => `
            <tr>
              <td>
                <div class="online-indicator">
                  <span class="online-dot ${m.is_online ? "online" : "offline"}"></span>
                  ${m.is_online ? "在线" : "离线"}
                </div>
              </td>
              <td style="font-family:monospace;font-size:13px;">${escapeHtml(m.key_code)}</td>
              <td>${escapeHtml(m.machine_name || "—")}</td>
              <td style="font-size:12px;font-family:monospace;color:var(--gray-500);">${escapeHtml(m.machine_id)}</td>
              <td>${m.online_duration_formatted}</td>
              <td style="font-size:13px;color:var(--gray-500);">
                ${m.last_heartbeat || "—"}
                ${m.last_heartbeat_ago ? `<br><span style="font-size:11px;color:var(--gray-400);">${m.last_heartbeat_ago}</span>` : ""}
              </td>
              <td style="font-size:13px;color:var(--gray-500);">${m.activated_at || "—"}</td>
            </tr>
          `
            )
            .join("")}
        </tbody>
      </table>
    `;
  } catch (err) {
    if (container) {
      container.innerHTML = `<div class="empty-state"><p>❌ 加载失败: ${escapeHtml(err.message)}</p></div>`;
    }
  }
}

// ============================================================
// 辅助
// ============================================================

function formatDuration(totalSeconds) {
  if (!totalSeconds || totalSeconds <= 0) return "0 分钟";
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  if (hours > 0) return `${hours} 小时 ${minutes} 分钟`;
  return `${minutes} 分钟`;
}

// ============================================================
// 应用初始化
// ============================================================

function renderApp() {
  if (STATE.token) {
    renderAdminLayout();
  } else {
    renderLogin();
  }
}

// 启动
renderApp();
