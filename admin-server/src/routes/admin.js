/**
 * 管理后台 API 路由
 *
 * 所有路由均需 JWT 认证（requireAdmin 中间件）。
 * 提供：登录、仪表盘、卡密管理、设备管理。
 */

import { Router } from "express";
import bcrypt from "bcryptjs";
import crypto from "crypto";
import { getDb } from "../db.js";
import { signToken, requireAdmin } from "../middleware/auth.js";
import config from "../config.js";
import { withBeijingDateTimes } from "../time.js";
import { buildExportFilter, normalizeKeyIds } from "../keyAdmin.js";

const router = Router();

// ============================================================
// 管理员登录（无需认证）
// ============================================================

router.post("/login", (req, res) => {
  const { username, password } = req.body || {};

  if (!username || !password) {
    return res.status(400).json({ error: "请输入用户名和密码" });
  }

  const db = getDb();
  const user = db.prepare("SELECT * FROM admin_users WHERE username = ?").get(username);

  if (!user) {
    return res.status(401).json({ error: "用户名或密码错误" });
  }

  const valid = bcrypt.compareSync(password, user.password_hash);
  if (!valid) {
    return res.status(401).json({ error: "用户名或密码错误" });
  }

  const token = signToken({ id: user.id, username: user.username });
  res.json({ token, username: user.username });
});

// ============================================================
// 仪表盘统计数据
// ============================================================

router.get("/dashboard", requireAdmin, (req, res) => {
  const db = getDb();

  const totalKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys").get();
  const activeKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys WHERE status = 'active'").get();
  const disabledKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys WHERE status = 'disabled'").get();
  const usedKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys WHERE used_count > 0").get();
  const onlineNow = db.prepare("SELECT COUNT(*) as count FROM activated_machines WHERE is_online = 1").get();
  const totalMachines = db.prepare("SELECT COUNT(*) as count FROM activated_machines").get();

  // 今日新增卡密
  const todayNewKeys = db.prepare(
    "SELECT COUNT(*) as count FROM license_keys WHERE date(created_at) = date('now')"
  ).get();

  // 今日新增激活
  const todayActivations = db.prepare(
    "SELECT COUNT(*) as count FROM activated_machines WHERE date(activated_at) = date('now')"
  ).get();

  // 最近 7 天激活趋势
  const weeklyTrend = db.prepare(`
    SELECT date(activated_at) as day, COUNT(*) as count
    FROM activated_machines
    WHERE activated_at >= datetime('now', '-7 days')
    GROUP BY date(activated_at)
    ORDER BY day
  `).all();

  // 在线时长排行（Top 10）
  const topOnline = db.prepare(`
    SELECT m.total_online_seconds, m.machine_name, m.machine_id, lk.key_code
    FROM activated_machines m
    JOIN license_keys lk ON m.key_id = lk.id
    ORDER BY m.total_online_seconds DESC
    LIMIT 10
  `).all();

  res.json({
    totalKeys: totalKeys.count,
    activeKeys: activeKeys.count,
    disabledKeys: disabledKeys.count,
    usedKeys: usedKeys.count,
    onlineNow: onlineNow.count,
    totalMachines: totalMachines.count,
    todayNewKeys: todayNewKeys.count,
    todayActivations: todayActivations.count,
    weeklyTrend,
    topOnline,
  });
});

// ============================================================
// 卡密列表
// ============================================================

router.get("/keys", requireAdmin, (req, res) => {
  const db = getDb();
  const {
    page = "1",
    pageSize = "20",
    status,
    search,
    sortBy = "created_at",
    sortOrder = "desc",
  } = req.query;

  const pageNum = Math.max(1, parseInt(page, 10) || 1);
  const limit = Math.min(100, Math.max(1, parseInt(pageSize, 10) || 20));
  const offset = (pageNum - 1) * limit;

  // 构建查询
  const conditions = [];
  const params = [];

  if (status && ["active", "disabled", "used", "expired"].includes(status)) {
    conditions.push("lk.status = ?");
    params.push(status);
  }

  if (search) {
    conditions.push("(lk.key_code LIKE ? OR lk.note LIKE ?)");
    params.push(`%${search}%`, `%${search}%`);
  }

  const where = conditions.length > 0 ? `WHERE ${conditions.join(" AND ")}` : "";

  // 允许的排序字段
  const allowedSort = ["created_at", "key_code", "status", "used_count", "updated_at"];
  const sortField = allowedSort.includes(sortBy) ? sortBy : "created_at";
  const sortDir = sortOrder === "asc" ? "ASC" : "DESC";

  const total = db.prepare(`SELECT COUNT(*) as count FROM license_keys lk ${where}`).get(...params);

  const keys = db.prepare(`
    SELECT lk.*,
      (SELECT COUNT(*) FROM activated_machines WHERE key_id = lk.id) as machine_count,
      (SELECT COUNT(*) FROM activated_machines WHERE key_id = lk.id AND is_online = 1) as online_count
    FROM license_keys lk
    ${where}
    ORDER BY lk.${sortField} ${sortDir}
    LIMIT ? OFFSET ?
  `).all(...params, limit, offset);

  res.json({
    keys: keys.map(withBeijingDateTimes),
    pagination: {
      page: pageNum,
      pageSize: limit,
      total: total.count,
      totalPages: Math.ceil(total.count / limit),
    },
  });
});

// ============================================================
// 生成卡密
// ============================================================

router.post("/keys", requireAdmin, (req, res) => {
  const { count = 1, duration_days = null, max_machines = 1, note = "" } = req.body || {};

  const generateCount = Math.min(100, Math.max(1, parseInt(count, 10) || 1));

  const db = getDb();
  const insert = db.prepare(`
    INSERT INTO license_keys (key_code, duration_days, max_machines, note, created_by)
    VALUES (?, ?, ?, ?, ?)
  `);

  const generated = [];

  const transaction = db.transaction(() => {
    for (let i = 0; i < generateCount; i++) {
      const keyCode = generateKeyCode();
      insert.run(keyCode, duration_days || null, max_machines, note, req.admin.username);
      generated.push(keyCode);
    }
  });

  transaction();

  res.json({
    success: true,
    count: generated.length,
    keys: generated,
  });
});

// ============================================================
// 导出卡密
// ============================================================

router.get("/keys/export", requireAdmin, (req, res) => {
  const db = getDb();
  const { label, where } = buildExportFilter(req.query.filter);
  const rows = db.prepare(`
    SELECT key_code
    FROM license_keys
    ${where}
    ORDER BY created_at DESC
  `).all();

  const content = rows.map((row) => row.key_code).join("\n");
  const stamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
  res.setHeader("Content-Type", "text/plain; charset=utf-8");
  res.setHeader("Content-Disposition", `attachment; filename=\"license-keys-${label}-${stamp}.txt\"`);
  res.send(content ? `${content}\n` : "");
});

// ============================================================
// 批量删除卡密
// ============================================================

router.delete("/keys/bulk", requireAdmin, (req, res) => {
  const ids = normalizeKeyIds(req.body?.ids);
  if (ids.length === 0) {
    return res.status(400).json({ error: "请选择要删除的卡密" });
  }

  const db = getDb();
  const placeholders = ids.map(() => "?").join(", ");
  const transaction = db.transaction(() => {
    db.prepare(`DELETE FROM activated_machines WHERE key_id IN (${placeholders})`).run(...ids);
    return db.prepare(`DELETE FROM license_keys WHERE id IN (${placeholders})`).run(...ids);
  });
  const result = transaction();

  res.json({ success: true, deleted: result.changes });
});

/**
 * 生成卡密：XXXX-XXXX-XXXX-XXXX（大写字母 + 数字）
 */
function generateKeyCode() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // 去除易混淆字符 I/O/0/1
  const segments = [];
  for (let seg = 0; seg < 4; seg++) {
    let segment = "";
    for (let i = 0; i < 4; i++) {
      segment += chars[crypto.randomInt(0, chars.length)];
    }
    segments.push(segment);
  }
  return segments.join("-");
}

// ============================================================
// 获取卡密详情（含关联设备）
// ============================================================

router.get("/keys/:id", requireAdmin, (req, res) => {
  const db = getDb();
  const key = db.prepare("SELECT * FROM license_keys WHERE id = ?").get(req.params.id);

  if (!key) {
    return res.status(404).json({ error: "卡密不存在" });
  }

  const machines = db.prepare(`
    SELECT * FROM activated_machines WHERE key_id = ? ORDER BY activated_at DESC
  `).all(key.id);

  // 计算累计在线时长（格式化）
  const machinesFormatted = machines.map((m) => ({
    ...withBeijingDateTimes(m),
    online_duration_formatted: formatDuration(m.total_online_seconds),
    is_online: isOnline(m.last_heartbeat),
  }));

  res.json({ ...withBeijingDateTimes(key), machines: machinesFormatted });
});

// ============================================================
// 更新卡密（禁用/启用/修改备注）
// ============================================================

router.put("/keys/:id", requireAdmin, (req, res) => {
  const { status, note, duration_days, max_machines } = req.body || {};
  const db = getDb();

  const key = db.prepare("SELECT * FROM license_keys WHERE id = ?").get(req.params.id);
  if (!key) {
    return res.status(404).json({ error: "卡密不存在" });
  }

  const updates = [];
  const params = [];

  if (status && ["active", "disabled", "expired"].includes(status)) {
    updates.push("status = ?");
    params.push(status);
  }
  if (note !== undefined) {
    updates.push("note = ?");
    params.push(note);
  }
  if (duration_days !== undefined) {
    updates.push("duration_days = ?");
    params.push(duration_days);
  }
  if (max_machines !== undefined) {
    updates.push("max_machines = ?");
    params.push(max_machines);
  }

  if (updates.length === 0) {
    return res.status(400).json({ error: "未提供需要更新的字段" });
  }

  updates.push("updated_at = datetime('now')");
  params.push(key.id);

  db.prepare(`UPDATE license_keys SET ${updates.join(", ")} WHERE id = ?`).run(...params);

  const updated = db.prepare("SELECT * FROM license_keys WHERE id = ?").get(key.id);
  res.json(withBeijingDateTimes(updated));
});

// ============================================================
// 删除单个卡密
// ============================================================

router.delete("/keys/:id", requireAdmin, (req, res) => {
  const id = Math.trunc(Number(req.params.id));
  if (!Number.isInteger(id) || id <= 0) {
    return res.status(400).json({ error: "卡密 ID 无效" });
  }

  const db = getDb();
  const key = db.prepare("SELECT id FROM license_keys WHERE id = ?").get(id);
  if (!key) {
    return res.status(404).json({ error: "卡密不存在" });
  }

  const transaction = db.transaction(() => {
    db.prepare("DELETE FROM activated_machines WHERE key_id = ?").run(id);
    return db.prepare("DELETE FROM license_keys WHERE id = ?").run(id);
  });
  const result = transaction();

  res.json({ success: true, deleted: result.changes });
});

// ============================================================
// 设备列表
// ============================================================

router.get("/machines", requireAdmin, (req, res) => {
  const db = getDb();
  const { page = "1", pageSize = "20", online } = req.query;

  const pageNum = Math.max(1, parseInt(page, 10) || 1);
  const limit = Math.min(100, Math.max(1, parseInt(pageSize, 10) || 20));
  const offset = (pageNum - 1) * limit;

  const conditions = [];
  const params = [];

  if (online === "1") {
    params.push(config.heartbeatTimeout);
    conditions.push(`m.last_heartbeat IS NOT NULL AND (unixepoch('now') - unixepoch(m.last_heartbeat)) < ?`);
  } else if (online === "0") {
    params.push(config.heartbeatTimeout);
    conditions.push(`(m.last_heartbeat IS NULL OR (unixepoch('now') - unixepoch(m.last_heartbeat)) >= ?)`);
  }

  const where = conditions.length > 0 ? `WHERE ${conditions.join(" AND ")}` : "";

  const total = db.prepare(`
    SELECT COUNT(*) as count FROM activated_machines m ${where}
  `).get(...params);

  const machines = db.prepare(`
    SELECT m.*, lk.key_code,
      CASE WHEN m.last_heartbeat IS NOT NULL AND (unixepoch('now') - unixepoch(m.last_heartbeat)) < ${config.heartbeatTimeout} THEN 1 ELSE 0 END as is_online,
      CASE WHEN m.last_heartbeat IS NOT NULL THEN CAST((unixepoch('now') - unixepoch(m.last_heartbeat)) AS INTEGER) ELSE NULL END as seconds_since_heartbeat
    FROM activated_machines m
    JOIN license_keys lk ON m.key_id = lk.id
    ${where}
    ORDER BY m.last_heartbeat DESC NULLS LAST
    LIMIT ? OFFSET ?
  `).all(...params, limit, offset);

  res.json({
    machines: machines.map((m) => ({
      ...withBeijingDateTimes(m),
      online_duration_formatted: formatDuration(m.total_online_seconds),
      last_heartbeat_ago: m.seconds_since_heartbeat != null
        ? formatTimeAgo(m.seconds_since_heartbeat)
        : "从未上报",
    })),
    pagination: {
      page: pageNum,
      pageSize: limit,
      total: total.count,
      totalPages: Math.ceil(total.count / limit),
    },
  });
});

// ============================================================
// 在线设备（简化版，用于监控页快捷查看）
// ============================================================

router.get("/machines/online", requireAdmin, (_req, res) => {
  const db = getDb();

  const onlineMachines = db.prepare(`
    SELECT m.id, m.machine_name, m.machine_id, m.last_heartbeat, m.total_online_seconds, lk.key_code
    FROM activated_machines m
    JOIN license_keys lk ON m.key_id = lk.id
    WHERE m.last_heartbeat IS NOT NULL
      AND (unixepoch('now') - unixepoch(m.last_heartbeat)) < ?
    ORDER BY m.last_heartbeat DESC
  `).all(config.heartbeatTimeout);

  res.json({
    count: onlineMachines.length,
    machines: onlineMachines.map((m) => ({
      ...withBeijingDateTimes(m),
      online_duration_formatted: formatDuration(m.total_online_seconds),
    })),
  });
});

// ============================================================
// 辅助函数
// ============================================================

/**
 * 判断设备是否在线（基于最后心跳时间）
 */
function isOnline(lastHeartbeat) {
  if (!lastHeartbeat) return false;
  const now = Math.floor(Date.now() / 1000);
  const heartbeat = Math.floor(new Date(`${lastHeartbeat}Z`).getTime() / 1000);
  return (now - heartbeat) < config.heartbeatTimeout;
}

/**
 * 格式化秒数为可读时长
 */
function formatDuration(totalSeconds) {
  if (!totalSeconds || totalSeconds <= 0) return "0 分钟";
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  if (hours > 0) {
    return `${hours} 小时 ${minutes} 分钟`;
  }
  return `${minutes} 分钟`;
}

/**
 * 格式化秒数差为可读文本
 */
function formatTimeAgo(seconds) {
  if (seconds < 60) return `${seconds} 秒前`;
  if (seconds < 3600) return `${Math.floor(seconds / 60)} 分钟前`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)} 小时前`;
  return `${Math.floor(seconds / 86400)} 天前`;
}

export default router;
