/**
 * 管理后台 API 路由
 *
 * 所有路由均需 JWT 认证（requireAdmin 中间件），除 /login 外。
 * 提供：管理员登录、仪表盘统计、卡密管理（CRUD/生成/导出）、设备管理（列表/状态）。
 *
 * 安全约束：
 * - 密码用 bcrypt 哈希比对，不存储明文
 * - 所有写操作先验证卡密存在性再执行
 * - 批量操作使用事务保证原子性
 * - 排序字段白名单防止 SQL 注入
 * - 卡密 ID 需为正整数（路由参数和批量 ID 统一做类型校验）
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

/**
 * POST /api/admin/login
 *
 * 登录流程：
 * 1. 根据 username 查找管理员
 * 2. 用 bcrypt.compareSync 比对密码哈希
 * 3. 验证通过后签发 JWT，返回给客户端
 *
 * 安全设计：
 * - 用户名或密码错误均返回相同提示（"用户名或密码错误"），防止用户名字典攻击
 * - 不暴露具体是用户名错误还是密码错误
 */
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

/**
 * GET /api/admin/dashboard
 *
 * 聚合查询说明：
 * - 基本的计数查询均走索引（status、is_online），性能良好
 * - 今日新增卡密/激活使用 date() 函数，查询当天数据
 * - 近 7 天激活趋势使用 GROUP BY + datetime('now', '-7 days') 窗口查询
 * - 在线时长排行使用 JOIN 获取卡密代码，LIMIT 10 控制返回量
 *
 * 这些查询在一次请求中并行执行（但由于 better-sqlite3 是同步的，实际为串行），
 * 考虑到数据量很小（< 10 万条），性能可接受。
 */
router.get("/dashboard", requireAdmin, (req, res) => {
  const db = getDb();

  // 各类汇总计数
  const totalKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys").get();
  const activeKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys WHERE status = 'active'").get();
  const disabledKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys WHERE status = 'disabled'").get();
  const usedKeys = db.prepare("SELECT COUNT(*) as count FROM license_keys WHERE used_count > 0").get();
  const onlineNow = db.prepare("SELECT COUNT(*) as count FROM activated_machines WHERE is_online = 1").get();
  const totalMachines = db.prepare("SELECT COUNT(*) as count FROM activated_machines").get();

  // 今日新增卡密（基于 created_at 的日期部分）
  const todayNewKeys = db.prepare(
    "SELECT COUNT(*) as count FROM license_keys WHERE date(created_at) = date('now')"
  ).get();

  // 今日新增激活（基于 activated_at 的日期部分）
  const todayActivations = db.prepare(
    "SELECT COUNT(*) as count FROM activated_machines WHERE date(activated_at) = date('now')"
  ).get();

  // 最近 7 天激活趋势（按天分组，用于前端图表展示）
  const weeklyTrend = db.prepare(`
    SELECT date(activated_at) as day, COUNT(*) as count
    FROM activated_machines
    WHERE activated_at >= datetime('now', '-7 days')
    GROUP BY date(activated_at)
    ORDER BY day
  `).all();

  // 在线时长排行（Top 10，用于展示最活跃的设备）
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
// 卡密列表（动态过滤 / 分页 / 排序）
// ============================================================

/**
 * GET /api/admin/keys
 *
 * 动态查询构建策略：
 * - 通过 status 和 search 两个可选参数构建 WHERE 子句
 * - search 同时匹配 key_code 和 note（LIKE 模糊查询）
 * - sortBy 白名单校验，防止 SQL 注入（不允许任意字段排序）
 * - 分页使用 LIMIT + OFFSET
 *
 * 子查询说明：
 * - machine_count / online_count 使用相关子查询获取实时计数
 * - 这比在应用层循环查询更高效
 */
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

  // 安全解析分页参数
  const pageNum = Math.max(1, parseInt(page, 10) || 1);
  const limit = Math.min(100, Math.max(1, parseInt(pageSize, 10) || 20));
  const offset = (pageNum - 1) * limit;

  // 动态构建 WHERE 条件
  const conditions = [];
  const params = [];

  if (status && ["active", "disabled", "used", "expired"].includes(status)) {
    conditions.push("lk.status = ?");
    params.push(status);
  }

  if (search) {
    // 同时模糊匹配卡密和备注
    conditions.push("(lk.key_code LIKE ? OR lk.note LIKE ?)");
    params.push(`%${search}%`, `%${search}%`);
  }

  const where = conditions.length > 0 ? `WHERE ${conditions.join(" AND ")}` : "";

  // 排序字段白名单：防止通过 sortBy 注入 SQL
  const allowedSort = ["created_at", "key_code", "status", "used_count", "updated_at"];
  const sortField = allowedSort.includes(sortBy) ? sortBy : "created_at";
  const sortDir = sortOrder === "asc" ? "ASC" : "DESC";

  // 总数查询（用于前端分页）
  const total = db.prepare(`SELECT COUNT(*) as count FROM license_keys lk ${where}`).get(...params);

  // 主数据查询 + 关联计数子查询
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

/**
 * POST /api/admin/keys
 *
 * 生成规则：
 * - 格式：XXXX-XXXX-XXXX-XXXX（4 段，每段 4 个大写字母/数字）
 * - 字符集排除易混淆字符：I、O、0、1 —> 减少人工输入错误
 * - 一次最多生成 100 张（防止误操作导致大量生成）
 * - 使用事务包裹：要么全部生成成功，要么全部回滚
 * - 每张卡密独立调用 INSERT，使用单条 prepared statement 复用
 */
router.post("/keys", requireAdmin, (req, res) => {
  const { count = 1, duration_days = null, max_machines = 1, note = "" } = req.body || {};

  // 限制单次生成数量，防止误操作
  const generateCount = Math.min(100, Math.max(1, parseInt(count, 10) || 1));

  const db = getDb();
  const insert = db.prepare(`
    INSERT INTO license_keys (key_code, duration_days, max_machines, note, created_by)
    VALUES (?, ?, ?, ?, ?)
  `);

  const generated = [];

  // 事务保证批量生成的原子性
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

/**
 * GET /api/admin/keys/export
 *
 * 以纯文本文件形式导出卡密，一行一个。
 * 支持按状态筛选（active 有效 / unused 未使用）。
 * 文件名含日期和筛选标签，方便归档管理。
 *
 * 响应为 text/plain 文件下载，浏览器触发下载对话框。
 */
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

/**
 * DELETE /api/admin/keys/bulk
 *
 * 批量删除流程：
 * 1. 接收 ids 数组，通过 normalizeKeyIds 做清洗（去重、过滤非正整数）
 * 2. 先删除关联设备记录（外键约束需要或避免孤立数据）
 * 3. 再删除卡密本身
 * 4. 使用事务保证两步删除的原子性
 */
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
 * 生成单张卡密字符串
 *
 * 字符集说明：
 * 使用 ABCDEFGHJKLMNPQRSTUVWXYZ23456789（共 30 个字符），
 * 去除了 I、O、0、1 这些容易被混淆的字符，减少用户输入错误。
 *
 * 格式：XXXX-XXXX-XXXX-XXXX，每段 4 个随机字符，共 16 位。
 * 通过 crypto.randomInt 生成密码学安全的随机数。
 *
 * @returns {string} 格式化的卡密字符串
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

/**
 * GET /api/admin/keys/:id
 *
 * 返回卡密基本信息 + 所有已绑定设备列表。
 * 设备列表附带在线状态判断（通过 isOnline 函数）、
 * 累计在线时长格式化、最后心跳相对时间。
 */
router.get("/keys/:id", requireAdmin, (req, res) => {
  const db = getDb();
  const key = db.prepare("SELECT * FROM license_keys WHERE id = ?").get(req.params.id);

  if (!key) {
    return res.status(404).json({ error: "卡密不存在" });
  }

  const machines = db.prepare(`
    SELECT * FROM activated_machines WHERE key_id = ? ORDER BY activated_at DESC
  `).all(key.id);

  // 格式化设备数据：添加可读的在线时长、在线状态、北京时区时间
  const machinesFormatted = machines.map((m) => ({
    ...withBeijingDateTimes(m),
    online_duration_formatted: formatDuration(m.total_online_seconds),
    is_online: isOnline(m.last_heartbeat),
  }));

  res.json({ ...withBeijingDateTimes(key), machines: machinesFormatted });
});

// ============================================================
// 更新卡密（禁用/启用/修改备注/有效期/设备上限）
// ============================================================

/**
 * PUT /api/admin/keys/:id
 *
 * 动态更新策略：
 * - 只更新请求中包含的字段，未提供的字段保持不变
 * - 支持更新：status、note、duration_days、max_machines
 * - status 仅允许设置为 active / disabled / expired（不可设为 used，used 是自动状态）
 * - 自动更新 updated_at 时间戳
 */
router.put("/keys/:id", requireAdmin, (req, res) => {
  const { status, note, duration_days, max_machines } = req.body || {};
  const db = getDb();

  const key = db.prepare("SELECT * FROM license_keys WHERE id = ?").get(req.params.id);
  if (!key) {
    return res.status(404).json({ error: "卡密不存在" });
  }

  // 动态构建 SET 子句
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

  // 自动更新修改时间
  updates.push("updated_at = datetime('now')");
  params.push(key.id);

  db.prepare(`UPDATE license_keys SET ${updates.join(", ")} WHERE id = ?`).run(...params);

  const updated = db.prepare("SELECT * FROM license_keys WHERE id = ?").get(key.id);
  res.json(withBeijingDateTimes(updated));
});

// ============================================================
// 删除单个卡密
// ============================================================

/**
 * DELETE /api/admin/keys/:id
 *
 * 与批量删除类似，先删关联设备再删卡密。
 * 对路由参数做了严格的类型校验（必须为正整数），
 * 防止 SQL 注入或无效参数导致意外行为。
 */
router.delete("/keys/:id", requireAdmin, (req, res) => {
  // 严格校验 ID 格式：必须为正整数
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
// 设备列表（分页 + 在线状态过滤）
// ============================================================

/**
 * GET /api/admin/machines
 *
 * 设备在线状态判断逻辑（区别于 is_online 冗余字段）：
 * 此处采用实时计算方式：通过 (unixepoch('now') - unixepoch(m.last_heartbeat)) < heartbeatTimeout
 * 来判断设备是否在线。这比依赖 is_online 字段更准确，因为：
 * 1. is_online 仅在心跳上报时更新为 1，不会自动变为 0
 * 2. 实时计算能反映超过 heartbeatTimeout 未上报的真实离线状态
 *
 * 过滤参数：
 * - online=1: 只显示当前在线设备
 * - online=0: 只显示当前离线设备
 * - 不传: 显示全部
 *
 * 返回字段补充：
 * - seconds_since_heartbeat: 距离上次心跳的秒数（用于前端显示相对时间）
 * - online_duration_formatted: 累计在线时长的可读格式
 * - last_heartbeat_ago: 最后心跳的相对时间描述
 */
router.get("/machines", requireAdmin, (req, res) => {
  const db = getDb();
  const { page = "1", pageSize = "20", online } = req.query;

  const pageNum = Math.max(1, parseInt(page, 10) || 1);
  const limit = Math.min(100, Math.max(1, parseInt(pageSize, 10) || 20));
  const offset = (pageNum - 1) * limit;

  const conditions = [];
  const params = [];

  // 在线状态过滤：使用实时计算而非 is_online 字段
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

/**
 * GET /api/admin/machines/online
 *
 * 仪表盘/监控页使用的精简在线设备列表。
 * 只返回当前在线设备（last_heartbeat 在 heartbeatTimeout 内）。
 * 相比 /machines 接口，不包含分页和离线设备，数据量更小。
 */
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
 * 判断设备是否在线
 *
 * 基于最后心跳时间与当前时间的差值是否小于 heartbeatTimeout。
 * 与服务端 /machines 接口中的 SQL 实时计算逻辑保持一致。
 *
 * @param {string|null} lastHeartbeat - SQLite datetime 字符串或 null
 * @returns {boolean}
 */
function isOnline(lastHeartbeat) {
  if (!lastHeartbeat) return false;
  const now = Math.floor(Date.now() / 1000);
  const heartbeat = Math.floor(new Date(`${lastHeartbeat}Z`).getTime() / 1000);
  return (now - heartbeat) < config.heartbeatTimeout;
}

/**
 * 格式化秒数为可读时长
 *
 * @param {number|null} totalSeconds - 累计在线秒数
 * @returns {string} 格式如 "2 小时 30 分钟" 或 "45 分钟"
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
 * 格式化秒数差为相对时间文本
 *
 * @param {number} seconds - 距现在的秒数
 * @returns {string} 如 "3 分钟前"、"2 小时前"、"1 天前"
 */
function formatTimeAgo(seconds) {
  if (seconds < 60) return `${seconds} 秒前`;
  if (seconds < 3600) return `${Math.floor(seconds / 60)} 分钟前`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)} 小时前`;
  return `${Math.floor(seconds / 86400)} 天前`;
}

export default router;
