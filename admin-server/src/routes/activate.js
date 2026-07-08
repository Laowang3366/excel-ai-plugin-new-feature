/**
 * 激活验证路由（公开接口，无需 JWT）
 *
 * POST /api/activate
 *
 * 激活流程：
 * 1. 卡密标准化（去空格、转大写）— 消除用户输入差异
 * 2. 验证卡密存在性 → 不存在返回 404
 * 3. 验证卡密状态（disabled / expired）→ 异常状态返回 403
 * 4. 检查是否已绑定该设备 → 已绑定则续心跳后直接返回成功（幂等）
 * 5. 检查已绑定设备数是否达到 max_machines → 达到上限返回 403
 * 6. 写入设备绑定记录，更新卡密 used_count
 * 7. 若 used_count >= max_machines，自动将卡密标记为 used
 * 8. 返回激活结果（含有效期等元信息）
 *
 * 安全性考虑：
 * - 卡密验证仅依赖 key_code，不验证请求来源 IP（适用于插件嵌入场景）
 * - machine_id 由客户端生成并持久化，服务端不验证其格式
 * - 设备绑定数受 max_machines 限制，防止一张卡密被无限扩散
 * - 已绑定设备重复激活不会增加 used_count
 */

import { Router } from "express";
import { getDb } from "../db.js";
import config from "../config.js";

const router = Router();

router.post("/", (req, res) => {
  const { key, machine_id, machine_name } = req.body || {};

  // ---------- 参数校验 ----------
  if (!key) {
    return res.status(400).json({ success: false, error: "请输入卡密" });
  }
  if (!machine_id) {
    return res.status(400).json({ success: false, error: "缺少设备标识" });
  }

  // 标准化卡密：去除首尾空格 + 转大写，消除人工输入时的常见差异
  const normalizedKey = key.trim().toUpperCase();

  const db = getDb();

  // ---------- 查询卡密 ----------
  const licenseKey = db.prepare("SELECT * FROM license_keys WHERE key_code = ?").get(normalizedKey);

  if (!licenseKey) {
    return res.status(404).json({ success: false, error: "卡密不存在" });
  }

  // ---------- 状态校验 ----------
  if (licenseKey.status === "disabled") {
    return res.status(403).json({ success: false, error: "卡密已被禁用" });
  }

  if (licenseKey.status === "expired") {
    return res.status(403).json({ success: false, error: "卡密已过期" });
  }

  // ---------- 重复激活检测 ----------
  // 如果该卡密下已存在相同 machine_id 的记录，说明设备已激活过。
  // 此时不做重复绑定，仅更新心跳时间并返回成功（幂等设计）。
  const existingMachine = db.prepare(
    "SELECT * FROM activated_machines WHERE key_id = ? AND machine_id = ?"
  ).get(licenseKey.id, machine_id);

  if (existingMachine) {
    // 续心跳、标记在线
    db.prepare(`
      UPDATE activated_machines
      SET last_heartbeat = datetime('now'), is_online = 1
      WHERE id = ?
    `).run(existingMachine.id);

    return res.json({
      success: true,
      activated: true,
      key: licenseKey.key_code,
      message: "该设备已激活，无需重复激活",
    });
  }

  // ---------- 设备上限检查 ----------
  const machineCount = db.prepare(
    "SELECT COUNT(*) as count FROM activated_machines WHERE key_id = ?"
  ).get(licenseKey.id);

  if (machineCount.count >= licenseKey.max_machines) {
    return res.status(403).json({
      success: false,
      error: `卡密已达到最大设备绑定数量（${licenseKey.max_machines} 台）`,
    });
  }

  // ---------- 执行绑定 ----------
  // 插入设备绑定记录，初始标记为在线
  db.prepare(`
    INSERT INTO activated_machines (key_id, machine_id, machine_name, last_heartbeat, is_online)
    VALUES (?, ?, ?, datetime('now'), 1)
  `).run(licenseKey.id, machine_id, machine_name || null);

  // 增加 used_count，如果达到 max_machines 则自动将状态标记为 used
  db.prepare(`
    UPDATE license_keys
    SET used_count = used_count + 1,
        status = CASE WHEN used_count + 1 >= max_machines THEN 'used' ELSE status END,
        updated_at = datetime('now')
    WHERE id = ?
  `).run(licenseKey.id);

  // 计算到期时间（如果设置了 duration_days）
  let expiresAt = null;
  if (licenseKey.duration_days) {
    expiresAt = new Date();
    expiresAt.setDate(expiresAt.getDate() + licenseKey.duration_days);
    // 转为 SQLite 可存储的格式：YYYY-MM-DD HH:MM:SS
    expiresAt = expiresAt.toISOString().slice(0, 19).replace("T", " ");
  }

  res.json({
    success: true,
    activated: true,
    key: licenseKey.key_code,
    expires_at: expiresAt,
    duration_days: licenseKey.duration_days,
    max_machines: licenseKey.max_machines,
    message: "激活成功",
  });
});

export default router;
