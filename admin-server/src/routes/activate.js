/**
 * 激活验证路由（公开接口）
 *
 * POST /api/activate
 * 验证卡密有效性，绑定设备并返回激活结果。
 */

import { Router } from "express";
import { getDb } from "../db.js";
import config from "../config.js";

const router = Router();

router.post("/", (req, res) => {
  const { key, machine_id, machine_name } = req.body || {};

  // 参数校验
  if (!key) {
    return res.status(400).json({ success: false, error: "请输入卡密" });
  }
  if (!machine_id) {
    return res.status(400).json({ success: false, error: "缺少设备标识" });
  }

  // 标准化卡密（去除空格、转大写）
  const normalizedKey = key.trim().toUpperCase();

  const db = getDb();

  // 查找卡密
  const licenseKey = db.prepare("SELECT * FROM license_keys WHERE key_code = ?").get(normalizedKey);

  if (!licenseKey) {
    return res.status(404).json({ success: false, error: "卡密不存在" });
  }

  // 检查卡密状态
  if (licenseKey.status === "disabled") {
    return res.status(403).json({ success: false, error: "卡密已被禁用" });
  }

  if (licenseKey.status === "expired") {
    return res.status(403).json({ success: false, error: "卡密已过期" });
  }

  // 检查是否已绑定当前设备
  const existingMachine = db.prepare(
    "SELECT * FROM activated_machines WHERE key_id = ? AND machine_id = ?"
  ).get(licenseKey.id, machine_id);

  if (existingMachine) {
    // 已绑定此设备，更新心跳后直接返回成功
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

  // 检查已绑定的设备数量是否达到上限
  const machineCount = db.prepare(
    "SELECT COUNT(*) as count FROM activated_machines WHERE key_id = ?"
  ).get(licenseKey.id);

  if (machineCount.count >= licenseKey.max_machines) {
    return res.status(403).json({
      success: false,
      error: `卡密已达到最大设备绑定数量（${licenseKey.max_machines} 台）`,
    });
  }

  // 绑定设备
  db.prepare(`
    INSERT INTO activated_machines (key_id, machine_id, machine_name, last_heartbeat, is_online)
    VALUES (?, ?, ?, datetime('now'), 1)
  `).run(licenseKey.id, machine_id, machine_name || null);

  // 更新卡密使用计数和状态（达到设备上限时标记为 used）
  db.prepare(`
    UPDATE license_keys
    SET used_count = used_count + 1,
        status = CASE WHEN used_count + 1 >= max_machines THEN 'used' ELSE status END,
        updated_at = datetime('now')
    WHERE id = ?
  `).run(licenseKey.id);

  // 计算到期时间
  let expiresAt = null;
  if (licenseKey.duration_days) {
    expiresAt = new Date();
    expiresAt.setDate(expiresAt.getDate() + licenseKey.duration_days);
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
