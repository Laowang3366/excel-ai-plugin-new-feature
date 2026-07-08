/**
 * 心跳上报路由（公开接口）
 *
 * POST /api/heartbeat
 * 应用定期上报在线状态，服务端累计在线时长。
 * 每 5 分钟上报一次，2 分钟无上报视为离线。
 */

import { Router } from "express";
import { getDb } from "../db.js";
import config from "../config.js";

const router = Router();

router.post("/", (req, res) => {
  const { key, machine_id } = req.body || {};

  if (!key || !machine_id) {
    return res.status(400).json({ success: false, error: "参数不完整" });
  }

  const db = getDb();

  // 查找绑定记录
  const machine = db.prepare(`
    SELECT m.id, m.last_heartbeat, lk.status as key_status
    FROM activated_machines m
    JOIN license_keys lk ON m.key_id = lk.id
    WHERE lk.key_code = ? AND m.machine_id = ?
  `).get(key, machine_id);

  if (!machine) {
    return res.status(404).json({ success: false, error: "未找到激活记录" });
  }

  if (machine.key_status === "disabled") {
    return res.status(403).json({ success: false, error: "卡密已被禁用" });
  }

  // 计算从上一次心跳到现在的秒数差
  let deltaSeconds = 0;
  if (machine.last_heartbeat) {
    const lastTime = new Date(machine.last_heartbeat + "Z").getTime();
    const nowTime = Date.now();
    deltaSeconds = Math.floor((nowTime - lastTime) / 1000);

    // 限制最大增量（防止时钟跳跃导致异常大值）
    deltaSeconds = Math.min(deltaSeconds, config.heartbeatInterval * 2);
  }

  // 更新心跳时间和在线时长
  db.prepare(`
    UPDATE activated_machines
    SET last_heartbeat = datetime('now'),
        is_online = 1,
        total_online_seconds = total_online_seconds + ?
    WHERE id = ?
  `).run(Math.max(0, deltaSeconds), machine.id);

  res.json({
    success: true,
    delta_seconds: deltaSeconds,
    message: "心跳上报成功",
  });
});

export default router;
