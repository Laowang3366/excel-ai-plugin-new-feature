/**
 * 心跳上报路由（公开接口，无需 JWT）
 *
 * POST /api/heartbeat
 *
 * 客户端（插件端）定期调用此接口告知服务端自身在线。
 * 服务端累计在线时长并记录最后心跳时间，用于判断设备在线/离线状态。
 *
 * 在线判定逻辑：
 * - 设备当前时间 - last_heartbeat < heartbeatTimeout（默认 120s）→ 在线
 * - 否则视为离线（离线判断由 /api/admin/machines 接口中的 SQL 完成）
 *
 * 心跳上报间隔约定：
 * - 客户端应每 heartbeatInterval 秒（默认 300s = 5min）上报一次
 * - 服务端不会检查上报间隔是否严格符合约定（容忍网络波动）
 *
 * 在线时长（total_online_seconds）计算：
 * - 每次上报时计算当前时间与上次心跳的时间差（deltaSeconds）
 * - 累加到 total_online_seconds
 * - deltaSeconds 上限为 heartbeatInterval * 2，防止时钟跳跃导致异常大值
 */

import { Router } from "express";
import { getDb } from "../db.js";
import config from "../config.js";

const router = Router();

router.post("/", (req, res) => {
  const { key, machine_id } = req.body || {};

  // ---------- 参数校验 ----------
  if (!key || !machine_id) {
    return res.status(400).json({ success: false, error: "参数不完整" });
  }

  const db = getDb();

  // ---------- 查找绑定记录 ----------
  // 必须同时匹配 key_code 和 machine_id，确保设备确实持有该卡密的合法激活记录
  const machine = db.prepare(`
    SELECT m.id, m.last_heartbeat, lk.status as key_status
    FROM activated_machines m
    JOIN license_keys lk ON m.key_id = lk.id
    WHERE lk.key_code = ? AND m.machine_id = ?
  `).get(key, machine_id);

  if (!machine) {
    return res.status(404).json({ success: false, error: "未找到激活记录" });
  }

  // 如果卡密已被管理员禁用，拒绝心跳上报
  if (machine.key_status === "disabled") {
    return res.status(403).json({ success: false, error: "卡密已被禁用" });
  }

  // ---------- 计算心跳间隔（delta）----------
  // 用于累计在线时长。若为首次上报（last_heartbeat 为空），delta 为 0。
  let deltaSeconds = 0;
  if (machine.last_heartbeat) {
    const lastTime = new Date(machine.last_heartbeat + "Z").getTime();
    const nowTime = Date.now();
    deltaSeconds = Math.floor((nowTime - lastTime) / 1000);

    // 限幅：防止客户端因休眠/时钟调整产生异常大的 delta
    // 上限设为 heartbeatInterval * 2（例如 300s * 2 = 600s）
    // 若超时未上报，缺失的时间段不计入在线时长
    deltaSeconds = Math.min(deltaSeconds, config.heartbeatInterval * 2);
  }

  // ---------- 更新记录 ----------
  // 更新最后心跳时间、在线标志、累计在线秒数
  db.prepare(`
    UPDATE activated_machines
    SET last_heartbeat = datetime('now'),
        is_online = 1,
        total_online_seconds = total_online_seconds + ?
    WHERE id = ?
  `).run(Math.max(0, deltaSeconds), machine.id);

  res.json({
    success: true,
    delta_seconds: deltaSeconds, // 返回本次累计的秒数，便于客户端调试
    message: "心跳上报成功",
  });
});

export default router;
