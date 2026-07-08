/**
 * 用户自助设备管理接口
 *
 * 已绑定设备可以查看同一卡密下的设备，并解绑设备释放名额。
 */

import { Router } from "express";
import { getDb } from "../db.js";
import config from "../config.js";
import { assertBoundMachine, normalizeDevicePayload } from "../selfServiceDevices.js";
import { withBeijingDateTimes } from "../time.js";

const router = Router();

router.post("/list", (req, res) => {
  const { key, machineId } = normalizeDevicePayload(req.body);
  if (!key || !machineId) {
    return res.status(400).json({ success: false, error: "参数不完整" });
  }

  const db = getDb();
  let binding;
  try {
    binding = assertBoundMachine(findBinding(db, key, machineId));
  } catch (err) {
    return res.status(err.statusCode || 500).json({ success: false, error: err.message });
  }

  const keyInfo = db.prepare("SELECT max_machines, used_count FROM license_keys WHERE id = ?").get(binding.keyId);
  const machines = db.prepare(`
    SELECT id, machine_id, machine_name, activated_at, last_heartbeat, is_online, total_online_seconds
    FROM activated_machines
    WHERE key_id = ?
    ORDER BY activated_at DESC
  `).all(binding.keyId);

  res.json({
    success: true,
    key,
    current_machine_id: machineId,
    max_machines: keyInfo?.max_machines || 0,
    used_count: machines.length,
    machines: machines.map((machine) => {
      const secondsSinceHeartbeat = machine.last_heartbeat
        ? Math.floor(Date.now() / 1000) - Math.floor(new Date(`${machine.last_heartbeat}Z`).getTime() / 1000)
        : null;
      return {
        ...withBeijingDateTimes(machine),
        is_current: machine.machine_id === machineId,
        is_online: isOnline(machine.last_heartbeat),
        online_duration_formatted: formatDuration(machine.total_online_seconds),
        last_heartbeat_ago: secondsSinceHeartbeat != null ? formatTimeAgo(Math.max(0, secondsSinceHeartbeat)) : "从未上报",
      };
    }),
  });
});

router.post("/unbind", (req, res) => {
  const { key, machineId, targetMachineId } = normalizeDevicePayload(req.body);
  if (!key || !machineId || !targetMachineId) {
    return res.status(400).json({ success: false, error: "参数不完整" });
  }

  const db = getDb();
  let binding;
  try {
    binding = assertBoundMachine(findBinding(db, key, machineId));
  } catch (err) {
    return res.status(err.statusCode || 500).json({ success: false, error: err.message });
  }

  const target = db.prepare(`
    SELECT id FROM activated_machines
    WHERE key_id = ? AND machine_id = ?
  `).get(binding.keyId, targetMachineId);

  if (!target) {
    return res.status(404).json({ success: false, error: "目标设备不存在" });
  }

  const transaction = db.transaction(() => {
    const result = db.prepare("DELETE FROM activated_machines WHERE id = ?").run(target.id);
    db.prepare(`
      UPDATE license_keys
      SET used_count = (SELECT COUNT(*) FROM activated_machines WHERE key_id = ?),
          status = CASE
            WHEN status = 'used' AND (SELECT COUNT(*) FROM activated_machines WHERE key_id = ?) = 0 THEN 'active'
            ELSE status
          END,
          updated_at = datetime('now')
      WHERE id = ?
    `).run(binding.keyId, binding.keyId, binding.keyId);
    return result;
  });

  const result = transaction();
  res.json({
    success: true,
    unbound: result.changes,
    current_device_unbound: targetMachineId === machineId,
  });
});

function findBinding(db, key, machineId) {
  return db.prepare(`
    SELECT lk.id, lk.key_code, lk.status, m.machine_id
    FROM license_keys lk
    JOIN activated_machines m ON m.key_id = lk.id
    WHERE lk.key_code = ? AND m.machine_id = ?
  `).get(key, machineId);
}

function isOnline(lastHeartbeat) {
  if (!lastHeartbeat) return false;
  const now = Math.floor(Date.now() / 1000);
  const heartbeat = Math.floor(new Date(`${lastHeartbeat}Z`).getTime() / 1000);
  return (now - heartbeat) < config.heartbeatTimeout;
}

function formatDuration(totalSeconds) {
  if (!totalSeconds || totalSeconds <= 0) return "0 分钟";
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  if (hours > 0) return `${hours} 小时 ${minutes} 分钟`;
  return `${minutes} 分钟`;
}

function formatTimeAgo(seconds) {
  if (seconds < 60) return `${seconds} 秒前`;
  if (seconds < 3600) return `${Math.floor(seconds / 60)} 分钟前`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)} 小时前`;
  return `${Math.floor(seconds / 86400)} 天前`;
}

export default router;
