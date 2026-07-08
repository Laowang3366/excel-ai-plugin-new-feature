/**
 * 用户自助设备管理接口（公开接口，无需 JWT）
 *
 * 客户端（插件端）可通过这些接口查看同卡密下的其他设备，
 * 以及主动解绑设备以释放绑定名额。
 *
 * 安全约束：
 * - 请求者必须提供 key + machine_id 证明自己是该卡密的合法持有者
 * - 每次操作前通过 findBinding + assertBoundMachine 验证身份
 * - 解绑操作不能越权：只能操作同一卡密下的设备
 *
 * 路由：
 * - POST /api/devices/list   — 查看同一卡密下的所有设备
 * - POST /api/devices/unbind — 解绑指定设备
 */

import { Router } from "express";
import { getDb } from "../db.js";
import config from "../config.js";
import { assertBoundMachine, normalizeDevicePayload } from "../selfServiceDevices.js";
import { withBeijingDateTimes } from "../time.js";

const router = Router();

/**
 * POST /api/devices/list
 *
 * 返回当前卡密下的所有已激活设备列表。
 * 请求参数：{ key, machineId }
 * 其中 machineId 用于验证请求者身份。
 */
router.post("/list", (req, res) => {
  const { key, machineId } = normalizeDevicePayload(req.body);
  if (!key || !machineId) {
    return res.status(400).json({ success: false, error: "参数不完整" });
  }

  const db = getDb();
  let binding;
  try {
    // 验证请求者身份：必须是该卡密的已绑定设备
    binding = assertBoundMachine(findBinding(db, key, machineId));
  } catch (err) {
    return res.status(err.statusCode || 500).json({ success: false, error: err.message });
  }

  // 查询卡密信息（设备上限、已用数量）
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
        is_current: machine.machine_id === machineId, // 标记当前设备
        is_online: isOnline(machine.last_heartbeat),
        online_duration_formatted: formatDuration(machine.total_online_seconds),
        last_heartbeat_ago: secondsSinceHeartbeat != null ? formatTimeAgo(Math.max(0, secondsSinceHeartbeat)) : "从未上报",
      };
    }),
  });
});

/**
 * POST /api/devices/unbind
 *
 * 解绑指定设备（释放一个绑定名额）。
 * 请求参数：{ key, machineId, targetMachineId }
 *
 * 解绑后如果卡密状态为 used 且不再有任何绑定设备，自动恢复为 active。
 * 如果是解绑自己的设备（targetMachineId === machineId），标记 current_device_unbound。
 */
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

  // 查找目标设备（必须在同一卡密下）
  const target = db.prepare(`
    SELECT id FROM activated_machines
    WHERE key_id = ? AND machine_id = ?
  `).get(binding.keyId, targetMachineId);

  if (!target) {
    return res.status(404).json({ success: false, error: "目标设备不存在" });
  }

  // 事务：删除设备 + 更新卡密计数和状态
  const transaction = db.transaction(() => {
    const result = db.prepare("DELETE FROM activated_machines WHERE id = ?").run(target.id);
    // 重新计算 used_count（实际绑定的设备数），并自动恢复状态
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

/**
 * 查询指定卡密和设备 ID 的绑定关系
 *
 * @returns {object|null} { keyId, keyCode, status, machineId } 或 null
 */
function findBinding(db, key, machineId) {
  return db.prepare(`
    SELECT lk.id, lk.key_code, lk.status, m.machine_id
    FROM license_keys lk
    JOIN activated_machines m ON m.key_id = lk.id
    WHERE lk.key_code = ? AND m.machine_id = ?
  `).get(key, machineId);
}

/** 判断设备是否在线（同 admin.js 中的实现） */
function isOnline(lastHeartbeat) {
  if (!lastHeartbeat) return false;
  const now = Math.floor(Date.now() / 1000);
  const heartbeat = Math.floor(new Date(`${lastHeartbeat}Z`).getTime() / 1000);
  return (now - heartbeat) < config.heartbeatTimeout;
}

/** 格式化秒数为可读时长 */
function formatDuration(totalSeconds) {
  if (!totalSeconds || totalSeconds <= 0) return "0 分钟";
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  if (hours > 0) return `${hours} 小时 ${minutes} 分钟`;
  return `${minutes} 分钟`;
}

/** 格式化秒数差为可读文本 */
function formatTimeAgo(seconds) {
  if (seconds < 60) return `${seconds} 秒前`;
  if (seconds < 3600) return `${Math.floor(seconds / 60)} 分钟前`;
  if (seconds < 86400) return `${Math.floor(seconds / 3600)} 小时前`;
  return `${Math.floor(seconds / 86400)} 天前`;
}

export default router;
