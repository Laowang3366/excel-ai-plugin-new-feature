/**
 * 设备自助服务工具函数
 *
 * 为 routes/devices.js 提供请求参数清洗和身份验证的辅助逻辑。
 */

/**
 * 标准化设备请求体
 *
 * 统一处理字段命名（将 snake_case 转为 camelCase）、去除首尾空格、
 * 卡密转大写，消除客户端编码差异。
 *
 * @param {object} [body={}] - 原始请求体
 * @returns {{ key: string, machineId: string, targetMachineId: string }}
 */
export function normalizeDevicePayload(body = {}) {
  return {
    key: String(body.key || "").trim().toUpperCase(),
    machineId: String(body.machine_id || "").trim(),
    targetMachineId: String(body.target_machine_id || "").trim(),
  };
}

/**
 * 断言设备已绑定且卡密状态正常
 *
 * 检查绑定记录是否存在以及卡密状态。
 * 若验证不通过，抛出带 statusCode 的 Error，便于调用方直接用于 HTTP 响应。
 *
 * @param {object|null} row - findBinding 的查询结果
 * @returns {{ keyId: number, keyCode: string, machineId: string }} 验证通过后的绑定信息
 * @throws {Error} 状态码 403（未绑定/禁用/过期）
 */
export function assertBoundMachine(row) {
  if (!row) {
    throw Object.assign(new Error("当前设备未绑定该卡密"), { statusCode: 403 });
  }
  if (row.status === "disabled") {
    throw Object.assign(new Error("卡密已被禁用"), { statusCode: 403 });
  }
  if (row.status === "expired") {
    throw Object.assign(new Error("卡密已过期"), { statusCode: 403 });
  }
  return {
    keyId: row.id,
    keyCode: row.key_code,
    machineId: row.machine_id,
  };
}
