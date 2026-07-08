export function normalizeDevicePayload(body = {}) {
  return {
    key: String(body.key || "").trim().toUpperCase(),
    machineId: String(body.machine_id || "").trim(),
    targetMachineId: String(body.target_machine_id || "").trim(),
  };
}

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
