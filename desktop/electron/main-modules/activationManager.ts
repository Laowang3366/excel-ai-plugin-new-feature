/**
 * 激活管理 - 主进程模块
 *
 * 负责：
 * - 激活状态持久化（electron-store）
 * - 与服务端通信（激活验证、心跳上报）
 * - 设备标识生成
 * - 心跳定时器管理
 */

import { getSettingsStore } from "../main-modules/settingsManager";
import { createLogger } from "../shared/logger";
import * as os from "os";
import * as crypto from "crypto";

const logger = createLogger("activation");

// ============================================================
// 常量
// ============================================================

/** 激活信息在 electron-store 中的键名 */
const STORAGE_KEY = "activation";

/** 默认服务器地址（用户可在设置中修改） */
const DEFAULT_SERVER_URL = "http://localhost:3456";

/** 心跳间隔（毫秒） */
const HEARTBEAT_INTERVAL = 5 * 60 * 1000; // 5 分钟

/** 离线容忍时间（毫秒） */
const OFFLINE_TOLERANCE = 72 * 60 * 60 * 1000; // 72 小时

// ============================================================
// 类型
// ============================================================

export interface ActivationInfo {
  /** 激活的卡密 */
  key: string;
  /** 设备唯一标识 */
  machineId: string;
  /** 激活时间 */
  activatedAt: string;
  /** 最后验证时间 */
  lastVerifiedAt: string;
  /** 卡密过期时间 */
  expiresAt: string | null;
  /** 服务器地址 */
  serverUrl: string;
}

interface ActivationState {
  activated: boolean;
  info: ActivationInfo | null;
}

// ============================================================
// 状态
// ============================================================

let heartbeatTimer: ReturnType<typeof setInterval> | null = null;

// ============================================================
// 设备标识
// ============================================================

/**
 * 生成设备唯一标识
 * 使用 hostname + 用户名 的 SHA256 哈希
 */
export function getMachineId(): string {
  const raw = `${os.hostname()}-${os.userInfo().username}`;
  return crypto.createHash("sha256").update(raw).digest("hex").slice(0, 16);
}

/**
 * 获取设备名称
 */
export function getMachineName(): string {
  return os.hostname();
}

// ============================================================
// 持久化
// ============================================================

function getActivationStore() {
  return getSettingsStore();
}

/**
 * 读取本地激活状态
 */
export function getLocalActivationState(): ActivationState {
  const store = getActivationStore();
  const saved = store.get(STORAGE_KEY) as ActivationState | undefined;
  if (saved && saved.activated && saved.info) {
    return saved;
  }
  return { activated: false, info: null };
}

/**
 * 保存激活状态到本地
 */
function saveActivationState(state: ActivationState): void {
  const store = getActivationStore();
  store.set(STORAGE_KEY, state);
}

/**
 * 清除本地激活状态
 */
export function clearActivationState(): void {
  const store = getActivationStore();
  store.set(STORAGE_KEY, { activated: false, info: null });
  stopHeartbeat();
  logger.info("激活状态已清除");
}

// ============================================================
// HTTP 请求（使用 Node.js 原生 fetch）
// ============================================================

async function postJson(url: string, body: unknown): Promise<Response> {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  return response;
}

async function readJsonResponse(response: Response): Promise<Record<string, any>> {
  const data = (await response.json()) as Record<string, any>;
  if (!response.ok) {
    throw new Error((data.error as string) || `服务器错误 (${response.status})`);
  }
  return data;
}

// ============================================================
// 核心 API
// ============================================================

/**
 * 获取已保存的服务器地址，或返回默认值
 */
export function getServerUrl(): string {
  const state = getLocalActivationState();
  return state.info?.serverUrl || DEFAULT_SERVER_URL;
}

/**
 * 保存服务器地址
 */
export function setServerUrl(url: string): void {
  const state = getLocalActivationState();
  if (state.info) {
    state.info.serverUrl = url;
    saveActivationState(state);
  }
}

/**
 * 激活验证
 *
 * 向服务端发送激活请求，验证卡密并绑定设备。
 *
 * @param key 卡密
 * @param serverUrl 服务器地址
 * @returns 激活结果
 */
export async function activate(key: string, serverUrl: string): Promise<{ success: boolean; error?: string; data?: any }> {
  const machineId = getMachineId();
  const machineName = getMachineName();

  // 标准化服务器地址
  const baseUrl = serverUrl.replace(/\/+$/, "");
  const url = `${baseUrl}/api/activate`;

  try {
    const response = await postJson(url, {
      key: key.toUpperCase(),
      machine_id: machineId,
      machine_name: machineName,
    });

    const data = (await response.json()) as Record<string, unknown>;

    if (!response.ok) {
      return { success: false, error: (data.error as string) || `服务器错误 (${response.status})` };
    }

    if (data.success) {
      // 保存激活状态
      const state: ActivationState = {
        activated: true,
        info: {
          key: data.key as string,
          machineId,
          activatedAt: new Date().toISOString(),
          lastVerifiedAt: new Date().toISOString(),
          expiresAt: (data.expires_at as string) || null,
          serverUrl: baseUrl,
        },
      };
      saveActivationState(state);

      // 启动心跳
      startHeartbeat();

      logger.info(`激活成功: ${data.key}`);
    }

    return { success: data.success as boolean, data };
  } catch (err: any) {
    const message = err.message || "网络错误，无法连接到服务器";
    logger.error("激活失败:", message);
    return { success: false, error: message };
  }
}

export async function listBoundDevices(): Promise<{ success: boolean; error?: string; data?: any }> {
  const state = getLocalActivationState();
  if (!state.activated || !state.info) {
    return { success: false, error: "应用未激活" };
  }

  try {
    const response = await postJson(`${state.info.serverUrl}/api/devices/list`, {
      key: state.info.key,
      machine_id: state.info.machineId,
    });
    return { success: true, data: await readJsonResponse(response) };
  } catch (err: any) {
    return { success: false, error: err.message || "设备列表加载失败" };
  }
}

export async function unbindDevice(targetMachineId: string): Promise<{ success: boolean; error?: string; currentDeviceUnbound?: boolean }> {
  const state = getLocalActivationState();
  if (!state.activated || !state.info) {
    return { success: false, error: "应用未激活" };
  }

  try {
    const response = await postJson(`${state.info.serverUrl}/api/devices/unbind`, {
      key: state.info.key,
      machine_id: state.info.machineId,
      target_machine_id: targetMachineId,
    });
    const data = await readJsonResponse(response);
    if (data.current_device_unbound) {
      clearActivationState();
    }
    return { success: true, currentDeviceUnbound: Boolean(data.current_device_unbound) };
  } catch (err: any) {
    return { success: false, error: err.message || "设备解绑失败" };
  }
}

/**
 * 发送心跳
 */
async function sendHeartbeat(): Promise<boolean> {
  const state = getLocalActivationState();
  if (!state.activated || !state.info) return false;

  const baseUrl = state.info.serverUrl;
  const url = `${baseUrl}/api/heartbeat`;

  try {
    const response = await postJson(url, {
      key: state.info.key,
      machine_id: state.info.machineId,
    });

    const responseData = (await response.json()) as Record<string, unknown>;

    if (!response.ok) {
      if (response.status === 403) {
        // 卡密被禁用
        logger.warn("卡密已被禁用，清除激活状态");
        clearActivationState();
        return false;
      }
      logger.warn("心跳失败:", (responseData.error as string) || response.statusText);
      return false;
    }

    // 更新最后验证时间
    state.info.lastVerifiedAt = new Date().toISOString();
    saveActivationState(state);
    return true;
  } catch (err: any) {
    logger.warn("心跳上报失败:", err.message);
    return false;
  }
}

/**
 * 检查激活状态是否有效
 * - 已激活且未超过离线容忍时间 → 有效
 * - 已激活但超过离线容忍时间 → 需要重新验证
 * - 未激活 → 无效
 */
export function checkActivationValid(): boolean {
  const state = getLocalActivationState();
  if (!state.activated || !state.info) return false;

  // 检查是否在离线容忍时间内
  const lastVerified = new Date(state.info.lastVerifiedAt).getTime();
  const now = Date.now();
  if (now - lastVerified > OFFLINE_TOLERANCE) {
    logger.warn("超过离线容忍时间，需要重新验证");
    return false;
  }

  // 检查卡密是否过期
  if (state.info.expiresAt) {
    const expires = new Date(state.info.expiresAt).getTime();
    if (now > expires) {
      logger.warn("卡密已过期");
      return false;
    }
  }

  return true;
}

// ============================================================
// 心跳定时器
// ============================================================

/**
 * 启动心跳定时器
 */
export function startHeartbeat(): void {
  if (heartbeatTimer) return;

  logger.info("启动心跳定时器");
  // 立即发送一次
  sendHeartbeat();
  heartbeatTimer = setInterval(sendHeartbeat, HEARTBEAT_INTERVAL);
}

/**
 * 停止心跳定时器
 */
export function stopHeartbeat(): void {
  if (heartbeatTimer) {
    // 关闭前发送最后一次心跳
    sendHeartbeat().finally(() => {
      clearInterval(heartbeatTimer!);
      heartbeatTimer = null;
      logger.info("心跳定时器已停止");
    });
  }
}

/**
 * 初始化激活系统
 * 在应用启动时调用，检查状态并启动心跳
 */
export function initActivation(): void {
  const state = getLocalActivationState();
  if (state.activated && state.info) {
    logger.info("检测到已激活，启动心跳");
    startHeartbeat();
  } else {
    logger.info("未激活，等待激活");
  }
}
