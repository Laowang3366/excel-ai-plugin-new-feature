/**
 * 渲染进程日志工具
 *
 * 日志输出策略（双通道）：
 * 1. console.* — 开发时在 DevTools 中即时查看
 * 2. IPC 转发 — 将日志发送到主进程，由主进程持久化写入日志文件
 *
 * 容错：IPC 转发失败（如 electronAPI 尚未就绪）时静默忽略，不干扰业务逻辑。
 */

import { ipcApi } from "../services/ipcApi";

type LogLevel = "info" | "warn" | "error";

/**
 * 将日志消息通过 IPC 发送到主进程持久化
 *
 * 外层 catch 静默处理 IPC 不可用的场景（如预加载脚本未注入 electronAPI）。
 * 注意：日志级别统一由这里控制，与 console 方法独立。
 */
async function sendLog(level: LogLevel, tag: string, message: string): Promise<void> {
  try {
    await ipcApi.app.log(level, tag, message);
  } catch {
    // IPC 不可用（如首次加载时的初始化阶段），不阻塞业务逻辑
  }
}

/**
 * 记录错误日志
 *
 * error 对象序列化策略：
 * - Error 实例 → 取 .message
 * - 非 Error（如普通对象/字符串）→ 转为字符串
 * - undefined/空 → 仅记录 message 本身
 *
 * 同时输出到 console.error（浏览器 DevTools）和 IPC（主进程日志文件）。
 */
export function logError(tag: string, message: string, error?: unknown): void {
  const detail = error instanceof Error ? error.message : String(error ?? "");
  const fullMsg = detail ? `${message}: ${detail}` : message;
  console.error(`[${tag}]`, message, error ?? "");
  sendLog("error", tag, fullMsg);
}

/**
 * 记录警告日志
 */
export function logWarn(tag: string, message: string): void {
  console.warn(`[${tag}]`, message);
  sendLog("warn", tag, message);
}

/**
 * 记录信息日志
 */
export function logInfo(tag: string, message: string): void {
  console.info(`[${tag}]`, message);
  sendLog("info", tag, message);
}
