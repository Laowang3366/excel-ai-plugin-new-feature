/**
 * 渲染进程日志工具
 *
 * 将日志转发到主进程持久化存储，同时保留控制台输出便于开发。
 */

import { ipcApi } from "../services/ipcApi";

type LogLevel = "info" | "warn" | "error";

async function sendLog(level: LogLevel, tag: string, message: string): Promise<void> {
  try {
    await ipcApi.app.log(level, tag, message);
  } catch {
    // 如果 IPC 不可用（如首次加载时），静默忽略
  }
}

export function logError(tag: string, message: string, error?: unknown): void {
  const detail = error instanceof Error ? error.message : String(error ?? "");
  const fullMsg = detail ? `${message}: ${detail}` : message;
  console.error(`[${tag}]`, message, error ?? "");
  sendLog("error", tag, fullMsg);
}

export function logWarn(tag: string, message: string): void {
  console.warn(`[${tag}]`, message);
  sendLog("warn", tag, message);
}

export function logInfo(tag: string, message: string): void {
  console.info(`[${tag}]`, message);
  sendLog("info", tag, message);
}
