/**
 * 结构化日志模块 — 基于 Node.js fs 模块实现
 *
 * 替代 electron-log（网络安装受限），提供：
 * - 分级日志（debug/info/warn/error）
* - 文件持久化（自动按日期分割）
 * - 控制台彩色输出（开发环境）
 * - 进程级未捕获异常记录
 *
 * 日志格式：[ISO时间] [级别] [来源] 消息 {结构化数据}
 */

import * as fs from "fs";
import * as path from "path";
import { app } from "electron";

type LogLevel = "debug" | "info" | "warn" | "error";

const LOG_LEVELS: Record<LogLevel, number> = {
  debug: 0,
  info: 1,
  warn: 2,
  error: 3,
};

const LOG_COLORS: Record<LogLevel, string> = {
  debug: "\x1b[90m",   // gray
  info: "\x1b[36m",    // cyan
  warn: "\x1b[33m",    // yellow
  error: "\x1b[31m",   // red
};
const RESET = "\x1b[0m";

const MIN_LEVEL: LogLevel = process.env.NODE_ENV === "development" ? "debug" : "info";

let logDir: string | null = null;

export function configureLogDirectory(directory: string): void {
  logDir = directory;
}

function getLogDir(): string {
  if (logDir) {
    try {
      if (!fs.existsSync(logDir)) {
        fs.mkdirSync(logDir, { recursive: true });
      }
      return logDir;
    } catch {
      logDir = null;
    }
  }
  try {
    logDir = path.join(app.getPath("userData"), "logs");
    if (!fs.existsSync(logDir)) {
      fs.mkdirSync(logDir, { recursive: true });
    }
  } catch {
    logDir = path.join(app.getPath("temp"), "excel-ai-logs");
    if (!fs.existsSync(logDir)) {
      fs.mkdirSync(logDir, { recursive: true });
    }
  }
  return logDir;
}

function getLogFilePath(): string {
  const now = new Date();
  const date = now.toISOString().split("T")[0]; // YYYY-MM-DD
  return path.join(getLogDir(), `app-${date}.log`);
}

function formatMessage(level: LogLevel, source: string, message: string, data?: unknown): string {
  const timestamp = new Date().toISOString();
  const base = `[${timestamp}] [${level.toUpperCase().padEnd(5)}] [${source}] ${message}`;
  if (data !== undefined) {
    const dataStr = typeof data === "string" ? data : JSON.stringify(data, null, 0);
    return `${base} ${dataStr}`;
  }
  return base;
}

function writeToFile(formatted: string): void {
  try {
    fs.appendFileSync(getLogFilePath(), formatted + "\n", { encoding: "utf-8" });
  } catch {
    // 写入失败时静默忽略，避免日志系统本身抛出异常
  }
}

function log(level: LogLevel, source: string, message: string, data?: unknown): void {
  if (LOG_LEVELS[level] < LOG_LEVELS[MIN_LEVEL]) return;

  const formatted = formatMessage(level, source, message, data);

  // 控制台输出（带颜色）
  const color = LOG_COLORS[level];
  console.log(`${color}${formatted}${RESET}`);

  // 文件持久化
  writeToFile(formatted);
}

/**
 * 创建一个绑定来源名称的 logger 实例
 *
 * @example
 * const logger = createLogger("AgentLoop");
 * logger.info("Turn started", { turnId });
 * logger.error("Tool execution failed", { toolName, error });
 */
export function createLogger(source: string) {
  return {
    debug: (message: string, data?: unknown) => log("debug", source, message, data),
    info: (message: string, data?: unknown) => log("info", source, message, data),
    warn: (message: string, data?: unknown) => log("warn", source, message, data),
    error: (message: string, data?: unknown) => log("error", source, message, data),
  };
}

/**
 * 全局日志实例（无来源标签，用于通用场景）
 */
export const logger = {
  debug: (message: string, data?: unknown) => log("debug", "app", message, data),
  info: (message: string, data?: unknown) => log("info", "app", message, data),
  warn: (message: string, data?: unknown) => log("warn", "app", message, data),
  error: (message: string, data?: unknown) => log("error", "app", message, data),
};

/**
 * 注册进程级错误捕获，确保未捕获异常写入日志文件
 */
export function setupGlobalErrorHandlers(): void {
  process.on("uncaughtException", (error) => {
    logger.error("Uncaught Exception:", {
      message: error.message,
      stack: error.stack,
    });
  });

  process.on("unhandledRejection", (reason) => {
    const info = reason instanceof Error
      ? { message: reason.message, stack: reason.stack }
      : { reason: String(reason) };
    logger.error("Unhandled Rejection:", info);
  });
}
