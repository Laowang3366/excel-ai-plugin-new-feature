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

/**
 * 将日志消息通过 IPC 发送到主进程持久化
 *
 * 外层 catch 静默处理 IPC 不可用的场景（如预加载脚本未注入 electronAPI）。
 * 注意：日志级别统一由这里控制，与 console 方法独立。
 */
async function sendLog(tag: string, message: string): Promise<void> {
  try {
    await ipcApi.app.log("warn", tag, message);
  } catch {
    // IPC 不可用（如首次加载时的初始化阶段），不阻塞业务逻辑
  }
}

/**
 * 记录警告日志
 */
export function logWarn(tag: string, message: string): void {
  console.warn(`[${tag}]`, message);
  sendLog(tag, message);
}
