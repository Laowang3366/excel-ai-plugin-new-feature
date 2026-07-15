/**
 * 会话级压缩配置构建。
 *
 * 关联模块：
 * - agentLoop.ts: 每个线程启动前根据当前模型上下文窗口计算阈值。
 * - runtime/compactionRuntime.ts: 提供全局压缩配置来源。
 */

import type { CompactionConfig } from "../../shared/types";

/**
 * 根据全局配置和会话的 contextWindowSize 计算会话级阈值，
 * 确保不同会话间的压缩隔离。
 */
export function buildSessionCompactionConfig(
  globalConfig: CompactionConfig,
  sessionContextWindowSize: number,
): CompactionConfig {
  const globalWindow = globalConfig.contextWindowSize || 128_000;
  const thresholdPercent = globalConfig.autoCompactTokenThreshold / globalWindow;
  return {
    ...globalConfig,
    contextWindowSize: sessionContextWindowSize,
    autoCompactTokenThreshold: Math.floor(sessionContextWindowSize * thresholdPercent),
  };
}
