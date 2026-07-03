import type { AIClientConfig } from "../../providers/aiClient";

/**
 * 模型压缩兼容性判断。
 *
 * 关联模块：
 * - agentLoop.ts: 切换模型配置时根据 compHash 决定是否需要预压缩。
 * - threadLifecycle.ts: 新建线程时记录初始 compHash，便于恢复和排查。
 */
export function resolveModelCompHash(config: AIClientConfig): string {
  const explicitHash = config.compHash?.trim();
  if (explicitHash) return explicitHash;

  const apiFormat = config.apiFormat?.trim() || config.provider;
  const model = config.model?.trim() || "unknown";
  return `${apiFormat}:${config.provider}:${model}`;
}

export function isModelCompHashCompatible(
  previous: AIClientConfig,
  next: AIClientConfig
): boolean {
  return resolveModelCompHash(previous) === resolveModelCompHash(next);
}
