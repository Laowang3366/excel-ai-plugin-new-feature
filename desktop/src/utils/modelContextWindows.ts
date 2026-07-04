/**
 * 上下文窗口工具（渲染器侧）
 *
 * 上下文窗口大小由用户在供应商配置中自定义填写（contextWindowSize 字段），
 * 不再使用硬编码映射表。此文件仅保留默认值和显示格式化工具。
 *
 * 与主进程侧 electron/agent/modelContextWindows.ts 保持同步。
 */

/** 默认上下文窗口大小（未设置 contextWindowSize 时的回退值） */
export const DEFAULT_CONTEXT_WINDOW = 128_000;

/**
 * 将 tokens 转换为 k/M 单位显示字符串
 *
 * @param tokens token 数量
 * @returns 如 "256k"、"1M"
 */
export function formatTokensAsK(tokens: number): string {
  if (tokens >= 1_000_000) {
    const m = tokens / 1_000_000;
    return m === Math.floor(m) ? `${m}M` : `${m.toFixed(1)}M`;
  }
  const k = Math.floor(tokens / 1000);
  return `${k}k`;
}

export function formatEstimatedUsedTokens(tokens: number): string {
  if (tokens > 0 && tokens < 1000) return "<1k";
  return `${Math.floor(tokens / 1000)}k`;
}
