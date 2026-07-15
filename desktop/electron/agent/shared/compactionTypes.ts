/** 压缩触发原因 */
export type CompactionReason =
  | "auto_token_limit" // 自动：token 接近上限
  | "auto_pre_turn" // 自动：新 Turn 开始前
  | "user_requested" // 用户手动触发
  | "interrupted_resume" // 中断后恢复时压缩
  | "model_changed" // 模型切换后压缩，避免旧模型上下文污染新模型
  | "context_window_changed"; // 上下文窗口变更后压缩，适配新的窗口预算

/** 压缩策略配置 */
export interface CompactionConfig {
  /** 是否启用自动压缩 */
  enabled: boolean;
  /** 触发自动压缩的 token 阈值（默认 100000） */
  autoCompactTokenThreshold: number;
  /** mid-turn 压缩触发比例，默认 0.9 */
  midTurnThresholdRatio?: number;
  /** 压缩后保留的最近用户消息最大 token 数（默认 20000） */
  retainedUserMessageMaxTokens: number;
  /** 压缩后最多保留最近多少条用户消息；未设置时只按 token 预算控制 */
  retainedRecentItemCount?: number;
  /** 摘要生成失败后的重试次数，默认 2 */
  summaryRetryCount?: number;
  /** 摘要生成重试的首次等待时间，默认由 AgentLoop 控制 */
  summaryRetryBaseDelayMs?: number;
  /** 摘要生成重试的单次等待上限，默认由 AgentLoop 控制 */
  summaryRetryMaxDelayMs?: number;
  /** 摘要生成重试的退避倍率，默认由 AgentLoop 控制 */
  summaryRetryBackoffFactor?: number;
  /** rollout 超过该字节数后可生成 gzip 归档快照 */
  archiveRolloutAfterBytes?: number;
  /** 压缩提示词（可自定义） */
  compactPrompt?: string;
  /** 压缩摘要生成方式：local 使用当前模型，remote 调用远程压缩服务 */
  compactionProvider?: "local" | "remote";
  /** 远程压缩服务地址，要求兼容 { instruction, input, model? } 请求 */
  remoteCompactUrl?: string;
  /** 远程压缩服务鉴权 token，可选 */
  remoteCompactApiKey?: string;
  /** 远程压缩服务模型名，可选 */
  remoteCompactModel?: string;
  /** 模型上下文窗口大小（参考 Codex model_context_window），用于感知实际窗口上限 */
  contextWindowSize?: number;
}

export const DEFAULT_COMPACTION_CONFIG: CompactionConfig = {
  enabled: true,
  autoCompactTokenThreshold: 100_000,
  midTurnThresholdRatio: 0.9,
  retainedUserMessageMaxTokens: 20_000,
  summaryRetryCount: 2,
};

/** 上下文压缩启动参数，用于让客户端观察压缩流程 */
export interface ThreadCompactStartParams {
  reason: CompactionReason;
  itemCount: number;
  tokensBefore: number;
  tokenThreshold: number;
  contextWindowSize?: number;
  retryCount: number;
  timestamp: number;
}
