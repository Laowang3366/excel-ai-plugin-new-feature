/**
 * 压缩提示词模板。
 *
 * 关联模块：
 * - core/agentLoop/summaryGenerator.ts: 生成上下文摘要时读取默认模板。
 * - memory/compaction.ts: 负责历史转文本与压缩结果构建，不再持有提示词正文。
 */

import defaultCompactPrompt from "./templates/compaction.zh-CN.md?raw";

export const DEFAULT_COMPACT_PROMPT = defaultCompactPrompt.trim();

export function getCompactionPromptTemplate(override?: string): string {
  const trimmed = override?.trim();
  return trimmed || DEFAULT_COMPACT_PROMPT;
}
