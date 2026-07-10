/**
 * 压缩摘要生成。
 *
 * 关联模块：
 * - agentLoop.ts: 在 pre-turn 和 mid-turn 压缩时调用本模块生成摘要。
 * - prompts/compactionPrompt.ts: 提供默认摘要提示词模板。
 */

import type { CompactionConfig } from "../../shared/types";
import { getCompactionPromptTemplate } from "../../prompts/compactionPrompt";

/**
 * 生成压缩摘要
 *
 * 使用 AI 客户端生成对话历史的摘要文本。
 * 失败时向上抛出，由 AgentLoop 统一重试并记录压缩失败。
 *
 * @param aiClient - AI 客户端
 * @param prompt - 要摘要的对话内容
 * @param compactionConfig - 压缩配置（含自定义 prompt）
 * @returns 摘要文本
 */
export async function generateSummary(
  aiClient: { chat: (params: any) => Promise<{ content?: string }> },
  prompt: string,
  compactionConfig?: CompactionConfig
): Promise<string> {
  const result = await aiClient.chat({
    messages: [
      {
        role: "user",
        content: getCompactionPromptTemplate(compactionConfig?.compactPrompt),
      },
      { role: "user", content: prompt },
    ],
    maxTokens: 2000,
    temperature: 0.3,
  });
  const summary = result.content?.trim();
  if (!summary) {
    throw new Error("压缩摘要为空");
  }
  return summary;
}
