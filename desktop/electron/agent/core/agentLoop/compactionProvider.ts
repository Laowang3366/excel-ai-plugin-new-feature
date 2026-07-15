/**
 * 压缩摘要 Provider。
 *
 * 关联模块：
 * - agentLoop.ts: 只依赖 CompactionProvider 接口，避免关心本地或远程实现。
 * - summaryGenerator.ts: 本地 Provider 复用现有 AI 客户端摘要生成逻辑。
 */

import type { CompactionConfig } from "../../shared/types";
import { getCompactionPromptTemplate } from "../../prompts/compactionPrompt";
import { generateSummary } from "./summaryGenerator";

export interface CompactionSummaryInput {
  historyPrompt: string;
  config?: CompactionConfig;
}

export interface CompactionProvider {
  generateSummary(input: CompactionSummaryInput): Promise<string>;
}

export interface RemoteCompactionProviderConfig {
  endpoint: string;
  apiKey?: string;
  model?: string;
}

export function createCompactionProvider(
  aiClient: { chat: (params: any) => Promise<{ content?: string }> },
  config?: CompactionConfig,
): CompactionProvider {
  if (config?.compactionProvider === "remote" && config.remoteCompactUrl) {
    return createRemoteCompactionProvider({
      endpoint: config.remoteCompactUrl,
      apiKey: config.remoteCompactApiKey,
      model: config.remoteCompactModel,
    });
  }
  return createLocalCompactionProvider(aiClient);
}

export function createLocalCompactionProvider(aiClient: {
  chat: (params: any) => Promise<{ content?: string }>;
}): CompactionProvider {
  return {
    generateSummary(input) {
      return generateSummary(aiClient, input.historyPrompt, input.config);
    },
  };
}

export function createRemoteCompactionProvider(
  config: RemoteCompactionProviderConfig,
): CompactionProvider {
  return {
    async generateSummary(input) {
      const headers: Record<string, string> = {
        "Content-Type": "application/json",
      };
      if (config.apiKey) {
        headers.Authorization = `Bearer ${config.apiKey}`;
      }

      const response = await fetch(config.endpoint, {
        method: "POST",
        headers,
        body: JSON.stringify({
          instruction: getCompactionPromptTemplate(input.config?.compactPrompt),
          input: input.historyPrompt,
          model: config.model,
        }),
      });

      if (!response.ok) {
        const text = await response.text();
        throw new Error(`远程压缩失败 (${response.status}): ${formatRemoteError(text)}`);
      }

      const json = (await response.json()) as any;
      const summary = json.summary || json.content || json.choices?.[0]?.message?.content;
      if (typeof summary !== "string" || !summary.trim()) {
        throw new Error("远程压缩返回空摘要");
      }
      return summary.trim();
    },
  };
}

function formatRemoteError(text: string): string {
  const compact = text
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return (compact || text).slice(0, 240);
}
