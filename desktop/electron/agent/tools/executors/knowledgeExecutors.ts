/**
 * 知识库工具执行器
 *
 * 关联 knowledge/retriever 与 knowledge/writer，注册知识库搜索与写入。
 */

import type { ToolExecutor } from "../../shared/types";
import type { Retriever } from "../../knowledge/retriever";
import type { KnowledgeWriter } from "../../knowledge/knowledgeWriter";
import { getKnowledgeRetriever, getKnowledgeWriter } from "../../knowledge/knowledgeRegistry";
import { validateArgs } from "./validation";

export interface KnowledgeExecutorDeps {
  knowledgeRetriever?: Retriever;
  knowledgeWriter?: KnowledgeWriter;
}

export function addKnowledgeExecutors(target: Map<string, ToolExecutor>, deps: KnowledgeExecutorDeps): void {
  target.set("knowledge.search", {
    name: "knowledge.search",
    execute: async (args: Record<string, unknown>) => {
      const retriever = getKnowledgeRetriever() ?? deps.knowledgeRetriever;
      if (!retriever) {
        return {
          success: false,
          error: "知识库尚未初始化，请在设置中配置 AI 供应商并添加知识来源",
        };
      }
      const err = validateArgs(args, { query: "string" });
      if (err) return { success: false, error: err };
      const query = args.query as string;
      const topK = typeof args.topK === "number" ? args.topK : 5;
      try {
        const results = await retriever.search({ text: query, topK });
        const formatted = retriever.formatForToolResult(results);
        return { success: true, data: formatted };
      } catch (e: any) {
        return { success: false, error: `知识搜索失败: ${e.message}` };
      }
    },
  });

  target.set("knowledge.write", {
    name: "knowledge.write",
    execute: async (args: Record<string, unknown>) => {
      const writer = getKnowledgeWriter() ?? deps.knowledgeWriter;
      if (!writer) {
        return {
          success: false,
          error: "知识库尚未初始化，请在设置中配置 AI 供应商后再写入知识库",
        };
      }
      const err = validateArgs(args, { content: "string" });
      if (err) return { success: false, error: err };
      const optionalErr = validateOptionalWriteArgs(args);
      if (optionalErr) return { success: false, error: optionalErr };
      try {
        const result = await writer.writeNote({
          content: args.content as string,
          title: typeof args.title === "string" ? args.title : undefined,
          tags: Array.isArray(args.tags) ? args.tags.filter((tag): tag is string => typeof tag === "string") : undefined,
          sourceName: typeof args.sourceName === "string" ? args.sourceName : undefined,
          metadata: { source: "tool:knowledge.write" },
        });
        return {
          success: true,
          data: {
            message: "已写入知识库",
            ...result,
          },
        };
      } catch (e: any) {
        return { success: false, error: `知识库写入失败: ${e.message}` };
      }
    },
  });
}

function validateOptionalWriteArgs(args: Record<string, unknown>): string | null {
  if (args.title !== undefined && typeof args.title !== "string") {
    return "参数 title 必须是 string";
  }
  if (args.sourceName !== undefined && typeof args.sourceName !== "string") {
    return "参数 sourceName 必须是 string";
  }
  if (args.tags !== undefined) {
    if (!Array.isArray(args.tags)) return "参数 tags 必须是 string 数组";
    if (args.tags.some((tag) => typeof tag !== "string")) {
      return "参数 tags 必须是 string 数组";
    }
  }
  return null;
}
