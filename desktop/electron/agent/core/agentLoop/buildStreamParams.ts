/**
 * 构建 AI 流式请求参数
 *
 * 从 AgentLoop.runAgentLoop 中提取的逻辑：
 * - 推理力度降级策略
 * - 系统提示词构建（含文件夹上下文）
 */

import { type ReasoningMode } from "../../providers/aiClient";
import { type AIClientConfig } from "../../providers/aiClient";
import {
  buildSystemPrompt,
  appendFolderContext,
  buildContextualPromptSections,
  type FolderFileItem,
  type PromptBuildContext,
} from "../../prompts/systemPrompt";
import * as fs from "fs";
import * as path from "path";
import { type StreamParams } from "./streamCollector";
import type { RuntimeLongTermMemoryRecord } from "../../memory/stateRuntimeTypes";
import type { StateRuntimeStore } from "../../memory/stateRuntimeStore";
import { isToolWritableMemoryKind } from "../../memory/longTerm/memoryTypes";
import {
  getWordBridge,
  getPresentationBridge,
} from "../../runtime/bridgeRegistry";

// ============================================================
// 推理力度（不再递减）
// ============================================================

/**
 * 计算当前轮次的有效推理力度
 *
 * 旧逻辑：round 2+ 自动降一档（max→high, high→medium…），目的是节省 token。
 * 问题：推理强度切换导致上下文连贯性断裂 — 第一轮深度思考写入了公式，
 *       第二轮降级后模型无法维持同等推理质量，反而反复纠结"公式对不对"
 *       而不调用 range.read 验证，浪费更多轮次。
 * 新逻辑：全程保持用户配置的推理力度，不再递减。
 *
 * @param configuredMode - 用户配置的推理力度
 * @param _round - 当前轮次编号（保留参数签名兼容，不再使用）
 * @returns 用户配置的推理力度
 */
export function getEffectiveReasoningMode(
  configuredMode: ReasoningMode,
  _round: number
): ReasoningMode {
  return configuredMode;
}

export function appendLongTermMemoryContext(
  prompt: string,
  memories: RuntimeLongTermMemoryRecord[]
): string {
  const visible = memories.filter((memory) =>
    memory.visibility === "user" &&
    memory.status === "active" &&
    isToolWritableMemoryKind(memory.kind)
  );
  if (visible.length === 0) return prompt;

  const lines = visible.slice(0, 8).map((memory) =>
    `- [${memory.kind}] ${memory.summary || memory.content}`
  );
  return `${prompt}\n\n## 用户长期记忆\n${lines.join("\n")}`;
}

export async function appendRuntimeLongTermMemoryContext(
  prompt: string,
  store: Pick<StateRuntimeStore, "listLongTermMemories"> | undefined
): Promise<string> {
  if (!store) return prompt;

  try {
    const allowed: RuntimeLongTermMemoryRecord[] = [];
    const pageSize = 8;
    for (let offset = 0; allowed.length < pageSize; offset += pageSize) {
      const memories = await store.listLongTermMemories({
        visibility: "user",
        status: "active",
        limit: pageSize,
        offset,
      });
      allowed.push(
        ...memories.filter((memory) => isToolWritableMemoryKind(memory.kind)),
      );
      if (memories.length < pageSize) break;
    }
    return appendLongTermMemoryContext(prompt, allowed.slice(0, pageSize));
  } catch {
    return prompt;
  }
}

export function appendRuntimeDateContext(prompt: string, now: Date = new Date()): string {
  const dateFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    weekday: "long",
  });
  const timeFormatter = new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  });
  return `${prompt}\n\n## 运行时上下文\n- 当前日期：${dateFormatter.format(now)}\n- 当前时间：${timeFormatter.format(now)}（Asia/Shanghai）\n- 处理“今天、昨天、最近、近 N 日、最新、今年”等相对时间时，以以上日期和时区为准；搜索时不要自行补入过期年份。`;
}

// ============================================================
// 系统提示词构建
// ============================================================

/**
 * 构建有效系统提示词（含动态文件夹上下文）
 *
 * @param basePrompt - 静态基础提示词
 * @param folderId - 关联的文件夹 ID（可选）
 * @returns 构建后的系统提示词
 */
export async function buildEffectiveSystemPrompt(
  basePrompt: string | undefined,
  folderId?: string,
  turnContext?: PromptBuildContext
): Promise<string> {
  let effectivePrompt = basePrompt || buildSystemPrompt();
  const contextualPrompt = buildContextualPromptSections({
    ...turnContext,
    folderId,
  });
  if (contextualPrompt) {
    effectivePrompt += "\n\n" + contextualPrompt;
  }

  // ── 注入文件夹上下文 ──
  if (folderId) {
    try {
      const folderName = folderId.split(/[\\/]/).pop() || folderId;
      const entries = await fs.promises.readdir(folderId, { withFileTypes: true });
      const excelExts = new Set([".xlsx", ".xls", ".xlsm", ".xlsb", ".csv"]);
      const filesPromises = entries
        .filter((e) => e.isFile() && excelExts.has(path.extname(e.name).toLowerCase()))
        .map(async (e) => {
          const fp = path.join(folderId, e.name);
          let size = 0;
          try { size = (await fs.promises.stat(fp)).size; } catch { /* ignore */ }
          return { fileName: e.name, filePath: fp, size };
        });
      const files: FolderFileItem[] = await Promise.all(filesPromises);
      effectivePrompt = appendFolderContext(effectivePrompt, folderId, folderName, files);
    } catch {
      // 文件夹路径不可访问时静默跳过
    }
  }

  // ── 注入 Office 连接状态 ──
  effectivePrompt = effectivePrompt.replace(
    /\{\{OFFICE_CONNECTION_STATUS\}\}/g,
    buildOfficeConnectionStatus()
  );

  return appendRuntimeDateContext(effectivePrompt);
}

/** 构建 Office 连接状态描述，供系统提示词使用 */
function buildOfficeConnectionStatus(): string {
  const wordBridge = getWordBridge();
  const pptBridge = getPresentationBridge();
  const parts: string[] = [];

  if (wordBridge?.isConnected()) {
    parts.push(`Word (${wordBridge.getHost()})`);
  } else {
    parts.push("Word(未连接)");
  }

  if (pptBridge?.isConnected()) {
    parts.push(`PPT (${pptBridge.getHost()})`);
  } else {
    parts.push("PPT(未连接)");
  }

  return parts.join(" | ");
}
