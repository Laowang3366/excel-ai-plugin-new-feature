/**
 * Build AI stream request parameters.
 *
 * Extracted from AgentLoop.runAgentLoop:
 * - system prompt assembly with folder context
 */

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
import { getAgentGlobalSettings } from "../../runtime/agentGlobalSettings";

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

export function appendDynamicArrayFunctionSupportContext(
  prompt: string,
  enabled = getAgentGlobalSettings().dynamicArrayFunctionsEnabled
): string {
  const settingLine = enabled
    ? "动态数组函数环境支持：已开启。生成 Excel/WPS 公式时默认允许 FILTER、UNIQUE、SORT、SEQUENCE、LET、XLOOKUP 等动态数组函数；不要反复质疑当前环境是否适配动态数组函数，除非工具回读明确返回 #NAME? 或用户关闭此设置。"
    : "动态数组函数环境支持：已关闭。生成 Excel/WPS 公式时不要依赖动态数组 spill，优先使用逐格独立公式、传统函数或辅助区域。";
  return `${prompt}\n\n## 公式函数环境设置\n- ${settingLine}`;
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

  effectivePrompt = appendDynamicArrayFunctionSupportContext(effectivePrompt);

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
