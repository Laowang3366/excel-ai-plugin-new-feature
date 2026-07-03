/**
 * IPC Schema 定义 — 单一事实来源
 *
 * 所有 IPC 通道的请求参数和响应类型在此定义。
 * ipcHandlers.ts 中使用 validateInput 在运行时校验参数。
 * preload.ts 和 electronApi.d.ts 的类型从此处推导，消除手工同步。
 */

import { z } from "zod";

// ============================================================
// 通用
// ============================================================

/** 运行时校验工具：校验失败抛出带有字段描述的 Error */
export function validateInput<T>(schema: z.ZodType<T>, data: unknown): T {
  const result = schema.safeParse(data);
  if (!result.success) {
    const issues = result.error.issues
      .map((i) => `${i.path.join(".")}: ${i.message}`)
      .join("; ");
    throw new Error(`IPC 参数校验失败: ${issues}`);
  }
  return result.data;
}

// ============================================================
// App
// ============================================================

export const MigrateDataPathInput = z.string().min(1, "路径不能为空");
export type MigrateDataPathInput = z.infer<typeof MigrateDataPathInput>;

export const SelectDataPathOutput = z.object({
  canceled: z.boolean(),
  filePaths: z.array(z.string()),
});

// ============================================================
// Window
// ============================================================

export const SetAlwaysOnTopInput = z.boolean();

// ============================================================
// Settings
// ============================================================

export const SettingsGetInput = z.string();
export const SettingsSetInput = z.tuple([z.string(), z.unknown()]);

// ============================================================
// Excel
// ============================================================

export const ExcelReadRangeInput = z.object({
  sheetName: z.string(),
  range: z.string(),
});

export const ExcelWriteRangeInput = z.object({
  sheetName: z.string(),
  range: z.string(),
  values: z.array(z.array(z.unknown())),
});

// ============================================================
// Agent
// ============================================================

export const AgentFileAttachment = z.object({
  filePath: z.string(),
  fileName: z.string(),
  fileType: z.enum(["image", "document"]),
  size: z.number().optional(),
});

export const AgentStartTurnInput = z.object({
  content: z.string(),
  attachments: z.array(AgentFileAttachment).optional(),
  clientId: z.string().optional(),
  threadId: z.string().optional().nullable(),
  isResume: z.boolean().optional(),
  resumeContext: z.string().optional(),
});
export type AgentStartTurnInput = z.infer<typeof AgentStartTurnInput>;

export const AgentContinueTurnInput = z.object({
  content: z.string(),
  attachments: z.array(AgentFileAttachment).optional(),
  clientId: z.string().optional(),
  threadId: z.string().optional().nullable(),
});
export type AgentContinueTurnInput = z.infer<typeof AgentContinueTurnInput>;

// ============================================================
// Thread
// ============================================================

export const ThreadIdInput = z.string().min(1);
export const ThreadNewInput = z.string().optional();
export const ThreadUpdateMetadataInput = z.object({
  threadId: z.string(),
  patch: z.record(z.string(), z.unknown()),
});
export const ThreadGraphEdgeInput = z.object({
  parentThreadId: z.string().min(1),
  childThreadId: z.string().min(1),
  label: z.string().optional(),
});
export const ThreadGraphCloseEdgeInput = z.object({
  parentThreadId: z.string().min(1),
  childThreadId: z.string().min(1),
});
export const ThreadGraphListDescendantsInput = z.object({
  parentThreadId: z.string().min(1),
  status: z.enum(["open", "closed", "all"]).optional(),
});

// ============================================================
// File & Folder
// ============================================================

export const FilePathInput = z.string().min(1);
export const FolderPathInput = z.string().min(1);

// ============================================================
// AI
// ============================================================

export const AiListModelsInput = z.object({
  baseUrl: z.string(),
  apiKey: z.string(),
  apiFormat: z.string(),
});

export const AiTestConnectionInput = z.object({
  baseUrl: z.string(),
  apiKey: z.string(),
  apiFormat: z.string(),
  model: z.string(),
});

// ============================================================
// Tool Approval
// ============================================================

export const ToolConfirmInput = z.object({
  toolCallId: z.string(),
  alwaysAllow: z.boolean().optional(),
});

export const ToolCancelInput = z.string();

// ============================================================
// Stats
// ============================================================

export const StatsGetSummaryInput = z.object({
  days: z.number().int().min(1).max(365).optional(),
}).optional();

// ============================================================
// Knowledge (RAG)
// ============================================================

export const KnowledgeIndexFileInput = z.object({
  filePath: z.string().min(1, "文件路径不能为空"),
});

export const KnowledgeIndexFolderInput = z.object({
  folderPath: z.string().min(1, "文件夹路径不能为空"),
});

export const KnowledgeDeleteInput = z.object({
  sourcePath: z.string().min(1, "来源路径不能为空"),
});

export const KnowledgeSearchInput = z.object({
  query: z.string().min(1, "搜索关键词不能为空"),
  topK: z.number().int().min(1).max(50).optional(),
});

export type KnowledgeIndexFileInput = z.infer<typeof KnowledgeIndexFileInput>;
export type KnowledgeIndexFolderInput = z.infer<typeof KnowledgeIndexFolderInput>;
export type KnowledgeDeleteInput = z.infer<typeof KnowledgeDeleteInput>;
export type KnowledgeSearchInput = z.infer<typeof KnowledgeSearchInput>;
