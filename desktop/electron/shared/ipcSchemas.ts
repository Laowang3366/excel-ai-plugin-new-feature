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

export const AppOpenPathInput = z.string().min(1, "路径不能为空");
export const AppOpenExternalInput = z.string().url("URL 格式不正确");
export const LaunchOfficeApplicationInput = z.enum(["wps", "excel", "word", "powerpoint"]);
export type LaunchOfficeApplicationInput = z.infer<typeof LaunchOfficeApplicationInput>;

// ============================================================
// Window
// ============================================================

export const SetAlwaysOnTopInput = z.boolean();
export const WindowDisplayModeInput = z.enum(["normal", "compact"]);

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
  expand: z.enum(["none", "spill", "currentArray", "currentRegion"]).optional(),
});

export const ExcelWriteRangeInput = z.object({
  sheetName: z.string(),
  range: z.string(),
  values: z.array(z.array(z.unknown())),
});

export const ExcelSelectHostInput = z.enum(["excel", "wps"]);

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

export const AgentInterruptInput = z.object({
  threadId: z.string().min(1).nullable().optional(),
}).optional();
export type AgentInterruptInput = z.infer<typeof AgentInterruptInput>;

// ============================================================
// Thread
// ============================================================

export const ThreadIdInput = z.string().min(1);
export const ThreadNewInput = z.string().optional();
export const ThreadUpdateMetadataInput = z.object({
  threadId: z.string().min(1),
  patch: z.object({
    name: z.string().max(200).optional(),
    folderId: z.string().optional(),
    modelProvider: z.string().min(1).optional(),
    model: z.string().optional(),
  }).strict(),
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
export const FolderPathsInput = z.array(FolderPathInput).min(1, "文件夹列表不能为空").max(100, "一次最多读取 100 个文件夹");
export const FileWriteTempFileInput = z.object({
  prefix: z.string().max(64).optional(),
  suffix: z.string().regex(/^\.[a-zA-Z0-9]{1,16}$/).optional(),
  data: z.string().min(1, "data 不能为空"),
});
export const OcrRecognizeInput = z.object({
  mode: z.enum(["image", "invoice"]).optional(),
  filePaths: z.array(z.string().min(1)).min(1, "文件列表不能为空"),
});

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
// Sandbox
// ============================================================

const SandboxPatternToken = z.union([
  z.string(),
  z.array(z.string()).min(1),
]);

export const SandboxUserRuleInput = z.object({
  first: z.string().min(1).optional(),
  pattern: z.array(SandboxPatternToken).min(1).optional(),
  rest: z.array(SandboxPatternToken).optional(),
  decision: z.enum(["allow", "prompt", "forbidden"]),
  justification: z.string().optional(),
}).refine((rule) => Boolean(rule.first || rule.pattern?.[0]), {
  message: "规则必须包含 first 或 pattern",
});

export const SandboxUserRulesInput = z.array(SandboxUserRuleInput);
export const SandboxWritableRootsInput = z.array(z.string().min(1));

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
