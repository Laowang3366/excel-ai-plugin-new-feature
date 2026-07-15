/**
 * IPC Schema 定义 — 单一事实来源
 *
 * 所有 IPC 通道的请求参数和响应类型在此定义。
 * ipcHandlers.ts 中使用 validateInput 在运行时校验参数。
 * preload.ts 和 electronApi.d.ts 的类型从此处推导，消除手工同步。
 */

import { z } from "zod";

import {
  DEFAULT_IPC_JSON_RESOURCE_BUDGET,
  inspectJsonResourceBudget,
} from "./jsonResourceBudget";
import { SETTINGS_SECRET_MASK } from "./settingsSecretContract";

// ============================================================
// 通用
// ============================================================

export const IPC_MAX_PATH_CHARS = 32_767;
export const IPC_MAX_CHAT_CONTENT_CHARS = 50_000;
export const IPC_MAX_RESUME_CONTEXT_CHARS = 200_000;
export const IPC_MAX_ATTACHMENTS = 20;
export const IPC_MAX_OCR_FILES = 20;
export const IPC_MAX_EXCEL_CELLS = 100_000;
export const IPC_MAX_EXCEL_ROWS = 10_000;
export const IPC_MAX_EXCEL_COLUMNS = 16_384;
export const IPC_MAX_CELL_TEXT_CHARS = 32_767;
export const IPC_MAX_FILE_TRANSFER_BYTES = 50 * 1024 * 1024;
export const IPC_MAX_BASE64_CHARS = Math.ceil(IPC_MAX_FILE_TRANSFER_BYTES / 3) * 4;

const IpcPath = z.string().min(1, "路径不能为空").max(IPC_MAX_PATH_CHARS, "路径过长");

export function estimateBase64DecodedBytes(value: string): number {
  if (!value) return 0;
  const padding = value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor(value.length * 3 / 4) - padding);
}

export function isBase64PayloadWithinLimit(
  value: string,
  maxBytes = IPC_MAX_FILE_TRANSFER_BYTES
): boolean {
  return estimateBase64DecodedBytes(value) <= maxBytes;
}

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

export const MigrateDataPathInput = IpcPath;
export type MigrateDataPathInput = z.infer<typeof MigrateDataPathInput>;

export const AppOpenPathInput = IpcPath;
export const AppOpenExternalInput = z.string().max(8_192).url("URL 格式不正确");
export const AppLogInput = z.object({
  level: z.enum(["debug", "info", "warn", "error"]),
  tag: z.string().max(128),
  message: z.string().max(50_000),
});
export const LaunchOfficeApplicationInput = z.enum(["wps", "excel", "word", "powerpoint"]);
export type LaunchOfficeApplicationInput = z.infer<typeof LaunchOfficeApplicationInput>;

export const UpdateKindInput = z.enum(["installer", "hotPatch"]);
export const UpdateCheckInput = z.boolean().optional();

// ============================================================
// Window
// ============================================================

export const SetAlwaysOnTopInput = z.boolean();
export const WindowDisplayModeInput = z.enum(["normal", "compact"]);

// ============================================================
// Settings
// ============================================================

export const SettingsKeyInput = z.enum([
  "activeProvider",
  "aiProviders",
  "closeToTray",
  "compactionConfig",
  "dataStoragePath",
  "dynamicArrayFunctionsEnabled",
  "knowledgeEnabled",
  "language",
  "mineruApiToken",
  "ocrMineruApiToken",
  "officeAutoCompactEnabled",
  "permissionMode",
  "pinnedFolders",
  "remoteDataProcessingEnabled",
  "showReasoning",
  "theme",
  "windowOpacity",
]);
export const SettingsGetInput = SettingsKeyInput;

const SettingsText = z.string().max(32_768);
const SettingsIdentifier = z.string().max(256);
const ReasoningModeInput = z.enum(["off", "low", "medium", "high", "max"]);
const PositiveTokenCount = z.number().int().positive().max(100_000_000);
const SecretInput = z.union([z.literal(SETTINGS_SECRET_MASK), SettingsText]);
const ModelConfigInput = z.object({
  name: z.string().min(1).max(1_024),
  contextWindowSize: PositiveTokenCount.optional(),
  compHash: z.string().max(512).optional(),
  reasoningMode: ReasoningModeInput.optional(),
}).strict();
const CustomHeadersInput = z.record(
  z.string().min(1).max(256),
  z.string().max(8_192),
).superRefine((headers, ctx) => {
  if (Object.keys(headers).length > 64) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "自定义请求头最多包含 64 项",
    });
  }
});
const AiProviderConfigInput = z.object({
  id: z.string().min(1).max(256),
  name: z.string().max(512),
  provider: z.string().max(256),
  apiKey: SecretInput,
  baseUrl: z.string().max(8_192),
  model: z.string().max(1_024),
  models: z.array(z.string().max(1_024)).max(1_000).optional(),
  modelConfigs: z.array(ModelConfigInput).max(1_000).optional(),
  defaultBaseUrl: z.string().max(8_192).optional(),
  defaultModel: z.string().max(1_024).optional(),
  apiFormat: z.string().max(128).optional(),
  customHeaders: CustomHeadersInput.optional(),
  contextWindowSize: PositiveTokenCount.optional(),
  compHash: z.string().max(512).optional(),
  reasoningMode: ReasoningModeInput.optional(),
}).strict();
const AiProviderMapInput = z.record(
  z.string().min(1).max(256),
  AiProviderConfigInput,
).superRefine((providers, ctx) => {
  if (Object.keys(providers).length > 50) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      message: "AI 提供商最多配置 50 个",
    });
  }
});
const CompactionConfigInput = z.object({
  enabled: z.boolean().optional(),
  autoCompactThresholdPercent: z.number().finite().min(1).max(100).optional(),
  autoCompactTokenThreshold: PositiveTokenCount.optional(),
  midTurnThresholdRatio: z.number().finite().min(0.1).max(1).optional(),
  retainedUserMessageMaxTokens: PositiveTokenCount.optional(),
  retainedRecentItemCount: z.number().int().nonnegative().max(100_000).optional(),
  summaryRetryCount: z.number().int().nonnegative().max(20).optional(),
  summaryRetryBaseDelayMs: z.number().int().nonnegative().max(3_600_000).optional(),
  summaryRetryMaxDelayMs: z.number().int().nonnegative().max(3_600_000).optional(),
  summaryRetryBackoffFactor: z.number().finite().min(1).max(100).optional(),
  archiveRolloutAfterBytes: z.number().int().positive().max(10 * 1024 * 1024 * 1024).optional(),
  compactPrompt: z.string().max(100_000).optional(),
  compactionProvider: z.enum(["local", "remote"]).optional(),
  remoteCompactUrl: z.string().max(8_192).optional(),
  remoteCompactApiKey: SecretInput.optional(),
  remoteCompactModel: z.string().max(1_024).optional(),
  contextWindowSize: PositiveTokenCount.optional(),
}).strict();
const PinnedFolderInput = z.object({
  path: IpcPath,
  name: z.string().min(1).max(255),
  addedAt: z.number().finite().nonnegative(),
  pinnedFiles: z.array(IpcPath).max(1_000).optional(),
}).strict();

type SettingsKey = z.infer<typeof SettingsKeyInput>;
const SettingsValueSchemas: Record<SettingsKey, z.ZodType> = {
  activeProvider: SettingsIdentifier,
  aiProviders: AiProviderMapInput,
  closeToTray: z.boolean(),
  compactionConfig: CompactionConfigInput,
  dataStoragePath: z.string().max(IPC_MAX_PATH_CHARS),
  dynamicArrayFunctionsEnabled: z.boolean(),
  knowledgeEnabled: z.boolean(),
  language: z.enum(["zh-CN", "en-US"]),
  mineruApiToken: SecretInput,
  ocrMineruApiToken: SecretInput,
  officeAutoCompactEnabled: z.boolean(),
  permissionMode: z.enum(["normal", "auto_approve_safe", "confirm_all"]),
  pinnedFolders: z.array(PinnedFolderInput).max(100),
  remoteDataProcessingEnabled: z.boolean(),
  showReasoning: z.boolean(),
  theme: z.enum(["light", "dark"]),
  windowOpacity: z.number().finite().min(0.55).max(1),
};

export const SettingsSetInput = z.tuple([SettingsKeyInput, z.unknown()])
  .superRefine(([key, value], ctx) => {
    const result = SettingsValueSchemas[key].safeParse(value);
    if (result.success) return;
    for (const issue of result.error.issues) {
      ctx.addIssue({
        ...issue,
        path: [1, ...issue.path],
      });
    }
  });

// ============================================================
// Excel
// ============================================================

export const ExcelReadRangeInput = z.object({
  sheetName: z.string().min(1).max(255),
  range: z.string().min(1).max(8_192),
  expand: z.enum(["none", "spill", "currentArray", "currentRegion"]).optional(),
});

const ExcelCellValue = z.union([
  z.string().max(IPC_MAX_CELL_TEXT_CHARS),
  z.number().finite(),
  z.boolean(),
  z.null(),
]);

export const ExcelWriteRangeInput = z.object({
  sheetName: z.string().min(1).max(255),
  range: z.string().min(1).max(8_192),
  values: z.array(
    z.array(ExcelCellValue).max(IPC_MAX_EXCEL_COLUMNS, "单行列数超过 Excel 上限")
  ).min(1, "写入矩阵不能为空").max(IPC_MAX_EXCEL_ROWS, "写入行数超过 IPC 上限"),
}).superRefine((input, ctx) => {
  const cellCount = input.values.reduce((total, row) => total + row.length, 0);
  if (cellCount > IPC_MAX_EXCEL_CELLS) {
    ctx.addIssue({
      code: z.ZodIssueCode.custom,
      path: ["values"],
      message: `写入矩阵最多包含 ${IPC_MAX_EXCEL_CELLS} 个单元格`,
    });
  }
});

export const ExcelSelectHostInput = z.enum(["excel", "wps"]);

// ============================================================
// Office 自动化管理
// ============================================================

export const OfficeAutomationAppInput = z.enum(["excel", "word", "presentation"]);
export const OfficeAutomationDocumentsListInput = z.object({ app: OfficeAutomationAppInput.optional() }).optional();
export const OfficeAutomationDocumentInput = z.object({
  app: OfficeAutomationAppInput,
  filePath: IpcPath,
  instanceId: z.string().min(1).max(256).optional(),
});
export const OfficeAutomationObjectsListInput = OfficeAutomationDocumentInput.extend({
  kind: z.string().min(1).max(256).optional(),
});
export const OfficeAutomationObjectActivateInput = OfficeAutomationDocumentInput.extend({
  locator: z.string().min(1).max(8_192),
});
export const OfficeAutomationIdInput = z.object({ id: z.string().uuid() });
export const OfficeAutomationForceInput = OfficeAutomationIdInput.extend({ force: z.boolean().optional() });
export const OfficeAutomationTemplateSaveInput = z.object({
  workflowId: z.string().uuid(),
  templateId: z.string().uuid().optional(),
  name: z.string().trim().min(1).max(120),
  description: z.string().trim().max(500).optional(),
});
export const OfficeAutomationTemplateRunInput = z.object({
  templateId: z.string().uuid(),
  variables: z.record(z.string().max(256), z.unknown()).superRefine((variables, ctx) => {
    const violation = inspectJsonResourceBudget(
      variables,
      DEFAULT_IPC_JSON_RESOURCE_BUDGET,
    );
    if (violation) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        path: violation.path,
        message: violation.message,
      });
    }
  }).optional(),
});

// ============================================================
// Agent
// ============================================================

export const AgentFileAttachment = z.object({
  filePath: IpcPath,
  fileName: z.string().min(1).max(255),
  fileType: z.enum(["image", "document"]),
  size: z.number().finite().nonnegative().max(IPC_MAX_FILE_TRANSFER_BYTES).optional(),
});

export const AgentStartTurnInput = z.object({
  content: z.string().max(IPC_MAX_CHAT_CONTENT_CHARS),
  attachments: z.array(AgentFileAttachment).max(IPC_MAX_ATTACHMENTS).optional(),
  clientId: z.string().max(256).optional(),
  threadId: z.string().max(256).optional().nullable(),
  isResume: z.boolean().optional(),
  resumeContext: z.string().max(IPC_MAX_RESUME_CONTEXT_CHARS).optional(),
});
export type AgentStartTurnInput = z.infer<typeof AgentStartTurnInput>;

export const AgentContinueTurnInput = z.object({
  content: z.string().max(IPC_MAX_CHAT_CONTENT_CHARS),
  attachments: z.array(AgentFileAttachment).max(IPC_MAX_ATTACHMENTS).optional(),
  clientId: z.string().max(256).optional(),
  threadId: z.string().max(256).optional().nullable(),
});
export type AgentContinueTurnInput = z.infer<typeof AgentContinueTurnInput>;

export const AgentInterruptInput = z.object({
  threadId: z.string().min(1).max(256).nullable().optional(),
}).optional();
export type AgentInterruptInput = z.infer<typeof AgentInterruptInput>;

// ============================================================
// Thread
// ============================================================

export const ThreadIdInput = z.string().min(1).max(256);
export const ThreadNewInput = z.string().max(256).optional();
export const ThreadUpdateMetadataInput = z.object({
  threadId: z.string().min(1).max(256),
  patch: z.object({
    name: z.string().max(200).optional(),
    folderId: z.string().max(256).optional(),
    modelProvider: z.string().min(1).max(256).optional(),
    model: z.string().max(1_024).optional(),
  }).strict(),
});
export const ThreadGraphEdgeInput = z.object({
  parentThreadId: z.string().min(1).max(256),
  childThreadId: z.string().min(1).max(256),
  label: z.string().max(500).optional(),
});
export const ThreadGraphCloseEdgeInput = z.object({
  parentThreadId: z.string().min(1).max(256),
  childThreadId: z.string().min(1).max(256),
});
export const ThreadGraphListDescendantsInput = z.object({
  parentThreadId: z.string().min(1).max(256),
  status: z.enum(["open", "closed", "all"]).optional(),
});

// ============================================================
// File & Folder
// ============================================================

export const FilePathInput = IpcPath;
export const FolderPathInput = IpcPath;
export const FolderPathsInput = z.array(FolderPathInput).min(1, "文件夹列表不能为空").max(100, "一次最多读取 100 个文件夹");
export const FileWriteTempFileInput = z.object({
  prefix: z.string().max(64).optional(),
  suffix: z.string().regex(/^\.[a-zA-Z0-9]{1,16}$/).optional(),
  data: z.string()
    .min(1, "data 不能为空")
    .max(IPC_MAX_BASE64_CHARS, "Base64 编码内容过大")
    .refine(
      (value) => isBase64PayloadWithinLimit(value),
      `临时文件最大支持 ${IPC_MAX_FILE_TRANSFER_BYTES / 1024 / 1024}MB`
    ),
});
export const OcrRecognizeInput = z.object({
  mode: z.enum(["image", "invoice"]).optional(),
  filePaths: z.array(IpcPath).min(1, "文件列表不能为空").max(IPC_MAX_OCR_FILES, `一次最多识别 ${IPC_MAX_OCR_FILES} 个文件`),
});

// ============================================================
// AI
// ============================================================

export const AiListModelsInput = z.object({
  baseUrl: z.string().max(8_192),
  apiKey: z.string().max(32_768),
  apiFormat: z.string().max(128),
  providerId: z.string().min(1).max(256).optional(),
});

export const AiTestConnectionInput = z.object({
  baseUrl: z.string().max(8_192),
  apiKey: z.string().max(32_768),
  apiFormat: z.string().max(128),
  model: z.string().max(1_024),
  providerId: z.string().min(1).max(256).optional(),
});

// ============================================================
// Tool Approval
// ============================================================

export const ToolConfirmInput = z.object({
  toolCallId: z.string().max(256),
  alwaysAllow: z.boolean().optional(),
});

export const ToolCancelInput = z.string().max(256);

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
  filePath: IpcPath,
});

export const KnowledgeIndexFolderInput = z.object({
  folderPath: IpcPath,
});

export const KnowledgeDeleteInput = z.object({
  sourcePath: IpcPath,
});

export const KnowledgeSearchInput = z.object({
  query: z.string().min(1, "搜索关键词不能为空").max(10_000, "搜索关键词过长"),
  topK: z.number().int().min(1).max(50).optional(),
});

export type KnowledgeIndexFileInput = z.infer<typeof KnowledgeIndexFileInput>;
export type KnowledgeIndexFolderInput = z.infer<typeof KnowledgeIndexFolderInput>;
export type KnowledgeDeleteInput = z.infer<typeof KnowledgeDeleteInput>;
export type KnowledgeSearchInput = z.infer<typeof KnowledgeSearchInput>;
