import { z } from "zod";

import { IPC_MAX_PATH_CHARS, IpcPath } from "./ipcSchemaPrimitives";
import { SETTINGS_SECRET_MASK } from "./settingsSecretContract";

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
const ModelConfigInput = z
  .object({
    name: z.string().min(1).max(1_024),
    contextWindowSize: PositiveTokenCount.optional(),
    compHash: z.string().max(512).optional(),
    reasoningMode: ReasoningModeInput.optional(),
  })
  .strict();
const CustomHeadersInput = z
  .record(z.string().min(1).max(256), z.string().max(8_192))
  .superRefine((headers, ctx) => {
    if (Object.keys(headers).length > 64) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "自定义请求头最多包含 64 项",
      });
    }
  });
const AiProviderConfigInput = z
  .object({
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
  })
  .strict();
const AiProviderMapInput = z
  .record(z.string().min(1).max(256), AiProviderConfigInput)
  .superRefine((providers, ctx) => {
    if (Object.keys(providers).length > 50) {
      ctx.addIssue({
        code: z.ZodIssueCode.custom,
        message: "AI 提供商最多配置 50 个",
      });
    }
  });
const CompactionConfigInput = z
  .object({
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
    archiveRolloutAfterBytes: z
      .number()
      .int()
      .positive()
      .max(10 * 1024 * 1024 * 1024)
      .optional(),
    compactPrompt: z.string().max(100_000).optional(),
    compactionProvider: z.enum(["local", "remote"]).optional(),
    remoteCompactUrl: z.string().max(8_192).optional(),
    remoteCompactApiKey: SecretInput.optional(),
    remoteCompactModel: z.string().max(1_024).optional(),
    contextWindowSize: PositiveTokenCount.optional(),
  })
  .strict();
const PinnedFolderInput = z
  .object({
    path: IpcPath,
    name: z.string().min(1).max(255),
    addedAt: z.number().finite().nonnegative(),
    pinnedFiles: z.array(IpcPath).max(1_000).optional(),
  })
  .strict();

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

export const SettingsSetInput = z
  .tuple([SettingsKeyInput, z.unknown()])
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
