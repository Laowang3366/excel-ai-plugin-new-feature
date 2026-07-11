/**
 * ModelSettings i18n 文本与模板分组
 *
 * 从 ModelSettings.tsx 提取的共享常量：
 * - MODEL_TEXT: 中英文双语文本
 * - DIRECT_TEMPLATES / AGGREGATION_TEMPLATES / OTHER_TEMPLATES: 按分类分组的供应商模板
 *
 * 消费方：
 * - ModelSettings.tsx — 主编排器
 * - ProviderCard.tsx — 供应商卡片
 * - AddProviderDialog.tsx — 添加供应商弹窗
 * - DeleteConfirmDialog.tsx — 删除确认弹窗
 */

import { PROVIDER_TEMPLATES } from "../../store/settingsStore";
import { formatTokensAsK } from "../../utils/modelContextWindows";

// ============================================================
// 双语文本常量
// ============================================================

export const MODEL_TEXT = {
  "zh-CN": {
    title: "模型配置",
    desc: "管理你的 AI 供应商和模型配置。",
    addProvider: "添加供应商",
    empty: "还没有配置任何 AI 供应商",
    unset: "未设置",
    noModel: "未选择模型",
    active: "使用中",
    defaultValue: "默认",
    apiKey: "API 密钥",
    testConnection: "测试连接",
    test: "测试",
    testFailed: "测试失败",
    connectionOk: (latency?: number) => `连接成功${latency ? ` (${latency}ms)` : ""}`,
    connectionFailed: "连接失败",
    model: "模型",
    refreshModels: "刷新模型列表",
    availableModels: (count: number) => `可用模型: ${count} 个`,
    reasoningMode: "思考等级",
    reasoningModeHint: "控制模型的推理深度，等级越高思考越深入但响应越慢",
    enableThinking: "启用思考模式",
    thinkingHint: "开启后模型会进行深度思考，提升回答质量",
    setActive: "设为当前使用",
    addDialogTitle: "添加 AI 供应商",
    editDialogTitle: (name: string) => `编辑 ${name}`,
    save: "保存",
    providerType: "选择供应商类型",
    providerName: "供应商名称",
    providerNamePlaceholder: "如：我的 DeepSeek",
    apiFormat: "API 格式",
    apiUrl: "API 地址",
    customProvider: "自定义供应商",
    cancel: "取消",
    add: "添加",
    testing: "测试中...",
    addAndTest: "添加并测试",
    deleteTitle: "确认删除",
    deleteMessage: (name: string) => (
      <>
        确定要删除供应商 <strong>{name}</strong> 吗？此操作不可撤销。
      </>
    ),
    delete: "删除",
    // 新增文本
    contextWindowSize: "上下文窗口大小 (tokens)",
    contextWindowHint: (value: number) => `默认: ${formatTokensAsK(value)}`,
    testModel: "测试模型",
    directProviders: "直连供应商",
    aggregationProviders: "聚合平台",
    other: "其他",
    modelInputPlaceholder: "输入模型名称，如 deepseek-v4-pro",
    contextWindowPlaceholder: "未设置，默认 128k",
    modelList: "模型列表",
    modelListHint: "每个模型可独立设置上下文窗口大小。按回车或点击 + 添加模型。",
    addModel: "添加模型",
    removeModel: "移除模型",
  },
  "en-US": {
    title: "Models",
    desc: "Manage your AI providers and model settings.",
    addProvider: "Add provider",
    empty: "No AI providers configured yet",
    unset: "Not set",
    noModel: "No model selected",
    active: "Active",
    defaultValue: "Default",
    apiKey: "API key",
    testConnection: "Test connection",
    test: "Test",
    testFailed: "Test failed",
    connectionOk: (latency?: number) => `Connected${latency ? ` (${latency}ms)` : ""}`,
    connectionFailed: "Connection failed",
    model: "Model",
    refreshModels: "Refresh model list",
    availableModels: (count: number) => `${count} models available`,
    reasoningMode: "Thinking level",
    reasoningModeHint: "Controls reasoning depth. Higher levels think deeper but respond slower.",
    enableThinking: "Enable thinking mode",
    thinkingHint: "When enabled, the model thinks deeply to improve answer quality.",
    setActive: "Set as active",
    addDialogTitle: "Add AI provider",
    editDialogTitle: (name: string) => `Edit ${name}`,
    save: "Save",
    providerType: "Choose provider type",
    providerName: "Provider name",
    providerNamePlaceholder: "e.g. My DeepSeek",
    apiFormat: "API format",
    apiUrl: "API URL",
    customProvider: "Custom provider",
    cancel: "Cancel",
    add: "Add",
    testing: "Testing...",
    addAndTest: "Add and test",
    deleteTitle: "Confirm delete",
    deleteMessage: (name: string) => (
      <>
        Delete provider <strong>{name}</strong>? This cannot be undone.
      </>
    ),
    delete: "Delete",
    // New text
    contextWindowSize: "Context window size (tokens)",
    contextWindowHint: (value: number) => `Default: ${formatTokensAsK(value)}`,
    testModel: "Test model",
    directProviders: "Direct providers",
    aggregationProviders: "Aggregation platforms",
    other: "Other",
    modelInputPlaceholder: "Enter model name, e.g. deepseek-v4-pro",
    contextWindowPlaceholder: "Not set, default 128k",
    modelList: "Model list",
    modelListHint:
      "Each model can have its own context window size. Press Enter or click + to add.",
    addModel: "Add model",
    removeModel: "Remove model",
  },
} as const;

// ============================================================
// 按分类分组模板
// ============================================================

export const DIRECT_TEMPLATES = PROVIDER_TEMPLATES.filter((t) => t.category === "direct");
export const AGGREGATION_TEMPLATES = PROVIDER_TEMPLATES.filter((t) => t.category === "aggregation");
export const OTHER_TEMPLATES = PROVIDER_TEMPLATES.filter((t) => t.category === "other");
