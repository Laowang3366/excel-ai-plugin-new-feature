import type { ReasoningMode } from "../electronApi";
import { DEFAULT_CONTEXT_WINDOW } from "../utils/modelContextWindows";

export type ProviderCategory = "direct" | "aggregation" | "other";

export interface ReasoningOption {
  value: ReasoningMode;
  label: string;
}

export interface ProviderTemplate {
  id: string;
  name: string;
  provider: string;
  baseUrl: string;
  defaultModel: string;
  apiFormat: string;
  presetModels?: string[];
  defaultContextWindowSize: number;
  category: ProviderCategory;
  reasoningOptions: ReasoningOption[];
  defaultReasoningMode: ReasoningMode;
}

const REASONING_TIERS_FULL: ReasoningOption[] = [
  { value: "off", label: "关闭" },
  { value: "low", label: "低" },
  { value: "medium", label: "中" },
  { value: "high", label: "高" },
  { value: "max", label: "极高" },
];

const REASONING_TIERS_HIGH_MAX: ReasoningOption[] = [
  { value: "off", label: "关闭" },
  { value: "high", label: "高" },
  { value: "max", label: "极高" },
];

const REASONING_TOGGLE: ReasoningOption[] = [
  { value: "off", label: "关闭" },
  { value: "high", label: "开启" },
];

export const PROVIDER_TEMPLATES: ProviderTemplate[] = [
  {
    id: "openai",
    name: "OpenAI",
    provider: "openai",
    baseUrl: "https://api.openai.com/v1",
    defaultModel: "gpt-5.4",
    apiFormat: "openai",
    presetModels: ["gpt-5.5", "gpt-5.4", "gpt-5.4-mini"],
    defaultContextWindowSize: 256_000,
    category: "direct",
    reasoningOptions: REASONING_TIERS_FULL,
    defaultReasoningMode: "medium",
  },
  {
    id: "deepseek",
    name: "DeepSeek",
    provider: "deepseek",
    baseUrl: "https://api.deepseek.com/v1",
    defaultModel: "deepseek-v4-flash",
    apiFormat: "openai",
    presetModels: ["deepseek-v4-pro", "deepseek-v4-flash"],
    defaultContextWindowSize: 1_000_000,
    category: "direct",
    reasoningOptions: REASONING_TIERS_HIGH_MAX,
    defaultReasoningMode: "high",
  },
  {
    id: "anthropic",
    name: "Anthropic (Claude)",
    provider: "anthropic",
    baseUrl: "https://api.anthropic.com/v1",
    defaultModel: "claude-sonnet-4-6",
    apiFormat: "anthropic",
    presetModels: [
      "claude-fable-5",
      "claude-opus-4-8",
      "claude-opus-4-7",
      "claude-opus-4-6",
      "claude-sonnet-4-6",
    ],
    defaultContextWindowSize: 1_000_000,
    category: "direct",
    reasoningOptions: REASONING_TIERS_FULL,
    defaultReasoningMode: "high",
  },
  {
    id: "kimi",
    name: "Kimi (月之暗面)",
    provider: "kimi",
    baseUrl: "https://api.moonshot.cn/v1",
    defaultModel: "kimi-k2.7-code",
    apiFormat: "openai",
    presetModels: [
      "kimi-k2.7-code",
      "kimi-k2.7-code-highspeed",
      "kimi-k2.6",
      "kimi-k2.5",
    ],
    defaultContextWindowSize: 262_144,
    category: "direct",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "zhipu",
    name: "智谱",
    provider: "zhipu",
    baseUrl: "https://open.bigmodel.cn/api/paas/v4",
    defaultModel: "glm-5.2",
    apiFormat: "openai",
    presetModels: ["glm-5.2", "glm-5.1", "glm-5", "glm-4-long"],
    defaultContextWindowSize: 1_048_576,
    category: "direct",
    reasoningOptions: REASONING_TIERS_HIGH_MAX,
    defaultReasoningMode: "max",
  },
  {
    id: "xiaomi",
    name: "小米",
    provider: "xiaomi",
    baseUrl: "https://token-plan-cn.xiaomimimo.com/v1",
    defaultModel: "mimo-v2.5-pro",
    apiFormat: "openai",
    presetModels: ["mimo-v2.5-pro", "mimo-v2.5-pro-ultraspeed", "mimo-v2.5"],
    defaultContextWindowSize: 1_000_000,
    category: "direct",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "aliyun",
    name: "阿里云百炼",
    provider: "aliyun",
    baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1",
    defaultModel: "",
    apiFormat: "openai",
    defaultContextWindowSize: 1_000_000,
    category: "aggregation",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "tencent",
    name: "腾讯云",
    provider: "tencent",
    baseUrl: "https://api.lkeap.cloud.tencent.com/plan/v3",
    defaultModel: "",
    apiFormat: "openai",
    defaultContextWindowSize: 1_000_000,
    category: "aggregation",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "volcengine",
    name: "火山方舟",
    provider: "volcengine",
    baseUrl: "https://ark.cn-beijing.volces.com/api/plan/v3",
    defaultModel: "",
    apiFormat: "openai",
    defaultContextWindowSize: 262_144,
    category: "aggregation",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "xunfei",
    name: "讯飞星辰",
    provider: "xunfei",
    baseUrl: "https://maas-coding-api.cn-huabei-1.xf-yun.com/v2",
    defaultModel: "",
    apiFormat: "openai",
    defaultContextWindowSize: 131_072,
    category: "aggregation",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "baidu",
    name: "百度千帆",
    provider: "baidu",
    baseUrl: "https://qianfan.baidubce.com/v2",
    defaultModel: "",
    apiFormat: "openai",
    defaultContextWindowSize: 131_072,
    category: "aggregation",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "jdcloud",
    name: "京东云",
    provider: "jdcloud",
    baseUrl: "https://modelservice.jdcloud.com/coding/openai/v1",
    defaultModel: "",
    apiFormat: "openai",
    defaultContextWindowSize: 1_000_000,
    category: "aggregation",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "high",
  },
  {
    id: "custom",
    name: "自定义兼容接口",
    provider: "custom",
    baseUrl: "",
    defaultModel: "",
    apiFormat: "openai",
    defaultContextWindowSize: DEFAULT_CONTEXT_WINDOW,
    category: "other",
    reasoningOptions: REASONING_TOGGLE,
    defaultReasoningMode: "off",
  },
  {
    id: "qwen",
    name: "千问 (Qwen)",
    provider: "qwen",
    baseUrl: "https://dashscope.aliyuncs.com/compatible-mode/v1",
    defaultModel: "qwen3-235b-a22b",
    apiFormat: "openai",
    presetModels: ["qwen3-235b-a22b", "qwen3-30b-a3b", "qwen-max", "qwen-plus", "qwen-turbo"],
    defaultContextWindowSize: 1_000_000,
    category: "direct",
    reasoningOptions: REASONING_TIERS_HIGH_MAX,
    defaultReasoningMode: "high",
  },
  {
    id: "minimax",
    name: "MiniMax",
    provider: "minimax",
    baseUrl: "https://api.minimax.chat/v1",
    defaultModel: "MiniMax-Text-01",
    apiFormat: "openai",
    presetModels: ["MiniMax-Text-01"],
    defaultContextWindowSize: 1_000_000,
    category: "direct",
    reasoningOptions: REASONING_TIERS_HIGH_MAX,
    defaultReasoningMode: "high",
  },
];

export const API_FORMATS = [
  { value: "openai", label: "Chat Completions (/chat/completions)" },
  { value: "anthropic", label: "Anthropic Messages (/v1/messages)" },
  { value: "responses", label: "OpenAI Responses (/responses)" },
] as const;
