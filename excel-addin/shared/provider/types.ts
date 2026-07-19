export type ReasoningMode = "off" | "low" | "medium" | "high" | "max";
export type ProviderCategory = "direct" | "aggregation" | "other";
export type ApiFormat = "openai" | "anthropic" | "responses";

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
  apiFormat: ApiFormat;
  presetModels?: string[];
  defaultContextWindowSize: number;
  category: ProviderCategory;
  reasoningOptions: ReasoningOption[];
  defaultReasoningMode: ReasoningMode;
}

export interface ProviderConfig {
  id: string;
  name: string;
  provider: string;
  apiKey: string;
  baseUrl: string;
  model: string;
  apiFormat: ApiFormat;
  contextWindowSize: number;
  reasoningMode: ReasoningMode;
  templateId?: string;
}

export interface ProviderPublicView {
  id: string;
  name: string;
  provider: string;
  baseUrl: string;
  model: string;
  apiFormat: ApiFormat;
  contextWindowSize: number;
  reasoningMode: ReasoningMode;
  templateId?: string;
  hasApiKey: boolean;
}

export interface CreateProviderInput {
  name: string;
  provider: string;
  apiKey: string;
  baseUrl: string;
  model: string;
  apiFormat: ApiFormat;
  contextWindowSize?: number;
  reasoningMode?: ReasoningMode;
  templateId?: string;
}

export interface UpdateProviderInput {
  name?: string;
  provider?: string;
  apiKey?: string;
  baseUrl?: string;
  model?: string;
  apiFormat?: ApiFormat;
  contextWindowSize?: number;
  reasoningMode?: ReasoningMode;
}

/** Narrow storage boundary: never use browser localStorage for apiKey. */
export interface ProviderSecretStore {
  get(id: string): string | undefined;
  set(id: string, apiKey: string): void;
  delete(id: string): void;
  clear(): void;
}
