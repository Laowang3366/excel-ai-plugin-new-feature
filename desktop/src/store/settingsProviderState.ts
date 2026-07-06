import type { AiProviderConfig } from "../electronApi";
import { normalizeProviderReasoningConfig } from "../utils/reasoningSupport";
import { PROVIDER_TEMPLATES, type ProviderTemplate } from "./settingsProviderTemplates";

interface ProviderStateInput {
  providers: Record<string, AiProviderConfig>;
  activeProviderId: string;
}

export function getProviderTemplate(provider: Pick<AiProviderConfig, "provider">): ProviderTemplate | undefined {
  return PROVIDER_TEMPLATES.find((template) => template.provider === provider.provider);
}

export function normalizeProviderConfig(provider: AiProviderConfig): AiProviderConfig {
  return normalizeProviderReasoningConfig(provider, getProviderTemplate(provider));
}

export function isProviderConfigured(provider?: AiProviderConfig): boolean {
  return !!(provider?.apiKey && provider?.baseUrl && (provider?.model || provider?.defaultModel));
}

export function checkProviderConfigured(
  providers: Record<string, AiProviderConfig>,
  activeProviderId: string,
): boolean {
  if (!activeProviderId) return false;
  return isProviderConfigured(providers[activeProviderId]);
}

export function buildUpdatedProviderState(
  state: ProviderStateInput,
  id: string,
  patch: Partial<AiProviderConfig>,
): ProviderStateInput & { isConfigured: boolean } {
  const nextProvider = normalizeProviderConfig({ ...state.providers[id], ...patch } as AiProviderConfig);
  const providers = {
    ...state.providers,
    [id]: nextProvider,
  };

  return {
    providers,
    activeProviderId: state.activeProviderId,
    isConfigured: checkProviderConfigured(providers, state.activeProviderId),
  };
}

export function buildAddedProviderState(
  state: ProviderStateInput,
  config: AiProviderConfig,
): ProviderStateInput & { isConfigured: boolean } {
  const normalizedConfig = normalizeProviderConfig(config);
  const providers = {
    ...state.providers,
    [normalizedConfig.id]: normalizedConfig,
  };
  const activeProviderId = state.activeProviderId || normalizedConfig.id;

  return {
    providers,
    activeProviderId,
    isConfigured: checkProviderConfigured(providers, activeProviderId),
  };
}

export function buildRemovedProviderState(
  state: ProviderStateInput,
  id: string,
): ProviderStateInput & { isConfigured: boolean } {
  const { [id]: _removed, ...providers } = state.providers;
  const remainingIds = Object.keys(providers);
  const activeProviderId = state.activeProviderId === id
    ? (remainingIds.length > 0 ? remainingIds[0] : "")
    : state.activeProviderId;

  return {
    providers,
    activeProviderId,
    isConfigured: checkProviderConfigured(providers, activeProviderId),
  };
}

export function buildProviderModelsState(
  providers: Record<string, AiProviderConfig>,
  id: string,
  models: string[],
): Record<string, AiProviderConfig> {
  return {
    ...providers,
    [id]: { ...providers[id], models } as AiProviderConfig,
  };
}

export function generateProviderId(): string {
  return `provider_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}
