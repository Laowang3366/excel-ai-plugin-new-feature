import type { AIClientConfig } from "../agent/providers/aiClient";
import { DEFAULT_CONTEXT_WINDOW } from "../agent/providers/modelContextWindows";
import { decryptProviderForRuntime, type SettingsSecretCipher } from "./settingsSecrets";

export function buildActiveAIConfig(
  activeProviderId: string,
  providers: Record<string, Record<string, unknown>>,
  cipher: SettingsSecretCipher,
): AIClientConfig {
  if (!activeProviderId || !providers[activeProviderId]) {
    return {
      provider: "openai",
      apiKey: "",
      baseUrl: "https://api.openai.com/v1",
      model: "gpt-4o",
    };
  }

  const provider = decryptProviderForRuntime(providers[activeProviderId], cipher) as any;
  const activeModelConfig = provider.modelConfigs?.find(
    (model: any) => model.name === provider.model,
  );
  return {
    provider: provider.provider,
    apiKey: provider.apiKey || "",
    baseUrl: provider.baseUrl || provider.defaultBaseUrl || "",
    model: provider.model || provider.defaultModel || "",
    apiFormat: provider.apiFormat,
    customHeaders: provider.customHeaders,
    contextWindowSize:
      activeModelConfig?.contextWindowSize || provider.contextWindowSize || DEFAULT_CONTEXT_WINDOW,
    compHash: activeModelConfig?.compHash || provider.compHash,
    reasoningMode: activeModelConfig?.reasoningMode || provider.reasoningMode,
  };
}
