import { normalizeConnectionMode } from "./endpointResolve";
import { MemorySecretStore } from "./memorySecretStore";
import { getProviderTemplate } from "./templates";
import type {
  ConnectionMode,
  CreateProviderInput,
  ProviderConfig,
  ProviderPublicView,
  ProviderSecretStore,
  UpdateProviderInput,
} from "./types";

function createId(): string {
  return `provider_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
}

function toPublic(config: ProviderConfig, apiKey: string): ProviderPublicView {
  return {
    id: config.id,
    name: config.name,
    provider: config.provider,
    baseUrl: config.baseUrl,
    model: config.model,
    apiFormat: config.apiFormat,
    connectionMode: config.connectionMode,
    gatewayBaseUrl: config.gatewayBaseUrl,
    gatewayUpstreamId: config.gatewayUpstreamId,
    contextWindowSize: config.contextWindowSize,
    reasoningMode: config.reasoningMode,
    templateId: config.templateId,
    hasApiKey: apiKey.length > 0,
  };
}

function gatewayFields(input: {
  connectionMode?: ConnectionMode;
  gatewayBaseUrl?: string;
  gatewayUpstreamId?: string;
}): {
  connectionMode: ConnectionMode;
  gatewayBaseUrl: string;
  gatewayUpstreamId: string;
} {
  return {
    connectionMode: normalizeConnectionMode(input.connectionMode),
    gatewayBaseUrl: input.gatewayBaseUrl?.trim() ?? "",
    gatewayUpstreamId: input.gatewayUpstreamId?.trim() ?? "",
  };
}

export class ProviderStore {
  private providers = new Map<string, ProviderConfig>();
  private activeId: string | null = null;

  constructor(private readonly secrets: ProviderSecretStore = new MemorySecretStore()) {}

  list(): ProviderPublicView[] {
    return [...this.providers.values()].map((config) =>
      toPublic(config, this.secrets.get(config.id) ?? ""),
    );
  }

  getActiveId(): string | null {
    return this.activeId;
  }

  getActive(): ProviderConfig | null {
    if (!this.activeId) return null;
    return this.getWithSecret(this.activeId);
  }

  getWithSecret(id: string): ProviderConfig | null {
    const config = this.providers.get(id);
    if (!config) return null;
    return {
      ...config,
      apiKey: this.secrets.get(id) ?? "",
    };
  }

  private viewOf(id: string): ProviderPublicView {
    const config = this.providers.get(id);
    if (!config) throw new Error(`Provider not found: ${id}`);
    return toPublic(config, this.secrets.get(id) ?? "");
  }

  add(input: CreateProviderInput): ProviderPublicView {
    const template = input.templateId ? getProviderTemplate(input.templateId) : undefined;
    const mode = gatewayFields(input);
    const id = createId();
    const config: ProviderConfig = {
      id,
      name: input.name,
      provider: input.provider,
      apiKey: "",
      baseUrl: input.baseUrl,
      model: input.model,
      apiFormat: input.apiFormat,
      connectionMode: mode.connectionMode,
      gatewayBaseUrl: mode.gatewayBaseUrl,
      gatewayUpstreamId: mode.gatewayUpstreamId,
      contextWindowSize:
        input.contextWindowSize ?? template?.defaultContextWindowSize ?? 128_000,
      reasoningMode: input.reasoningMode ?? template?.defaultReasoningMode ?? "off",
      templateId: input.templateId,
    };
    this.providers.set(id, config);
    if (mode.connectionMode === "direct" && input.apiKey.trim()) {
      this.secrets.set(id, input.apiKey.trim());
    } else {
      this.secrets.delete(id);
    }
    if (!this.activeId) this.activeId = id;
    return this.viewOf(id);
  }

  addFromTemplate(
    templateId: string,
    apiKey = "",
    options?: {
      connectionMode?: ConnectionMode;
      gatewayBaseUrl?: string;
      gatewayUpstreamId?: string;
    },
  ): ProviderPublicView {
    const template = getProviderTemplate(templateId);
    if (!template) throw new Error(`Unknown template: ${templateId}`);
    return this.add({
      name: template.name,
      provider: template.provider,
      apiKey,
      baseUrl: template.baseUrl,
      model: template.defaultModel,
      apiFormat: template.apiFormat,
      connectionMode: options?.connectionMode,
      gatewayBaseUrl: options?.gatewayBaseUrl,
      gatewayUpstreamId: options?.gatewayUpstreamId,
      contextWindowSize: template.defaultContextWindowSize,
      reasoningMode: template.defaultReasoningMode,
      templateId: template.id,
    });
  }

  update(id: string, patch: UpdateProviderInput): ProviderPublicView {
    const existing = this.providers.get(id);
    if (!existing) throw new Error(`Provider not found: ${id}`);
    const connectionMode = normalizeConnectionMode(
      patch.connectionMode ?? existing.connectionMode,
    );
    const next: ProviderConfig = {
      ...existing,
      name: patch.name ?? existing.name,
      provider: patch.provider ?? existing.provider,
      baseUrl: patch.baseUrl ?? existing.baseUrl,
      model: patch.model ?? existing.model,
      apiFormat: patch.apiFormat ?? existing.apiFormat,
      connectionMode,
      gatewayBaseUrl: patch.gatewayBaseUrl?.trim() ?? existing.gatewayBaseUrl,
      gatewayUpstreamId:
        patch.gatewayUpstreamId?.trim() ?? existing.gatewayUpstreamId,
      contextWindowSize: patch.contextWindowSize ?? existing.contextWindowSize,
      reasoningMode: patch.reasoningMode ?? existing.reasoningMode,
      apiKey: "",
    };
    this.providers.set(id, next);
    if (connectionMode === "gateway") {
      this.secrets.delete(id);
    } else if (patch.apiKey !== undefined) {
      const apiKey = patch.apiKey.trim();
      if (apiKey) this.secrets.set(id, apiKey);
      else this.secrets.delete(id);
    }
    return this.viewOf(id);
  }

  remove(id: string): void {
    if (!this.providers.has(id)) throw new Error(`Provider not found: ${id}`);
    this.providers.delete(id);
    this.secrets.delete(id);
    if (this.activeId === id) {
      this.activeId = this.providers.keys().next().value ?? null;
    }
  }

  setActive(id: string): void {
    if (!this.providers.has(id)) throw new Error(`Provider not found: ${id}`);
    this.activeId = id;
  }
}
