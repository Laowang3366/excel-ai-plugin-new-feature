import { getProviderTemplate } from "./templates";
import type {
  CreateProviderInput,
  ProviderConfig,
  ProviderPublicView,
  ProviderSecretStore,
  UpdateProviderInput,
} from "./types";
import { MemorySecretStore } from "./memorySecretStore";

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
    contextWindowSize: config.contextWindowSize,
    reasoningMode: config.reasoningMode,
    templateId: config.templateId,
    hasApiKey: apiKey.length > 0,
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
    const id = createId();
    const config: ProviderConfig = {
      id,
      name: input.name,
      provider: input.provider,
      apiKey: "",
      baseUrl: input.baseUrl,
      model: input.model,
      apiFormat: input.apiFormat,
      contextWindowSize:
        input.contextWindowSize ?? template?.defaultContextWindowSize ?? 128_000,
      reasoningMode: input.reasoningMode ?? template?.defaultReasoningMode ?? "off",
      templateId: input.templateId,
    };
    this.providers.set(id, config);
    this.secrets.set(id, input.apiKey);
    if (!this.activeId) this.activeId = id;
    return this.viewOf(id);
  }

  addFromTemplate(templateId: string, apiKey = ""): ProviderPublicView {
    const template = getProviderTemplate(templateId);
    if (!template) {
      throw new Error(`Unknown template: ${templateId}`);
    }
    return this.add({
      name: template.name,
      provider: template.provider,
      apiKey,
      baseUrl: template.baseUrl,
      model: template.defaultModel,
      apiFormat: template.apiFormat,
      contextWindowSize: template.defaultContextWindowSize,
      reasoningMode: template.defaultReasoningMode,
      templateId: template.id,
    });
  }

  update(id: string, patch: UpdateProviderInput): ProviderPublicView {
    const existing = this.providers.get(id);
    if (!existing) throw new Error(`Provider not found: ${id}`);
    const next: ProviderConfig = {
      ...existing,
      name: patch.name ?? existing.name,
      provider: patch.provider ?? existing.provider,
      baseUrl: patch.baseUrl ?? existing.baseUrl,
      model: patch.model ?? existing.model,
      apiFormat: patch.apiFormat ?? existing.apiFormat,
      contextWindowSize: patch.contextWindowSize ?? existing.contextWindowSize,
      reasoningMode: patch.reasoningMode ?? existing.reasoningMode,
      apiKey: "",
    };
    this.providers.set(id, next);
    if (patch.apiKey !== undefined) {
      this.secrets.set(id, patch.apiKey);
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
