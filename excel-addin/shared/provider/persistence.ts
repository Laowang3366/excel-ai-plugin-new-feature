import type {
  ApiFormat,
  ConnectionMode,
  ProviderConfig,
  ReasoningMode,
} from "./types";

export const PROVIDER_PERSISTENCE_KEY = "wengge.excel-addin.provider-state";
export const PROVIDER_PERSISTENCE_VERSION = 1;

export interface ProviderPersistenceStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

export interface PersistedProviderState {
  providers: ProviderConfig[];
  activeId: string | null;
}

interface LoadedPersistedProviderState extends PersistedProviderState {
  shouldRewrite: boolean;
}

const API_FORMATS = new Set<ApiFormat>(["openai", "anthropic", "responses"]);
const CONNECTION_MODES = new Set<ConnectionMode>(["direct", "gateway"]);
const REASONING_MODES = new Set<ReasoningMode>([
  "off",
  "low",
  "medium",
  "high",
  "max",
]);

function isRecord(value: unknown): value is Record<string, unknown> {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function readString(
  value: Record<string, unknown>,
  key: string,
  maxLength: number,
  allowEmpty = true,
): string | null {
  const field = value[key];
  if (typeof field !== "string" || field.length > maxLength) return null;
  if (!allowEmpty && field.trim() === "") return null;
  return field;
}

function parseProvider(value: unknown): ProviderConfig | null {
  if (!isRecord(value)) return null;

  const id = readString(value, "id", 128, false);
  const name = readString(value, "name", 200, false);
  const provider = readString(value, "provider", 100, false);
  const baseUrl = readString(value, "baseUrl", 2_048);
  const model = readString(value, "model", 200);
  const gatewayBaseUrl = readString(value, "gatewayBaseUrl", 2_048);
  const gatewayUpstreamId = readString(value, "gatewayUpstreamId", 64);
  const apiFormat = value.apiFormat;
  const connectionMode = value.connectionMode;
  const contextWindowSize = value.contextWindowSize;
  const reasoningMode = value.reasoningMode;
  const templateId = value.templateId;

  if (
    id == null ||
    name == null ||
    provider == null ||
    baseUrl == null ||
    model == null ||
    gatewayBaseUrl == null ||
    gatewayUpstreamId == null ||
    typeof apiFormat !== "string" ||
    !API_FORMATS.has(apiFormat as ApiFormat) ||
    typeof connectionMode !== "string" ||
    !CONNECTION_MODES.has(connectionMode as ConnectionMode) ||
    typeof contextWindowSize !== "number" ||
    !Number.isInteger(contextWindowSize) ||
    contextWindowSize < 1_024 ||
    contextWindowSize > 10_000_000 ||
    typeof reasoningMode !== "string" ||
    !REASONING_MODES.has(reasoningMode as ReasoningMode) ||
    (templateId !== undefined &&
      (typeof templateId !== "string" || templateId.length > 100))
  ) {
    return null;
  }

  return {
    id,
    name,
    provider,
    apiKey: "",
    baseUrl,
    model,
    apiFormat: apiFormat as ApiFormat,
    connectionMode: connectionMode as ConnectionMode,
    gatewayBaseUrl,
    gatewayUpstreamId,
    contextWindowSize,
    reasoningMode: reasoningMode as ReasoningMode,
    ...(typeof templateId === "string" ? { templateId } : {}),
  };
}

export function loadPersistedProviderState(
  storage?: ProviderPersistenceStorage,
): LoadedPersistedProviderState {
  const empty = { providers: [], activeId: null, shouldRewrite: false };
  if (!storage) return empty;

  try {
    const raw = storage.getItem(PROVIDER_PERSISTENCE_KEY);
    if (!raw) return empty;
    const parsed: unknown = JSON.parse(raw);
    if (!isRecord(parsed) || parsed.version !== PROVIDER_PERSISTENCE_VERSION) {
      return empty;
    }
    if (!Array.isArray(parsed.providers) || parsed.providers.length > 100) {
      return empty;
    }

    const providers: ProviderConfig[] = [];
    const ids = new Set<string>();
    for (const candidate of parsed.providers) {
      const provider = parseProvider(candidate);
      if (!provider || ids.has(provider.id)) continue;
      ids.add(provider.id);
      providers.push(provider);
    }

    const requestedActiveId =
      typeof parsed.activeId === "string" && parsed.activeId.length <= 128
        ? parsed.activeId
        : null;
    const activeId =
      requestedActiveId && ids.has(requestedActiveId)
        ? requestedActiveId
        : (providers[0]?.id ?? null);
    return { providers, activeId, shouldRewrite: true };
  } catch {
    return empty;
  }
}

function toPersistedProvider(config: ProviderConfig): Record<string, unknown> {
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
    ...(config.templateId ? { templateId: config.templateId } : {}),
  };
}

export function persistProviderState(
  storage: ProviderPersistenceStorage | undefined,
  state: PersistedProviderState,
): void {
  if (!storage) return;
  try {
    storage.setItem(
      PROVIDER_PERSISTENCE_KEY,
      JSON.stringify({
        version: PROVIDER_PERSISTENCE_VERSION,
        activeId: state.activeId,
        providers: state.providers.map(toPersistedProvider),
      }),
    );
  } catch {
    // Storage access may be disabled or full; keep the in-memory store usable.
  }
}

export function getBrowserProviderPersistenceStorage():
  | ProviderPersistenceStorage
  | undefined {
  try {
    if (typeof window === "undefined") return undefined;
    return window.localStorage;
  } catch {
    return undefined;
  }
}
