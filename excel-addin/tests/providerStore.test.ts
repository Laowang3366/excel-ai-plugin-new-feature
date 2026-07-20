import { describe, expect, it } from "vitest";
import {
  API_FORMATS,
  MemorySecretStore,
  PROVIDER_PERSISTENCE_KEY,
  PROVIDER_PERSISTENCE_VERSION,
  PROVIDER_TEMPLATES,
  ProviderStore,
  type ProviderPersistenceStorage,
} from "../shared/provider";

class TestProviderStorage implements ProviderPersistenceStorage {
  private readonly values = new Map<string, string>();

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }

  raw(): string {
    return this.getItem(PROVIDER_PERSISTENCE_KEY) ?? "";
  }
}

describe("ProviderStore", () => {
  it("includes required vendor templates and api formats", () => {
    const ids = PROVIDER_TEMPLATES.map((template) => template.id);
    for (const id of [
      "openai",
      "deepseek",
      "anthropic",
      "kimi",
      "zhipu",
      "xiaomi",
      "aliyun",
      "tencent",
      "volcengine",
      "xunfei",
      "baidu",
      "jdcloud",
      "qwen",
      "minimax",
      "custom",
    ]) {
      expect(ids).toContain(id);
    }
    expect(API_FORMATS.map((item) => item.value).sort()).toEqual(
      ["anthropic", "openai", "responses"].sort(),
    );
  });

  it("supports CRUD, contextWindow, reasoningMode without localStorage", () => {
    const secrets = new MemorySecretStore();
    const store = new ProviderStore(secrets);
    const created = store.addFromTemplate("openai", "sk-test");
    expect(created.hasApiKey).toBe(true);
    expect(created.contextWindowSize).toBeGreaterThan(0);
    expect(created.reasoningMode).toBeTruthy();
    expect(store.getActiveId()).toBe(created.id);

    const second = store.addFromTemplate("deepseek", "sk-2");
    store.setActive(second.id);
    expect(store.getActiveId()).toBe(second.id);

    store.update(created.id, {
      name: "OpenAI Work",
      apiKey: "sk-new",
      contextWindowSize: 200_000,
      reasoningMode: "high",
      model: "gpt-5.4-mini",
    });
    const updated = store.list().find((item) => item.id === created.id);
    expect(updated?.contextWindowSize).toBe(200_000);
    expect(updated?.reasoningMode).toBe("high");
    expect(updated?.model).toBe("gpt-5.4-mini");
    expect(store.getWithSecret(created.id)?.apiKey).toBe("sk-new");
    expect(JSON.stringify(updated)).not.toContain("sk-new");

    store.remove(second.id);
    expect(store.list().some((item) => item.id === second.id)).toBe(false);
    expect(secrets.get(second.id)).toBeUndefined();
  });

  it("never exposes api key on public list views", () => {
    const store = new ProviderStore();
    store.addFromTemplate("anthropic", "secret-key");
    const listed = store.list()[0];
    expect(listed).toBeDefined();
    expect(JSON.stringify(listed)).not.toContain("secret-key");
    expect(listed.hasApiKey).toBe(true);
  });
});


describe("ProviderStore gateway mode", () => {
  it("creates gateway provider without storing api key", () => {
    const secrets = new MemorySecretStore();
    const store = new ProviderStore(secrets);
    const created = store.addFromTemplate("openai", "should-ignore", {
      connectionMode: "gateway",
      gatewayUpstreamId: "openai",
      gatewayBaseUrl: "https://app.example",
    });
    expect(created.connectionMode).toBe("gateway");
    expect(created.gatewayUpstreamId).toBe("openai");
    expect(created.hasApiKey).toBe(false);
    expect(created.baseUrl).toBe("https://api.openai.com/v1");
    expect(created.gatewayBaseUrl).toBe("https://app.example");
    expect(secrets.get(created.id)).toBeUndefined();
    const withSecret = store.getWithSecret(created.id);
    expect(withSecret?.apiKey).toBe("");
    expect(JSON.stringify(created)).not.toContain("should-ignore");
  });

  it("switching to gateway drops browser api key", () => {
    const secrets = new MemorySecretStore();
    const store = new ProviderStore(secrets);
    const created = store.addFromTemplate("openai", "sk-keep");
    expect(secrets.get(created.id)).toBe("sk-keep");
    store.update(created.id, {
      connectionMode: "gateway",
      gatewayUpstreamId: "openai",
      gatewayBaseUrl: "",
    });
    expect(secrets.get(created.id)).toBeUndefined();
    const view = store.list().find((p) => p.id === created.id);
    expect(view?.connectionMode).toBe("gateway");
    expect(view?.hasApiKey).toBe(false);
    expect(view?.baseUrl).toBe("https://api.openai.com/v1");
  });

  it("keeps direct and gateway URLs separate while editing", () => {
    const store = new ProviderStore();
    const created = store.addFromTemplate("openai", "sk-direct");
    store.update(created.id, {
      connectionMode: "gateway",
      gatewayBaseUrl: "https://plugin.example",
      gatewayUpstreamId: "openai",
    });
    const gateway = store.getWithSecret(created.id);
    expect(gateway?.baseUrl).toBe("https://api.openai.com/v1");
    expect(gateway?.gatewayBaseUrl).toBe("https://plugin.example");
    expect(gateway?.apiKey).toBe("");

    store.update(created.id, { connectionMode: "direct", apiKey: "sk-new" });
    const direct = store.getWithSecret(created.id);
    expect(direct?.baseUrl).toBe("https://api.openai.com/v1");
    expect(direct?.gatewayBaseUrl).toBe("https://plugin.example");
    expect(direct?.apiKey).toBe("sk-new");
  });
});

describe("ProviderStore persistence", () => {
  it("restores non-sensitive provider state and activeId without restoring keys", () => {
    const storage = new TestProviderStorage();
    const directKey = "sk-direct-never-persist";
    const ignoredGatewayKey = "sk-gateway-never-persist";
    const first = new ProviderStore(new MemorySecretStore(), storage);
    const direct = first.addFromTemplate("openai", directKey);
    const gateway = first.addFromTemplate("anthropic", ignoredGatewayKey, {
      connectionMode: "gateway",
      gatewayBaseUrl: "https://addin.example.com",
      gatewayUpstreamId: "anthropic",
    });
    first.setActive(gateway.id);

    const raw = storage.raw();
    expect(raw).not.toContain(directKey);
    expect(raw).not.toContain(ignoredGatewayKey);
    expect(raw).not.toMatch(/apiKey|authorization|secret/i);
    expect(JSON.parse(raw)).toMatchObject({
      version: PROVIDER_PERSISTENCE_VERSION,
      activeId: gateway.id,
    });

    const injected = JSON.parse(raw) as {
      providers: Array<Record<string, unknown>>;
      Authorization?: string;
      secret?: string;
    };
    injected.providers[0].apiKey = directKey;
    injected.Authorization = `Bearer ${directKey}`;
    injected.secret = ignoredGatewayKey;
    storage.setItem(PROVIDER_PERSISTENCE_KEY, JSON.stringify(injected));

    const restored = new ProviderStore(new MemorySecretStore(), storage);
    expect(storage.raw()).not.toContain(directKey);
    expect(storage.raw()).not.toContain(ignoredGatewayKey);
    expect(storage.raw()).not.toMatch(/apiKey|authorization|secret/i);
    expect(restored.list()).toHaveLength(2);
    expect(restored.getActiveId()).toBe(gateway.id);
    expect(restored.getWithSecret(direct.id)?.apiKey).toBe("");
    expect(restored.list().find((item) => item.id === direct.id)?.hasApiKey).toBe(
      false,
    );
    expect(restored.getWithSecret(gateway.id)).toMatchObject({
      connectionMode: "gateway",
      gatewayBaseUrl: "https://addin.example.com",
      gatewayUpstreamId: "anthropic",
      apiKey: "",
    });
  });

  it("persists provider deletion and active-provider fallback", () => {
    const storage = new TestProviderStorage();
    const store = new ProviderStore(new MemorySecretStore(), storage);
    const first = store.addFromTemplate("openai", "sk-one");
    const second = store.addFromTemplate("deepseek", "sk-two");
    store.setActive(second.id);
    store.remove(first.id);

    const afterFirstDelete = new ProviderStore(new MemorySecretStore(), storage);
    expect(afterFirstDelete.list().map((item) => item.id)).toEqual([second.id]);
    expect(afterFirstDelete.getActiveId()).toBe(second.id);

    afterFirstDelete.remove(second.id);
    const afterLastDelete = new ProviderStore(new MemorySecretStore(), storage);
    expect(afterLastDelete.list()).toEqual([]);
    expect(afterLastDelete.getActiveId()).toBeNull();
    expect(JSON.parse(storage.raw())).toMatchObject({
      version: PROVIDER_PERSISTENCE_VERSION,
      activeId: null,
      providers: [],
    });
  });

  it("ignores damaged, unknown-version, and invalid persisted data", () => {
    const storage = new TestProviderStorage();

    storage.setItem(PROVIDER_PERSISTENCE_KEY, "{broken-json");
    expect(new ProviderStore(new MemorySecretStore(), storage).list()).toEqual([]);

    storage.setItem(
      PROVIDER_PERSISTENCE_KEY,
      JSON.stringify({ version: 999, activeId: null, providers: [] }),
    );
    expect(new ProviderStore(new MemorySecretStore(), storage).list()).toEqual([]);

    storage.setItem(
      PROVIDER_PERSISTENCE_KEY,
      JSON.stringify({
        version: PROVIDER_PERSISTENCE_VERSION,
        activeId: "bad",
        providers: [
          {
            id: "bad",
            name: "Bad",
            provider: "bad",
            baseUrl: "https://api.example.com/v1",
            model: "bad-model",
            apiFormat: "unknown",
            connectionMode: "direct",
            gatewayBaseUrl: "",
            gatewayUpstreamId: "",
            contextWindowSize: 128_000,
            reasoningMode: "off",
            apiKey: "must-not-be-loaded",
          },
        ],
      }),
    );
    const invalid = new ProviderStore(new MemorySecretStore(), storage);
    expect(invalid.list()).toEqual([]);
    expect(invalid.getActiveId()).toBeNull();
  });

  it("keeps the in-memory store usable when browser storage throws", () => {
    const unavailable: ProviderPersistenceStorage = {
      getItem() {
        throw new Error("storage disabled");
      },
      setItem() {
        throw new Error("storage full");
      },
    };
    const store = new ProviderStore(new MemorySecretStore(), unavailable);
    const created = store.addFromTemplate("openai", "sk-memory-only");
    expect(store.getActiveId()).toBe(created.id);
    expect(store.getWithSecret(created.id)?.apiKey).toBe("sk-memory-only");
  });
});
