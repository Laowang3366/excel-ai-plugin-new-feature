import { describe, expect, it } from "vitest";
import {
  API_FORMATS,
  MemorySecretStore,
  PROVIDER_TEMPLATES,
  ProviderStore,
} from "../shared/provider";

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
