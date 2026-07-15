import { describe, expect, it, vi } from "vitest";

import {
  createCompactionProvider,
  createLocalCompactionProvider,
  createRemoteCompactionProvider,
} from "./compactionProvider";
import { DEFAULT_COMPACT_PROMPT } from "../../prompts/compactionPrompt";

describe("LocalCompactionProvider", () => {
  it("uses the template prompt for local summary generation", async () => {
    const aiClient = {
      chat: vi.fn(async () => ({ content: "摘要" })),
    };
    const provider = createLocalCompactionProvider(aiClient);

    await provider.generateSummary({ historyPrompt: "历史内容" });

    expect(aiClient.chat).toHaveBeenCalledWith(
      expect.objectContaining({
        messages: [
          { role: "user", content: DEFAULT_COMPACT_PROMPT },
          { role: "user", content: "历史内容" },
        ],
      }),
    );
  });

  it("keeps compactPrompt overrides compatible with existing settings", async () => {
    const aiClient = {
      chat: vi.fn(async () => ({ content: "摘要" })),
    };
    const provider = createLocalCompactionProvider(aiClient);

    await provider.generateSummary({
      historyPrompt: "历史内容",
      config: { compactPrompt: "覆盖提示词" } as any,
    });

    expect(aiClient.chat).toHaveBeenCalledWith(
      expect.objectContaining({
        messages: [
          { role: "user", content: "覆盖提示词" },
          { role: "user", content: "历史内容" },
        ],
      }),
    );
  });
});

describe("RemoteCompactionProvider", () => {
  it("posts compaction input to the remote endpoint and returns its summary", async () => {
    const originalFetch = globalThis.fetch;
    globalThis.fetch = vi.fn(
      async () =>
        new Response(
          JSON.stringify({
            summary: "远程摘要",
          }),
          { status: 200, headers: { "Content-Type": "application/json" } },
        ),
    ) as any;

    try {
      const provider = createRemoteCompactionProvider({
        endpoint: "https://compact.example.test/v2",
        apiKey: "remote-key",
        model: "compact-model",
      });

      const summary = await provider.generateSummary({ historyPrompt: "历史内容" });

      expect(summary).toBe("远程摘要");
      expect(globalThis.fetch).toHaveBeenCalledWith(
        "https://compact.example.test/v2",
        expect.objectContaining({
          method: "POST",
          headers: expect.objectContaining({
            "Content-Type": "application/json",
            Authorization: "Bearer remote-key",
          }),
          body: JSON.stringify({
            instruction: DEFAULT_COMPACT_PROMPT,
            input: "历史内容",
            model: "compact-model",
          }),
        }),
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("reports remote compaction failures with status and compact body", async () => {
    const originalFetch = globalThis.fetch;
    globalThis.fetch = vi.fn(
      async () => new Response("<html>not found</html>", { status: 404 }),
    ) as any;

    try {
      const provider = createRemoteCompactionProvider({
        endpoint: "https://compact.example.test/v2",
      });

      await expect(provider.generateSummary({ historyPrompt: "历史内容" })).rejects.toThrow(
        "远程压缩失败 (404): not found",
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("rejects successful remote responses that do not contain a summary", async () => {
    const originalFetch = globalThis.fetch;
    globalThis.fetch = vi.fn(
      async () =>
        new Response(
          JSON.stringify({
            summary: "   ",
          }),
          { status: 200, headers: { "Content-Type": "application/json" } },
        ),
    ) as any;

    try {
      const provider = createRemoteCompactionProvider({
        endpoint: "https://compact.example.test/v2",
      });

      await expect(provider.generateSummary({ historyPrompt: "历史内容" })).rejects.toThrow(
        "远程压缩返回空摘要",
      );
    } finally {
      globalThis.fetch = originalFetch;
    }
  });

  it("selects the remote provider only when remote settings are complete", async () => {
    const aiClient = {
      chat: vi.fn(async () => ({ content: "本地摘要" })),
    };

    const localProvider = createCompactionProvider(aiClient, {
      compactionProvider: "remote",
    } as any);

    await localProvider.generateSummary({ historyPrompt: "历史内容" });

    expect(aiClient.chat).toHaveBeenCalled();
  });
});
