import { describe, expect, it, vi } from "vitest";

import { generateSummary } from "./summaryGenerator";

describe("generateSummary", () => {
  it("rejects empty model output instead of treating it as a valid summary", async () => {
    const aiClient = {
      chat: vi.fn().mockResolvedValue({ content: "   " }),
    };

    await expect(generateSummary(aiClient, "历史内容")).rejects.toThrow("压缩摘要为空");
  });

  it("propagates provider failures without converting them to summary text", async () => {
    const failure = new Error("upstream failed");
    const aiClient = {
      chat: vi.fn().mockRejectedValue(failure),
    };

    await expect(generateSummary(aiClient, "历史内容")).rejects.toBe(failure);
  });
});
