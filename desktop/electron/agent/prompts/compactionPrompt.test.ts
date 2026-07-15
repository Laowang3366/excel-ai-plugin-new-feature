import { describe, expect, it } from "vitest";

import { DEFAULT_COMPACT_PROMPT, getCompactionPromptTemplate } from "./compactionPrompt";

describe("compactionPrompt", () => {
  it("loads the default compaction prompt from the prompt template", () => {
    const prompt = getCompactionPromptTemplate();

    expect(prompt).toContain("用户的核心需求和目标");
    expect(prompt).toContain("无缝继续对话");
    expect(prompt).toBe(DEFAULT_COMPACT_PROMPT);
  });

  it("uses an explicit compact prompt override when provided", () => {
    expect(getCompactionPromptTemplate("自定义压缩提示词")).toBe("自定义压缩提示词");
  });
});
