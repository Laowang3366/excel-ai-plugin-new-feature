import { describe, expect, it } from "vitest";

import { loadMemoryPromptTemplate } from "./memoryPrompt";

describe("memoryPrompt", () => {
  it.each([
    [
      "stage_one_system",
      ["用户可见记忆", "tool_success_profile", "不要进入普通对话提示词"],
    ],
    [
      "consolidation",
      ["长期记忆候选清洗合并", "同一用户纠正多次出现", "tool_success_profile"],
    ],
    [
      "instructions",
      ["长期记忆注入说明", "普通对话上下文", "内部工具策略"],
    ],
  ] as const)("loads the %s prompt template", (name, expectedContents) => {
    const prompt = loadMemoryPromptTemplate(name);

    for (const expectedContent of expectedContents) {
      expect(prompt).toContain(expectedContent);
    }
  });
});
