import { describe, expect, test } from "vitest";
import {
  appendPromptSections,
  composePromptSections,
  renderPromptTemplate,
} from "./promptComposer";

describe("promptComposer", () => {
  test("preserves section order and deduplicates by key", () => {
    const prompt = composePromptSections([
      { key: "base", content: " base " },
      { key: "formula", content: "formula" },
      { key: "formula", content: "duplicate" },
      { key: "runtime", content: "runtime" },
    ]);

    expect(prompt).toBe("base\n\nformula\n\nruntime");
  });

  test("ignores empty sections when appending", () => {
    expect(
      appendPromptSections("base", [
        { key: "empty", content: "  " },
        { key: "runtime", content: "runtime" },
      ]),
    ).toBe("base\n\nruntime");
  });

  test("does not let an empty section reserve a deduplication key", () => {
    expect(
      composePromptSections([
        { key: "runtime", content: "" },
        { key: "runtime", content: "runtime" },
      ]),
    ).toBe("runtime");
  });

  test("renders every declared variable", () => {
    expect(
      renderPromptTemplate("{{FIRST}}/{{SECOND}}", {
        FIRST: "A",
        SECOND: "B",
      }),
    ).toBe("A/B");
  });

  test("rejects templates with missing variables", () => {
    expect(() => renderPromptTemplate("{{MISSING}}", {})).toThrow("缺少提示词模板变量：MISSING");
  });

  test("rejects unsupported placeholder syntax", () => {
    expect(() => renderPromptTemplate("{{lowercase}}", {})).toThrow(
      "存在未替换的提示词模板变量：{{lowercase}}",
    );
  });
});
