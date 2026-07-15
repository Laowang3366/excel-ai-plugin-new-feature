import { describe, expect, it } from "vitest";
import { buildReasoningAutoHint } from "./providerReasoningHint";

describe("buildReasoningAutoHint", () => {
  it("formats the Chinese auto adaptation hint", () => {
    expect(buildReasoningAutoHint(["off", "low", "high"], "zh-CN")).toBe(
      "已根据当前供应商/API/模型自动适配：关闭 / 低 / 高",
    );
  });

  it("formats the English auto adaptation hint", () => {
    expect(buildReasoningAutoHint(["off", "medium"], "en-US")).toBe(
      "Automatically adapted for this provider/API/model: Off / Medium",
    );
  });
});
