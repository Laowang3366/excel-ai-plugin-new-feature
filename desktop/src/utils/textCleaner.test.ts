import { describe, expect, it } from "vitest";
import { cleanReasoningText } from "./textCleaner";

describe("cleanReasoningText", () => {
  it("removes token-level spaces around CJK text, cell references, and punctuation", () => {
    expect(cleanReasoningText("我已 经 读取 了 Sheet 2 ! D 2 : F 8")).toBe(
      "我已经读取了Sheet2!D2:F8",
    );
  });

  it("keeps normal English word spacing", () => {
    expect(cleanReasoningText("the quick brown fox")).toBe("the quick brown fox");
  });

  it("preserves markdown table structure while cleaning cell text", () => {
    expect(cleanReasoningText("| 字 段 | value |\n| --- | --- |")).toBe(
      "| 字段 | value |\n| --- | --- |",
    );
  });

  it("normalizes repeated blank lines but keeps paragraph breaks", () => {
    expect(cleanReasoningText("第一 段\n\n\n第二 段")).toBe("第一段\n\n第二段");
  });
});
