import { describe, expect, it } from "vitest";
import { doneResult, needsComResult, unsupportedResult } from "./results";

describe("office action results", () => {
  it("builds consistent status results", () => {
    expect(doneResult({
      engine: "openxml",
      app: "word",
      action: "style",
      operation: "styleTables",
      summary: "已美化 Word 表格",
    }).status).toBe("done");

    expect(needsComResult({
      app: "word",
      action: "insert",
      operation: "insertOrUpdateToc",
      summary: "目录字段需要 Word 刷新",
    }).engine).toBe("openxml");

    expect(unsupportedResult({
      app: "presentation",
      action: "insert",
      operation: "editAnimationTimeline",
      summary: "首阶段不支持动画时间轴",
    }).status).toBe("unsupported");
  });
});
