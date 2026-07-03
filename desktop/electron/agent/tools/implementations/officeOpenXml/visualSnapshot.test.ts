import { describe, expect, it } from "vitest";
import { selectSnapshotPlan } from "./visualSnapshot";

describe("visualSnapshot", () => {
  it("selects Open XML renderer before COM fallback", () => {
    expect(selectSnapshotPlan({ preferEngine: "openxml", hasHeadlessRenderer: true, hasComFallback: true }).engine)
      .toBe("openxml");
    expect(selectSnapshotPlan({ preferEngine: "openxml", hasHeadlessRenderer: false, hasComFallback: true }).engine)
      .toBe("com");
    expect(selectSnapshotPlan({ preferEngine: "com", hasHeadlessRenderer: true, hasComFallback: true }).engine)
      .toBe("com");
  });

  it("throws a clear error when no renderer is available", () => {
    expect(() => selectSnapshotPlan({
      preferEngine: "openxml",
      hasHeadlessRenderer: false,
      hasComFallback: false,
    })).toThrow("没有可用的 Office 视觉快照渲染器");
  });
});
