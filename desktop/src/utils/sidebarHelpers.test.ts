import { describe, expect, it } from "vitest";
import { matchesSidebarSearch } from "./sidebarHelpers";

describe("matchesSidebarSearch", () => {
  it("matches a trimmed query across multiple sidebar text fields", () => {
    expect(matchesSidebarSearch(["清除当前子表所有内容", "练习题"], " 子表 ")).toBe(true);
  });

  it("matches case-insensitively for project and chat text", () => {
    expect(matchesSidebarSearch(["Project Reports", "任务说明"], "reports")).toBe(true);
  });

  it("treats an empty query as a match", () => {
    expect(matchesSidebarSearch(["任何会话"], "   ")).toBe(true);
  });
});
