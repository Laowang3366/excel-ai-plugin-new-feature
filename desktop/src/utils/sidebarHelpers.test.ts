import { describe, expect, it } from "vitest";
import { buildSidebarDerivedLists, matchesSidebarSearch, sortSidebarItems } from "./sidebarHelpers";

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

describe("sortSidebarItems", () => {
  const items = [
    { preview: "Beta", updatedAt: 20 },
    { preview: "Alpha", updatedAt: 10 },
    { preview: "Gamma", updatedAt: 30 },
  ];

  it("sorts sidebar items by recent time and display name", () => {
    expect(sortSidebarItems(items, "recentDesc", "en-US").map((item) => item.preview)).toEqual(["Gamma", "Beta", "Alpha"]);
    expect(sortSidebarItems(items, "recentAsc", "en-US").map((item) => item.preview)).toEqual(["Alpha", "Beta", "Gamma"]);
    expect(sortSidebarItems(items, "nameAsc", "en-US").map((item) => item.preview)).toEqual(["Alpha", "Beta", "Gamma"]);
    expect(sortSidebarItems(items, "nameDesc", "en-US").map((item) => item.preview)).toEqual(["Gamma", "Beta", "Alpha"]);
  });
});

describe("buildSidebarDerivedLists", () => {
  it("groups folder threads separately from ungrouped conversations", () => {
    const result = buildSidebarDerivedLists({
      threads: [
        { threadId: "thread-1", preview: "Ungrouped", modelProvider: "test", createdAt: 1, updatedAt: 5 },
        { threadId: "thread-2", preview: "Folder recent", modelProvider: "test", createdAt: 1, updatedAt: 20, folderId: "C:/Work" },
        { threadId: "thread-3", preview: "Folder older", modelProvider: "test", createdAt: 1, updatedAt: 10, folderId: "C:/Work" },
      ],
      pinnedFolders: [
        { path: "C:/Work", name: "Work", addedAt: 100 },
      ],
      folderFiles: {
        "C:/Work": [{ fileName: "plan.docx", filePath: "C:/Work/plan.docx", size: 10, lastModified: 1 }],
      },
      projectSortMode: "recentDesc",
      conversationSortMode: "recentDesc",
      language: "en-US",
    });

    expect(result.ungroupedThreads.map((thread) => thread.threadId)).toEqual(["thread-1"]);
    expect(result.groupedByFolder).toHaveLength(1);
    expect(result.groupedByFolder[0].threads.map((thread) => thread.threadId)).toEqual(["thread-2", "thread-3"]);
    expect(result.groupedByFolder[0].files.map((file) => file.fileName)).toEqual(["plan.docx"]);
    expect(result.hasProjectItems).toBe(true);
    expect(result.hasConversationItems).toBe(true);
    expect(result.showNoSearchResults).toBe(false);
  });
});
