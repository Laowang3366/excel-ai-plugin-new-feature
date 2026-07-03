import { describe, expect, it } from "vitest";

import type { FolderFileInfo, ThreadMetadata } from "../electronApi";
import { buildSidebarSearchResults } from "./sidebarSearch";

const threads: ThreadMetadata[] = [
  {
    threadId: "t-old",
    preview: "写一份周报",
    modelProvider: "openai",
    createdAt: 1,
    updatedAt: 10,
  },
  {
    threadId: "t-report",
    preview: "优化销售报告排版",
    modelProvider: "openai",
    createdAt: 2,
    updatedAt: 30,
    folderId: "D:/work/reports",
  },
];

const folderFiles: Record<string, FolderFileInfo[]> = {
  "D:/work/reports": [
    {
      fileName: "销售报告.docx",
      filePath: "D:/work/reports/销售报告.docx",
      size: 1024,
      lastModified: 20,
    },
    {
      fileName: "预算表.xlsx",
      filePath: "D:/work/reports/预算表.xlsx",
      size: 2048,
      lastModified: 40,
    },
  ],
};

const folders = [
  {
    path: "D:/work/reports",
    name: "报告项目",
    addedAt: 1,
  },
];

describe("buildSidebarSearchResults", () => {
  it("finds matching threads, files, and actions from one query", () => {
    const results = buildSidebarSearchResults({
      query: "报告",
      threads,
      folders,
      folderFiles,
      actions: [
        { id: "newThread", label: "新建会话" },
        { id: "settings", label: "设置" },
      ],
    });

    expect(results.threads.map((item) => item.thread.threadId)).toEqual(["t-report"]);
    expect(results.files.map((item) => item.file.fileName)).toEqual(["销售报告.docx"]);
    expect(results.actions).toEqual([]);
  });

  it("returns recent threads, files, and suggested actions for an empty query", () => {
    const results = buildSidebarSearchResults({
      query: " ",
      threads,
      folders,
      folderFiles,
      actions: [
        { id: "newThread", label: "新建会话" },
        { id: "addFolder", label: "添加文件夹" },
      ],
    });

    expect(results.threads.map((item) => item.thread.threadId)).toEqual(["t-report", "t-old"]);
    expect(results.files.map((item) => item.file.fileName)).toEqual(["预算表.xlsx", "销售报告.docx"]);
    expect(results.actions.map((item) => item.id)).toEqual(["newThread", "addFolder"]);
  });
});
