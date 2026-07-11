import { describe, expect, it } from "vitest";

import {
  createEmptyFormulaDraft,
  getTaskDraftsForKey,
  moveTaskDraftStore,
  updateTaskDraftStore,
  type TaskDraftStore,
} from "./useTaskDrafts";

describe("task draft store helpers", () => {
  it("clears formula task data while preserving the detected host", () => {
    expect(createEmptyFormulaDraft({
      dataSourceRanges: ["Sheet1!A1:B10"],
      dataSourceInput: "Sheet1!C1:C10",
      referenceSampleRange: "Sheet1!D1:D3",
      referenceSampleMode: "complete",
      outputRange: "Sheet1!F1",
      hostEnvironment: "wps",
      task: "提取规格",
    })).toEqual({
      dataSourceRanges: [],
      dataSourceInput: "",
      referenceSampleRange: "",
      referenceSampleMode: "partial",
      outputRange: "",
      hostEnvironment: "wps",
      task: "",
    });
  });

  it("returns a complete formula draft without changing the store", () => {
    const formulaDraft = {
      dataSourceRanges: ["Sheet1!A1:B10"],
      dataSourceInput: "Sheet1!A1:B10",
      referenceSampleRange: "Sheet1!D1:D3",
      referenceSampleMode: "partial" as const,
      outputRange: "Sheet1!C1:C10",
      hostEnvironment: "microsoft_excel" as const,
      task: "计算同比增长率",
    };
    const store: TaskDraftStore = {
      "thread-1": {
        formula: formulaDraft,
      },
    };
    const originalStore = structuredClone(store);

    const drafts = getTaskDraftsForKey(store, "thread-1");

    expect(drafts.formula).toEqual(formulaDraft);
    expect(store).toEqual(originalStore);
  });

  it("keeps task drafts isolated by draft key", () => {
    let store: TaskDraftStore = {};

    store = updateTaskDraftStore(store, "thread-1", {
      clean: { range: "Sheet1!A1", task: "去重" },
    });
    store = updateTaskDraftStore(store, "thread-2", {
      clean: { range: "Sheet2!B2", task: "排序" },
    });

    expect(getTaskDraftsForKey(store, "thread-1").clean).toEqual({
      range: "Sheet1!A1",
      task: "去重",
    });
    expect(getTaskDraftsForKey(store, "thread-2").clean).toEqual({
      range: "Sheet2!B2",
      task: "排序",
    });
  });

  it("applies functional updates only to the active draft key", () => {
    const store = updateTaskDraftStore(
      {
        "thread-1": { chart: { range: "A1:B2", task: "柱状图" } },
        "thread-2": { chart: { range: "C1:D2", task: "折线图" } },
      },
      "thread-1",
      (prev) => ({
        ...prev,
        chart: {
          range: prev.chart?.range ?? "",
          task: "饼图",
        },
      })
    );

    expect(getTaskDraftsForKey(store, "thread-1").chart).toEqual({
      range: "A1:B2",
      task: "饼图",
    });
    expect(getTaskDraftsForKey(store, "thread-2").chart).toEqual({
      range: "C1:D2",
      task: "折线图",
    });
  });

  it("moves a new-thread draft to its created thread", () => {
    const store: TaskDraftStore = {
      "new:folder-1": {
        clean: { range: "Sheet1!A1:B8", task: "去重" },
        chart: { range: "Sheet1!A1:B8", task: "柱状图" },
      },
    };

    expect(moveTaskDraftStore(store, "new:folder-1", "thread-1")).toEqual({
      "thread-1": store["new:folder-1"],
    });
  });

  it("does not overwrite an existing target-thread draft", () => {
    const store: TaskDraftStore = {
      new: { clean: { range: "A1:B2", task: "去重" } },
      "thread-1": { clean: { range: "C1:D2", task: "排序" } },
    };

    expect(moveTaskDraftStore(store, "new", "thread-1")).toBe(store);
  });
});
