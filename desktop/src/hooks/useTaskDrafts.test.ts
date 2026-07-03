import { describe, expect, it } from "vitest";

import {
  getTaskDraftsForKey,
  updateTaskDraftStore,
  type TaskDraftStore,
} from "./useTaskDrafts";

describe("task draft store helpers", () => {
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
});
