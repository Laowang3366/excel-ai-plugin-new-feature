import { describe, expect, it } from "vitest";

import { StateRuntimeStore } from "../stateRuntimeStore";
import { runMemoryStartupTask } from "./memoryStartupTask";

describe("memory startup task", () => {
  it("does not throw and reports zero processed events when there are no new events", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();

    try {
      await expect(
        runMemoryStartupTask({ runtime, pipelineId: "test" }),
      ).resolves.toEqual({ processed: 0 });
    } finally {
      await runtime.close();
    }
  });

  it("uses the default pipeline when pipelineId is omitted", async () => {
    const runtime = new StateRuntimeStore(":memory:");
    await runtime.init();

    try {
      await runtime.setMemoryPipelineCursor("default", 7);
      const pipelineIds: string[] = [];
      const getCursor = runtime.getMemoryPipelineCursor.bind(runtime);
      const setCursor = runtime.setMemoryPipelineCursor.bind(runtime);
      runtime.getMemoryPipelineCursor = async (pipelineId) => {
        pipelineIds.push(pipelineId);
        return getCursor(pipelineId);
      };
      runtime.setMemoryPipelineCursor = async (pipelineId, cursor) => {
        pipelineIds.push(pipelineId);
        return setCursor(pipelineId, cursor);
      };

      await expect(runMemoryStartupTask({ runtime })).resolves.toEqual({
        processed: 0,
      });
      expect(pipelineIds).toEqual(["default", "default"]);
      expect(await runtime.getMemoryPipelineCursor("default")).toBe(7);
    } finally {
      await runtime.close();
    }
  });
});
