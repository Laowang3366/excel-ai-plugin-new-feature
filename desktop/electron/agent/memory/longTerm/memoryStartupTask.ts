import type { StateRuntimeStore } from "../stateRuntimeStore";

export interface MemoryStartupTaskOptions {
  runtime: StateRuntimeStore;
  pipelineId?: string;
}

export interface MemoryStartupTaskResult {
  processed: number;
}

export async function runMemoryStartupTask(
  options: MemoryStartupTaskOptions,
): Promise<MemoryStartupTaskResult> {
  const pipelineId = options.pipelineId ?? "default";
  const cursor = await options.runtime.getMemoryPipelineCursor(pipelineId);
  await options.runtime.setMemoryPipelineCursor(pipelineId, cursor);
  return { processed: 0 };
}
