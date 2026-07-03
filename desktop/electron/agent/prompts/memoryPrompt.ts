import stageOneSystemPrompt from "./templates/memory/stage_one_system.zh-CN.md?raw";
import consolidationPrompt from "./templates/memory/consolidation.zh-CN.md?raw";
import instructionsPrompt from "./templates/memory/instructions.zh-CN.md?raw";

export type MemoryPromptTemplateName =
  | "stage_one_system"
  | "consolidation"
  | "instructions";

const memoryPromptTemplates: Record<MemoryPromptTemplateName, string> = {
  stage_one_system: stageOneSystemPrompt,
  consolidation: consolidationPrompt,
  instructions: instructionsPrompt,
};

export function loadMemoryPromptTemplate(
  name: MemoryPromptTemplateName,
): string {
  return memoryPromptTemplates[name].trim();
}
