import type { ToolExecutor } from "../../shared/types";
import { normalizeFormulaPreparation } from "../../core/agentLoop/formulaTaskContract";

export function addFormulaWorkflowExecutors(target: Map<string, ToolExecutor>): void {
  target.set("formula.prepare", {
    name: "formula.prepare",
    execute: async (args: Record<string, unknown>) => {
      const preparation = normalizeFormulaPreparation(args);
      if (typeof preparation === "string") {
        return { success: false, error: preparation };
      }
      return { success: true, data: preparation };
    },
  });

  target.set("formula.verify", {
    name: "formula.verify",
    execute: async () => ({
      success: false,
      error: "formula.verify 只能在公式写入后由运行时执行。",
    }),
  });
}
