import type {
  RuntimeMemoryKind,
  RuntimeMemoryVisibility,
} from "../stateRuntimeTypes";

export type MemoryConsolidationAction = "add" | "ignore";

export interface ConsolidationCandidate {
  kind: RuntimeMemoryKind;
  visibility: RuntimeMemoryVisibility;
  namespace?: string;
  content: string;
  metadata?: Record<string, unknown>;
}

export function chooseConsolidationAction(
  candidate: ConsolidationCandidate,
): MemoryConsolidationAction {
  if (!candidate.content.trim()) return "ignore";

  if (candidate.kind === "tool_success_profile") {
    return "ignore";
  }

  if (candidate.kind === "correction" && candidate.visibility === "user") {
    return "add";
  }

  return candidate.visibility === "user" ? "add" : "ignore";
}
