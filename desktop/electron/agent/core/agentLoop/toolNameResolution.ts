import type { ToolExecutor } from "../../shared/types";
import { desanitizeToolName, sanitizeToolName } from "../../providers/openaiCompatibleClient";

export function resolveExecutableToolName(
  name: string,
  executors?: Map<string, ToolExecutor>
): string | null {
  const desanitized = desanitizeToolName(name);
  const candidates = Array.from(new Set([
    name,
    desanitized,
    sanitizeToolName(name),
    sanitizeToolName(desanitized),
    name.replace(/\.(?=[^.]+$)/, "_"),
    desanitized.replace(/\.(?=[^.]+$)/, "_"),
  ]));
  for (const candidate of candidates) {
    if (executors?.has(candidate)) return candidate;
  }
  return null;
}
