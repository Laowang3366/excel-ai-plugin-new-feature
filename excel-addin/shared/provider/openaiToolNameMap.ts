import type { ToolDefinition } from "../tools/types";

export interface ToolNameMaps {
  internalToExternal: Map<string, string>;
  externalToInternal: Map<string, string>;
}

/**
 * Map Excel tool names (may contain ".") to OpenAI function names (no dots, <=64).
 * Built only from this request's tools; no static table or reverse guessing.
 */
export function buildToolNameMaps(tools: ToolDefinition[]): ToolNameMaps | { error: string } {
  const internalToExternal = new Map<string, string>();
  const externalToInternal = new Map<string, string>();
  for (const tool of tools) {
    const internal = tool.name;
    const external = internal.replace(/\./g, "_");
    if (!external || external.length > 64) {
      return { error: `tool name mapping invalid or too long: ${internal}` };
    }
    if (!/^[A-Za-z0-9_-]+$/.test(external)) {
      return { error: `tool name has illegal OpenAI function characters: ${internal}` };
    }
    const prevExt = internalToExternal.get(internal);
    if (prevExt != null && prevExt !== external) {
      return { error: `tool name mapping conflict for ${internal}` };
    }
    const prevInt = externalToInternal.get(external);
    if (prevInt != null && prevInt !== internal) {
      return {
        error: `tool name mapping collision: ${internal} and ${prevInt} both map to ${external}`,
      };
    }
    internalToExternal.set(internal, external);
    externalToInternal.set(external, internal);
  }
  return { internalToExternal, externalToInternal };
}

export function isToolNameMaps(
  value: ToolNameMaps | { error: string },
): value is ToolNameMaps {
  return "internalToExternal" in value;
}
