import { describe, expect, it } from "vitest";

import { desanitizeToolName, sanitizeToolName } from "./openaiToolNames";

describe("openaiToolNames", () => {
  it("sanitizes dotted internal tool names for OpenAI-compatible APIs", () => {
    expect(sanitizeToolName("office.action.apply")).toBe("office_action_apply");
    expect(sanitizeToolName("range.read")).toBe("range_read");
  });

  it("restores known tool namespaces returned by providers", () => {
    expect(desanitizeToolName("office_action_apply")).toBe("office.action.apply");
    expect(desanitizeToolName("memory_delete")).toBe("memory.delete");
    expect(desanitizeToolName("range_read")).toBe("range.read");
    expect(desanitizeToolName("custom_tool_name")).toBe("custom_tool_name");
  });
});
