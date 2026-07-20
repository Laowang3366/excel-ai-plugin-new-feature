import { describe, expect, it, vi } from "vitest";
import {
  CHAT_READONLY_DENY_ERROR,
  CHAT_READONLY_TOOL_ALLOWLIST,
  GuardedChatExecutor,
  listChatReadOnlyTools,
} from "../shared/agentChat";
import { TOOL_DEFINITIONS, TOOL_DEFINITION_MAP } from "../shared/tools";
import type { ToolCall, ToolName, ToolResult } from "../shared/tools/types";

describe("chat readonly allowlist", () => {
  it("exact set equality; all safe; pageLayout.get in, set out", () => {
    expect([...CHAT_READONLY_TOOL_ALLOWLIST]).toEqual([
      "host.status",
      "selection.get",
      "range.read",
      "range.format.read",
      "formula.read",
      "formula.context",
      "formula.protection.inspect",
      "formula.dependencies.inspect",
      "formula.backups.inspect",
      "sheet.list",
      "table.list",
      "chart.list",
      "chart.series.list",
      "chart.image.get",
      "range.image.get",
      "workbook.inspect",
      "workbook.objects.inspect",
      "conditionalFormat.list",
      "dataValidation.read",
      "sheet.visibility.get",
      "sheet.protection.get",
      "namedRange.list",
      "sheet.display.get",
      "sheet.freeze.get",
      "sheet.pageLayout.get",
      "shape.list",
    ]);

    const listed = listChatReadOnlyTools();
    expect(listed.map((t) => t.name)).toEqual([...CHAT_READONLY_TOOL_ALLOWLIST]);
    for (const t of listed) {
      expect(t.riskLevel).toBe("safe");
      expect(TOOL_DEFINITION_MAP[t.name].riskLevel).toBe("safe");
    }

    expect(listed.some((t) => t.name === "sheet.pageLayout.get")).toBe(true);
    expect(listed.some((t) => t.name === "sheet.pageLayout.set")).toBe(false);

    // Fresh array; does not share identity with registry entries mutation surface.
    const again = listChatReadOnlyTools();
    expect(again).not.toBe(listed);
    again.pop();
    expect(listChatReadOnlyTools()).toHaveLength(CHAT_READONLY_TOOL_ALLOWLIST.length);
  });

  it("exposes no moderate tools", () => {
    const names = new Set(listChatReadOnlyTools().map((t) => t.name));
    for (const def of TOOL_DEFINITIONS) {
      if (def.riskLevel === "moderate") {
        expect(names.has(def.name)).toBe(false);
      }
    }
  });
});

describe("GuardedChatExecutor", () => {
  it("rejects range.write / sheet.delete without calling inner; allows range.read", async () => {
    const innerExecute = vi.fn(async (call: ToolCall): Promise<ToolResult> => ({
      ok: true,
      tool: call.name,
      data: { via: "inner" },
    }));
    const guard = new GuardedChatExecutor({ execute: innerExecute });

    const write = await guard.execute({
      name: "range.write" as ToolName,
      arguments: { sheetName: "S", range: "A1", values: [["x"]] },
    });
    expect(write.ok).toBe(false);
    if (!write.ok) {
      expect(write.error).toContain(CHAT_READONLY_DENY_ERROR);
      expect(write.error).toContain("range.write");
    }
    expect(innerExecute).not.toHaveBeenCalled();

    const del = await guard.execute({
      name: "sheet.delete" as ToolName,
      arguments: { sheetName: "S" },
    });
    expect(del.ok).toBe(false);
    if (!del.ok) expect(del.error).toContain("sheet.delete");
    expect(innerExecute).not.toHaveBeenCalled();

    const read = await guard.execute({
      name: "range.read",
      arguments: { sheetName: "S", range: "A1" },
    });
    expect(read).toEqual({ ok: true, tool: "range.read", data: { via: "inner" } });
    expect(innerExecute).toHaveBeenCalledTimes(1);

    const unknown = await guard.execute({
      name: "not.a.tool" as ToolName,
      arguments: {},
    });
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) expect(unknown.error).toContain(CHAT_READONLY_DENY_ERROR);
    expect(innerExecute).toHaveBeenCalledTimes(1);
  });
});
