import { describe, expect, it } from "vitest";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";

const SLICER_TOOLS = [
  "slicer.list",
  "slicer.create",
  "slicer.update",
  "slicer.delete",
  "slicer.filter.get",
  "slicer.filter.apply",
  "slicer.filter.clear",
] as const;

describe("phase54 slicer contract", () => {
  it("registers 7 slicer tools; total 96 unique names", () => {
    const names = TOOL_DEFINITIONS.map((d) => d.name);
    expect(TOOL_DEFINITIONS).toHaveLength(96);
    expect(new Set(names).size).toBe(96);
    for (const name of SLICER_TOOLS) expect(names).toContain(name);
  });

  it("read-only allowlist includes list and filter.get only", () => {
    expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("slicer.list");
    expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("slicer.filter.get");
    expect(CHAT_READONLY_TOOL_ALLOWLIST).not.toContain("slicer.create");
  });

  it("schemas are closed; create requires advancedIntent; empty keys allowed on apply", () => {
    for (const name of SLICER_TOOLS) {
      const def = TOOL_DEFINITIONS.find((d) => d.name === name)!;
      const params = def.parameters as { additionalProperties?: boolean };
      expect(params.additionalProperties).toBe(false);
    }
    const create = TOOL_DEFINITIONS.find((d) => d.name === "slicer.create")!;
    const props = (create.parameters as { properties: Record<string, unknown> }).properties;
    expect((props.advancedIntent as { const?: string }).const).toBe("interactive-pivot");
    const apply = TOOL_DEFINITIONS.find((d) => d.name === "slicer.filter.apply")!;
    const keys = (apply.parameters as { properties: { keys: { maxItems?: number } } }).properties
      .keys;
    expect(keys.maxItems).toBe(500);
  });

  it("executor rejects null/unknown before Host; empty update; empty keys=select all", async () => {
    const host = new MockHostAdapter();
    const ex = new ToolExecutor(host);
    let hostCreate = 0;
    const orig = host.createSlicer.bind(host);
    host.createSlicer = async (input) => {
      hostCreate += 1;
      return orig(input);
    };

    const nullName = await ex.execute({
      name: "slicer.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceType: "table",
        sourceName: null,
        sourceField: "Dept",
        destinationSheet: "Sheet1",
      },
    });
    expect(nullName.ok).toBe(false);
    expect(hostCreate).toBe(0);

    const unknown = await ex.execute({
      name: "slicer.list",
      arguments: { extra: 1 },
    });
    expect(unknown.ok).toBe(false);

    // seed table for create
    await host.createTable({ sheetName: "Sheet1", address: "A1:B2", name: "SalesTable" });
    const created = await ex.execute({
      name: "slicer.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceType: "table",
        sourceName: "SalesTable",
        sourceField: "Dept",
        destinationSheet: "Sheet1",
        name: "DeptSlicer",
      },
    });
    expect(created.ok).toBe(true);
    expect(hostCreate).toBe(1);
    if (created.ok) {
      const data = created.data as { requestedSource: { sourceType: string }; name: string };
      expect(data.name).toBe("DeptSlicer");
      expect(data.requestedSource.sourceType).toBe("table");
    }

    const emptyUpdate = await ex.execute({
      name: "slicer.update",
      arguments: { name: "DeptSlicer" },
    });
    expect(emptyUpdate.ok).toBe(false);

    const applyAll = await ex.execute({
      name: "slicer.filter.apply",
      arguments: { name: "DeptSlicer", keys: [] },
    });
    expect(applyAll.ok).toBe(true);
    if (applyAll.ok) {
      const data = applyAll.data as { isFilterCleared: boolean; selectedKeys: string[] };
      expect(data.isFilterCleared).toBe(true);
      expect(data.selectedKeys.length).toBe(3);
    }

    const applySome = await ex.execute({
      name: "slicer.filter.apply",
      arguments: { name: "DeptSlicer", keys: ["A", "C"] },
    });
    expect(applySome.ok).toBe(true);
    if (applySome.ok) {
      const data = applySome.data as { selectedKeys: string[]; isFilterCleared: boolean };
      expect(data.selectedKeys.sort()).toEqual(["A", "C"]);
      expect(data.isFilterCleared).toBe(false);
    }

    const cleared = await ex.execute({
      name: "slicer.filter.clear",
      arguments: { name: "DeptSlicer" },
    });
    expect(cleared.ok).toBe(true);
    if (cleared.ok) {
      expect((cleared.data as { isFilterCleared: boolean }).isFilterCleared).toBe(true);
    }

    const deleted = await ex.execute({
      name: "slicer.delete",
      arguments: { name: "DeptSlicer" },
    });
    expect(deleted.ok).toBe(true);
    const listed = await ex.execute({ name: "slicer.list", arguments: {} });
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect((listed.data as { slicers: unknown[] }).slicers).toHaveLength(0);
    }
  });
});
