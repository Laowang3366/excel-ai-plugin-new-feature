import { describe, expect, it, vi } from "vitest";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import {
  CORE_TOOL_ARGUMENT_ALLOWLIST,
  RANGE_FORMAT_FIELD_KEYS,
} from "../shared/tools/argValidation";
import type { HostAdapter } from "../shared/host/types";
import { ok } from "../shared/host/types";
import { MockHostAdapter } from "./mockHost";

const CORE_TOOLS = Object.keys(CORE_TOOL_ARGUMENT_ALLOWLIST) as (keyof typeof CORE_TOOL_ARGUMENT_ALLOWLIST)[];

const CHART_LIST_DELETE = ["chart.list", "chart.delete"] as const;

const ALL_27 = [...CORE_TOOLS, ...CHART_LIST_DELETE] as const;

function schemaPropertyKeys(toolName: string): string[] {
  const def = TOOL_DEFINITIONS.find((d) => d.name === toolName);
  expect(def, toolName).toBeTruthy();
  const props = (def!.parameters as { properties?: Record<string, unknown> }).properties ?? {};
  return Object.keys(props).sort();
}

function minimalArgs(toolName: string): Record<string, unknown> {
  switch (toolName) {
    case "host.status":
    case "selection.get":
    case "sheet.list":
    case "workbook.inspect":
    case "workbook.objects.inspect":
    case "workbook.save":
      return {};
    case "range.read":
    case "range.clear":
    case "range.format.read":
    case "formula.read":
    case "conditionalFormat.list":
    case "dataValidation.read":
    case "dataValidation.clear":
      return { sheetName: "Sheet1", range: "A1" };
    case "range.write":
      return { sheetName: "Sheet1", range: "A1", values: [["x"]] };
    case "range.format.write":
      return {
        sheetName: "Sheet1",
        range: "A1",
        format: { fontBold: true },
      };
    case "formula.write":
      return { sheetName: "Sheet1", range: "A1", formula: "=1" };
    case "formula.context":
      return { sheetName: "Sheet1" };
    case "sheet.operation":
      return { operation: "add", sheetName: "S2" };
    case "sheet.add":
    case "sheet.delete":
      return { sheetName: "S2" };
    case "sheet.rename":
      return { sheetName: "Sheet1", newName: "Renamed" };
    case "table.list":
    case "chart.list":
      return {};
    case "table.create":
      return { sheetName: "Sheet1", range: "A1:B2" };
    case "table.delete":
      return { sheetName: "Sheet1", tableName: "T1" };
    case "conditionalFormat.add":
      return {
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE" },
      };
    case "conditionalFormat.delete":
      return { sheetName: "Sheet1", range: "A1", id: "cf1" };
    case "dataValidation.write":
      return {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
      };
    case "chart.delete":
      return { sheetName: "Sheet1", chartName: "C1" };
    default:
      throw new Error(`no minimal args for ${toolName}`);
  }
}

describe("phase36 unknown legacy tool arguments", () => {
  it("core allowlist keys match TOOL_DEFINITIONS property keys exactly", () => {
    expect(CORE_TOOLS).toHaveLength(27);
    for (const name of CORE_TOOLS) {
      const allowed = [...CORE_TOOL_ARGUMENT_ALLOWLIST[name]].sort();
      expect(allowed, name).toEqual(schemaPropertyKeys(name));
    }
  });

  it("chart.list/delete schema keys match runtime allowlists", () => {
    expect(schemaPropertyKeys("chart.list")).toEqual(["sheetName"]);
    expect(schemaPropertyKeys("chart.delete")).toEqual(["chartName", "sheetName"].sort());
  });

  it("rejects __unknown on all 29 paths before missing-arg or host work", async () => {
    const host = new MockHostAdapter();
    const spyStatus = vi.spyOn(host, "getStatus");
    const spyRead = vi.spyOn(host, "readRange");
    const spyListCharts = vi.spyOn(host, "listCharts");
    const executor = new ToolExecutor(host);

    for (const name of ALL_27) {
      const result = await executor.execute({
        name,
        arguments: { ...minimalArgs(name), __unknown: 1 },
      });
      expect(result.ok, name).toBe(false);
      if (!result.ok) {
        expect(result.error, name).toMatch(/unknown field: __unknown/);
      }
    }

    expect(spyStatus).not.toHaveBeenCalled();
    expect(spyRead).not.toHaveBeenCalled();
    expect(spyListCharts).not.toHaveBeenCalled();
  });

  it("range.format.write rejects unknown nested format fields", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());
    const result = await executor.execute({
      name: "range.format.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        format: { fontBold: true, __unknown: true },
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error).toMatch(/unknown format field: __unknown/);
    }
  });

  it("range.format.write accepts all nine closed format fields", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const format: Record<string, unknown> = {};
    for (const key of RANGE_FORMAT_FIELD_KEYS) {
      if (key === "fontSize") format[key] = 12;
      else if (key === "fontBold" || key === "wrapText") format[key] = true;
      else format[key] = "x";
    }
    const result = await executor.execute({
      name: "range.format.write",
      arguments: { sheetName: "Sheet1", range: "A1", format },
    });
    expect(result.ok).toBe(true);
  });

  it("allowed optional fields are not treated as unknown", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);

    const cases: Array<{ name: string; args: Record<string, unknown> }> = [
      { name: "range.read", args: { sheetName: "Sheet1", range: "A1", expand: "none" } },
      {
        name: "range.write",
        args: { sheetName: "Sheet1", range: "A1", values: [["1"]], verify: false },
      },
      {
        name: "formula.write",
        args: { sheetName: "Sheet1", range: "A1", formula: "=1+1", verify: false },
      },
      { name: "formula.context", args: { sheetName: "Sheet1", range: "A1:B2" } },
      {
        name: "sheet.operation",
        args: { operation: "move", sheetName: "Sheet1", newName: "X", position: 1 },
      },
      { name: "table.list", args: { sheetName: "Sheet1" } },
      {
        name: "table.create",
        args: { sheetName: "Sheet1", range: "A1:B2", name: "T1", hasHeaders: true },
      },
      { name: "chart.list", args: { sheetName: "Sheet1" } },
    ];

    for (const item of cases) {
      const result = await executor.execute({
        name: item.name as never,
        arguments: item.args,
      });
      // may fail for host reasons (e.g. missing table) but never unknown field
      if (!result.ok) {
        expect(result.error, item.name).not.toMatch(/unknown field/);
      }
    }
  });

  it("CF/DV nested rule unknown still uses unknown rule field", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());
    const cf = await executor.execute({
      name: "conditionalFormat.add",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { kind: "custom", formula: "=TRUE", __unknown: 1 },
      },
    });
    expect(cf.ok).toBe(false);
    if (!cf.ok) expect(cf.error).toMatch(/unknown rule field: __unknown/);

    const dv = await executor.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"], __unknown: 1 },
      },
    });
    expect(dv.ok).toBe(false);
    if (!dv.ok) expect(dv.error).toMatch(/unknown rule field: __unknown/);
  });

  it("already-strict non-batch tool still rejects unknown (pageLayout.set sample)", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());
    const result = await executor.execute({
      name: "sheet.pageLayout.set",
      arguments: { sheetName: "Sheet1", orientation: "portrait", __unknown: 1 },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toMatch(/unknown field: __unknown/);
  });

  it("chart.create unknown rejection unchanged", async () => {
    const host = {
      kind: "office-js" as const,
      createChart: vi.fn(async () => ok({ name: "C1" })),
    } as unknown as HostAdapter;
    const executor = new ToolExecutor(host);
    const result = await executor.execute({
      name: "chart.create",
      arguments: {
        sheetName: "Sheet1",
        sourceRange: "A1:B2",
        __unknown: true,
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error).toMatch(/unknown field: __unknown/);
    expect((host as unknown as { createChart: { mock: { calls: unknown[] } } }).createChart.mock.calls).toHaveLength(
      0,
    );
  });
});
