import { describe, expect, it } from "vitest";
import {
  computeNextPivotAddress,
  parsePivotDestination,
} from "../shared/host/officeJsPivotDestination";
import { buildPivotFieldPlan } from "../shared/host/officeJsPivotFields";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { MockHostAdapter } from "./mockHost";

describe("phase45 pivot contract", () => {
  it("registers pivot tools and readonly allowlist", () => {
    const names = TOOL_DEFINITIONS.map((d) => d.name);
    expect(names).toContain("pivot.list");
    expect(names).toContain("pivot.create");
    expect(names).toContain("pivot.refresh");
    expect(TOOL_DEFINITIONS).toHaveLength(88);
    for (const name of ["pivot.list", "pivot.create", "pivot.refresh"] as const) {
      const def = TOOL_DEFINITIONS.find((d) => d.name === name)!;
      expect(def.parameters.additionalProperties).toBe(false);
    }
    expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("pivot.list");
    expect(CHAT_READONLY_TOOL_ALLOWLIST).not.toContain("pivot.create");
  });

  it("executor rejects unknown args, missing intent, non-data function/caption", async () => {
    const host = new MockHostAdapter();
    let createCalls = 0;
    const original = host.createPivot.bind(host);
    host.createPivot = async (input) => {
      createCalls += 1;
      return original(input);
    };
    const ex = new ToolExecutor(host);

    const bad = await ex.execute({
      name: "pivot.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        extra: true,
      },
    });
    expect(bad.ok).toBe(false);
    if (!bad.ok) expect(bad.error).toMatch(/unknown field/i);

    const noIntent = await ex.execute({
      name: "pivot.create",
      arguments: { sourceSheetName: "Sheet1", sourceAddress: "A1:C10" },
    });
    expect(noIntent.ok).toBe(false);
    if (!noIntent.ok) expect(noIntent.error).toMatch(/advancedIntent/);

    const rowFn = await ex.execute({
      name: "pivot.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        rowFields: [{ name: "Region", function: "sum" }],
      },
    });
    expect(rowFn.ok).toBe(false);
    if (!rowFn.ok) expect(rowFn.error).toMatch(/rowFields does not accept function\/caption/);
    expect(createCalls).toBe(0);

    const colCap = await ex.execute({
      name: "pivot.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        columnFields: [{ name: "Product", caption: "P" }],
        dataFields: ["Sales"],
      },
    });
    expect(colCap.ok).toBe(false);
    if (!colCap.ok) expect(colCap.error).toMatch(/columnFields does not accept function\/caption/);
    expect(createCalls).toBe(0);
  });

  it("empty sheetName/name rejected; destination empty uses Pivots default", async () => {
    const host = new MockHostAdapter();
    const ex = new ToolExecutor(host);

    const emptySheet = await ex.execute({
      name: "pivot.list",
      arguments: { sheetName: "" },
    });
    expect(emptySheet.ok).toBe(false);
    if (!emptySheet.ok) expect(emptySheet.error).toMatch(/sheetName must be non-empty/);

    const emptyName = await ex.execute({
      name: "pivot.refresh",
      arguments: { advancedIntent: "interactive-pivot", name: "   " },
    });
    expect(emptyName.ok).toBe(false);
    if (!emptyName.ok) expect(emptyName.error).toMatch(/name must be non-empty/);

    const created = await ex.execute({
      name: "pivot.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        destination: "",
        rowFields: ["Region"],
        dataFields: ["Sales"],
      },
    });
    expect(created.ok).toBe(true);
    if (created.ok) {
      expect((created.data as { sheetName: string }).sheetName).toBe("Pivots");
    }
  });

  it("field plan: conflicts, empty layout, multi data agg; destination parse", () => {
    expect(() =>
      buildPivotFieldPlan({
        sourceSheetName: "S",
        sourceAddress: "A1",
        rowFields: ["Region"],
        columnFields: ["Region"],
      }),
    ).toThrow(/both rowFields and columnFields/);

    expect(() =>
      buildPivotFieldPlan({ sourceSheetName: "S", sourceAddress: "A1" }),
    ).toThrow(/at least one field/);

    expect(() =>
      buildPivotFieldPlan({
        sourceSheetName: "S",
        sourceAddress: "A1",
        dataFields: [{ name: "Sales", function: "median" as "sum" }],
      }),
    ).toThrow(/sum\|count/);

    const multi = buildPivotFieldPlan({
      sourceSheetName: "S",
      sourceAddress: "A1",
      dataFields: [
        { name: "Sales", function: "sum", caption: "Sum Sales" },
        { name: "Sales", function: "count", caption: "Count Sales" },
      ],
    });
    expect(multi.dataFields).toHaveLength(2);
    expect(multi.dataFields.map((f) => f.function)).toEqual(["sum", "count"]);

    expect(() => parsePivotDestination("Sheet1!A1,B1")).toThrow(/multi-area/);
    expect(() => parsePivotDestination("[Book.xlsx]Sheet1!A1")).toThrow(/structured|external/i);
    expect(parsePivotDestination(undefined).useDedicatedSheet).toBe(true);
    expect(parsePivotDestination("").useDedicatedSheet).toBe(true);
    expect(parsePivotDestination("A5")).toEqual({
      useDedicatedSheet: false,
      sheetName: null,
      address: "A5",
    });
    expect(parsePivotDestination("'Sheet 2'!B2").sheetName).toBe("Sheet 2");
  });

  it("computeNextPivotAddress matches desktop lastBottom+3", () => {
    expect(computeNextPivotAddress(0)).toBe("A1");
    expect(computeNextPivotAddress(-1)).toBe("A1");
    expect(computeNextPivotAddress(1)).toBe("A4");
    expect(computeNextPivotAddress(8)).toBe("A11");
    expect(computeNextPivotAddress(20)).toBe("A23");
  });

  it("executor rejects explicit null for sheetName/name/destination before Host", async () => {
    const host = new MockHostAdapter();
    let listCalls = 0;
    let createCalls = 0;
    let refreshCalls = 0;
    const listOrig = host.listPivots.bind(host);
    const createOrig = host.createPivot.bind(host);
    const refreshOrig = host.refreshPivots.bind(host);
    host.listPivots = async (input) => {
      listCalls += 1;
      return listOrig(input);
    };
    host.createPivot = async (input) => {
      createCalls += 1;
      return createOrig(input);
    };
    host.refreshPivots = async (input) => {
      refreshCalls += 1;
      return refreshOrig(input);
    };
    const ex = new ToolExecutor(host);

    const listNull = await ex.execute({
      name: "pivot.list",
      arguments: { sheetName: null },
    });
    expect(listNull.ok).toBe(false);
    if (!listNull.ok) expect(listNull.error).toMatch(/sheetName must not be null/);
    expect(listCalls).toBe(0);

    const refreshNull = await ex.execute({
      name: "pivot.refresh",
      arguments: { advancedIntent: "interactive-pivot", name: null },
    });
    expect(refreshNull.ok).toBe(false);
    if (!refreshNull.ok) expect(refreshNull.error).toMatch(/name must not be null/);
    expect(refreshCalls).toBe(0);

    const destNull = await ex.execute({
      name: "pivot.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        destination: null,
        rowFields: ["Region"],
        dataFields: ["Sales"],
      },
    });
    expect(destNull.ok).toBe(false);
    if (!destNull.ok) expect(destNull.error).toMatch(/destination must not be null/);
    expect(createCalls).toBe(0);

    // Explicit "" destination remains default Pivots (documented), and Host is called.
    const destEmpty = await ex.execute({
      name: "pivot.create",
      arguments: {
        advancedIntent: "interactive-pivot",
        sourceSheetName: "Sheet1",
        sourceAddress: "A1:C10",
        destination: "",
        rowFields: ["Region"],
        dataFields: ["Sales"],
      },
    });
    expect(destEmpty.ok).toBe(true);
    expect(createCalls).toBe(1);
    if (destEmpty.ok) {
      expect((destEmpty.data as { sheetName: string }).sheetName).toBe("Pivots");
    }
  });
});
