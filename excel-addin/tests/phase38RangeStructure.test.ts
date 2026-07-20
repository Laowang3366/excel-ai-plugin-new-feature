import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { buildAdvancedExcelBoundary } from "../shared/prompts/advancedExcelBoundary";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { MockHostAdapter } from "./mockHost";
import { installRangeStructureExcel } from "./fakes/officeJsRangeStructureFake";

describe("phase38 range structure", () => {
  let fake: ReturnType<typeof installRangeStructureExcel>;

  beforeEach(() => {
    fake = installRangeStructureExcel({
      insertedAddress: "HostSheet!$A$1:$B$2",
      deletedAddress: "HostSheet!$C$3:$D$4",
      autofitAddress: "HostSheet!$E$5:$F$6",
      columnWidth: 81.25,
      rowHeight: 24.5,
    });
  });

  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
  });

  it("registers three moderate, closed schemas", () => {
    for (const [name, required, enumValues] of [
      ["range.insert", ["sheetName", "range", "shift"], ["down", "right"]],
      ["range.delete", ["sheetName", "range", "shift"], ["up", "left"]],
      ["range.autofit", ["sheetName", "range", "direction"], ["rows", "columns", "both"]],
    ] as const) {
      const definition = TOOL_DEFINITIONS.find((item) => item.name === name);
      expect(definition?.riskLevel, name).toBe("moderate");
      const params = definition?.parameters as {
        additionalProperties?: boolean;
        properties?: Record<string, { enum?: string[] }>;
        required?: string[];
      };
      expect(params.additionalProperties, name).toBe(false);
      expect(params.required, name).toEqual(required);
      expect(params.properties?.[required[2]]?.enum, name).toEqual(enumValues);
    }
  });

  it("maps insert shift and returns the loaded host address", async () => {
    const result = await new OfficeJsAdapter().insertRange({
      sheetName: "Sheet1",
      address: "A1:B2",
      shift: "down",
    });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.data).toEqual({
        sheetName: "Sheet1",
        address: "HostSheet!$A$1:$B$2",
        shift: "down",
        operation: "insert",
      });
    }
    expect(fake.insertShift()).toBe("Down");
    expect(fake.syncCalls()).toBe(2);
  });

  it("maps delete shift and captures the address before deletion", async () => {
    const result = await new OfficeJsAdapter().deleteRange({
      sheetName: "Sheet1",
      address: "C3:D4",
      shift: "left",
    });
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.data).toEqual({
        sheetName: "Sheet1",
        address: "HostSheet!$C$3:$D$4",
        shift: "left",
        operation: "delete",
      });
    }
    expect(fake.deleteShift()).toBe("Left");
    expect(fake.syncCalls()).toBe(2);
  });

  it("autofits the requested dimensions and reads back actual sizes", async () => {
    const adapter = new OfficeJsAdapter();
    const rows = await adapter.autofitRange({
      sheetName: "Sheet1",
      address: "E5:F6",
      direction: "rows",
    });
    const columns = await adapter.autofitRange({
      sheetName: "Sheet1",
      address: "E5:F6",
      direction: "columns",
    });
    const both = await adapter.autofitRange({
      sheetName: "Sheet1",
      address: "E5:F6",
      direction: "both",
    });
    for (const result of [rows, columns, both]) {
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.address).toBe("HostSheet!$E$5:$F$6");
        expect(result.data.columnWidth).toBe(81.25);
        expect(result.data.rowHeight).toBe(24.5);
      }
    }
    expect(fake.autofitRowsCalls()).toBe(2);
    expect(fake.autofitColumnsCalls()).toBe(2);
  });

  it("gates autofit on ExcelApi 1.2 before Excel.run", async () => {
    for (const options of [
      { excelApi12: false },
      { missingIsSetSupported: true },
      { isSetSupportedThrows: true },
    ]) {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      fake = installRangeStructureExcel(options);
      const result = await new OfficeJsAdapter().autofitRange({
        sheetName: "Sheet1",
        address: "A1",
        direction: "both",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.2/);
      }
      expect(fake.runCalls()).toBe(0);
      expect(fake.autofitRowsCalls()).toBe(0);
      expect(fake.autofitColumnsCalls()).toBe(0);
    }
  });

  it("classifies missing host methods and malformed dimensions as ordinary failures", async () => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
    installRangeStructureExcel({ missingAutofitRows: true });
    const missing = await new OfficeJsAdapter().autofitRange({
      sheetName: "Sheet1",
      address: "A1",
      direction: "rows",
    });
    expect(missing.ok).toBe(false);
    if (!missing.ok) {
      expect(missing.unsupported).not.toBe(true);
      expect(missing.reason).toMatch(/autofitRows/);
    }

    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
    installRangeStructureExcel({ rowHeight: "not-a-number" });
    const malformed = await new OfficeJsAdapter().autofitRange({
      sheetName: "Sheet1",
      address: "A1",
      direction: "rows",
    });
    expect(malformed.ok).toBe(false);
    if (!malformed.ok) expect(malformed.unsupported).not.toBe(true);
  });

  it("executor trims identifiers and rejects unknown, null, undefined, and enum errors", async () => {
    const host = new MockHostAdapter();
    const executor = new ToolExecutor(host);
    const inserted = await executor.execute({
      name: "range.insert",
      arguments: { sheetName: "  Sheet1 ", range: " A1:B2 ", shift: "down" },
    });
    expect(inserted.ok).toBe(true);
    if (inserted.ok) expect(inserted.data).toMatchObject({ address: "Sheet1!A1:B2" });

    const valid = [
      { name: "range.delete", arguments: { sheetName: "Sheet1", range: "A1", shift: "up" } },
      {
        name: "range.autofit",
        arguments: { sheetName: "Sheet1", range: "A1", direction: "both" },
      },
    ] as const;
    for (const call of valid) expect((await executor.execute(call)).ok).toBe(true);

    const invalid = [
      { name: "range.insert", arguments: { sheetName: "Sheet1", range: "A1" } },
      { name: "range.insert", arguments: { sheetName: "Sheet1", range: "A1", shift: null } },
      { name: "range.insert", arguments: { sheetName: "Sheet1", range: "A1", shift: undefined } },
      { name: "range.insert", arguments: { sheetName: "Sheet1", range: "A1", shift: "up" } },
      { name: "range.delete", arguments: { sheetName: "Sheet1", range: "A1", shift: "right" } },
      {
        name: "range.autofit",
        arguments: { sheetName: "Sheet1", range: "A1", direction: "width" },
      },
      {
        name: "range.autofit",
        arguments: { sheetName: "Sheet1", range: "A1", direction: "rows", extra: true },
      },
    ] as const;
    for (const call of invalid) expect((await executor.execute(call as never)).ok).toBe(false);
  });

  it("returns typed unsupported for all three operations on WPS", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    const calls = [
      { name: "range.insert" as const, arguments: { sheetName: "Sheet1", range: "A1", shift: "down" } },
      { name: "range.delete" as const, arguments: { sheetName: "Sheet1", range: "A1", shift: "up" } },
      { name: "range.autofit" as const, arguments: { sheetName: "Sheet1", range: "A1", direction: "rows" } },
    ];
    for (const call of calls) {
      const result = await executor.execute(call);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });

  it("documents the API gates and WPS boundary", () => {
    const boundary = buildAdvancedExcelBoundary({});
    expect(boundary).toContain("`range.insert`");
    expect(boundary).toContain("`range.delete`");
    expect(boundary).toContain("`range.autofit`");
    expect(boundary).toMatch(/ExcelApi 1\.2/);
    expect(boundary).toMatch(/WPS.*unsupported/);
  });
});
