import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor, TOOL_DEFINITIONS } from "../shared/tools";
import { installDisplayExcel } from "./fakes/officeJsDisplayFake";
import { MockHostAdapter } from "./mockHost";

describe("phase9 sheet.display", () => {
  it("registers sheet.display.get and sheet.display.set", () => {
    const names = TOOL_DEFINITIONS.map((tool) => tool.name);
    expect(names).toContain("sheet.display.get");
    expect(names).toContain("sheet.display.set");
  });

  describe("Office.js", () => {
    beforeEach(() => {
      installDisplayExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("reads default display props", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.getSheetDisplay("Sheet1");
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sheetName).toBe("Sheet1");
        expect(result.data.tabColor).toBe("");
        expect(result.data.showGridlines).toBe(true);
        expect(result.data.showHeadings).toBe(true);
      }
    });

    it("sets single and multi fields with writeback", async () => {
      const adapter = new OfficeJsAdapter();
      const color = await adapter.setSheetDisplay({
        sheetName: "Sheet1",
        tabColor: "#FF0000",
      });
      expect(color.ok).toBe(true);
      if (color.ok) expect(color.data.tabColor).toBe("#FF0000");

      const multi = await adapter.setSheetDisplay({
        sheetName: "Sheet1",
        showGridlines: false,
        showHeadings: false,
      });
      expect(multi.ok).toBe(true);
      if (multi.ok) {
        expect(multi.data.tabColor).toBe("#FF0000");
        expect(multi.data.showGridlines).toBe(false);
        expect(multi.data.showHeadings).toBe(false);
      }

      const auto = await adapter.setSheetDisplay({
        sheetName: "Sheet1",
        tabColor: "",
      });
      expect(auto.ok).toBe(true);
      if (auto.ok) expect(auto.data.tabColor).toBe("");
    });
  });

  it("executor success and validation branches", async () => {
    const executor = new ToolExecutor(new MockHostAdapter());

    expect(
      (
        await executor.execute({
          name: "sheet.display.get",
          arguments: { sheetName: "Sheet1" },
        })
      ).ok,
    ).toBe(true);

    expect(
      (
        await executor.execute({
          name: "sheet.display.set",
          arguments: { sheetName: "Sheet1", tabColor: "00ff00" },
        })
      ).ok,
    ).toBe(true);

    const normalized = await executor.execute({
      name: "sheet.display.set",
      arguments: { sheetName: "Sheet1", tabColor: "#aabbcc" },
    });
    expect(normalized.ok).toBe(true);
    if (normalized.ok) {
      expect((normalized.data as { tabColor: string }).tabColor).toBe("#AABBCC");
    }

    expect(
      (
        await executor.execute({
          name: "sheet.display.set",
          arguments: { sheetName: "Sheet1" },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "sheet.display.set",
          arguments: { sheetName: "Sheet1", tabColor: "red" },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "sheet.display.set",
          arguments: { sheetName: "Sheet1", unknown: true },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "sheet.display.set",
          arguments: { sheetName: "Sheet1", showGridlines: "yes" },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "sheet.display.set",
          arguments: { sheetName: "Sheet1", tabColor: "#AABBCC", showGridlines: null },
        })
      ).ok,
    ).toBe(false);

    expect(
      (
        await executor.execute({
          name: "sheet.display.set",
          arguments: { sheetName: "Sheet1", tabColor: null },
        })
      ).ok,
    ).toBe(false);
  });

  it("WPS returns unsupported for display get/set", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    for (const name of ["sheet.display.get", "sheet.display.set"] as const) {
      const result = await executor.execute({
        name,
        arguments:
          name === "sheet.display.get"
            ? { sheetName: "Sheet1" }
            : { sheetName: "Sheet1", showGridlines: false },
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });
});
