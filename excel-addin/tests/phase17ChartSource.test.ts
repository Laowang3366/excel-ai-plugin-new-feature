import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import {
  normalizeSameSheetSourceRange,
  parseChartSourceRange,
} from "../shared/host/officeJsChartSource";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSourceExcel } from "./fakes/officeJsChartSourceFake";
import { MockHostAdapter } from "./mockHost";

describe("phase17 chart source update", () => {
  describe("parseChartSourceRange", () => {
    it("parses bare A1 and same-sheet prefix", () => {
      expect(parseChartSourceRange("Sheet1", "a1:b3")).toEqual({
        sourceSheetName: "Sheet1",
        bareA1: "A1:B3",
        displaySourceRange: "A1:B3",
      });
      expect(parseChartSourceRange("Sheet1", "Sheet1!A1:C2").displaySourceRange).toBe("A1:C2");
    });

    it("parses cross-sheet and quoted sheet names", () => {
      const cross = parseChartSourceRange("Sheet1", "Sheet2!A1:B10");
      expect(cross.sourceSheetName).toBe("Sheet2");
      expect(cross.bareA1).toBe("A1:B10");
      expect(cross.displaySourceRange).toBe("Sheet2!A1:B10");

      const spaced = parseChartSourceRange("Sheet1", "'Sheet 2'!A1:B2");
      expect(spaced.sourceSheetName).toBe("Sheet 2");
      expect(spaced.displaySourceRange).toBe("'Sheet 2'!A1:B2");
    });

    it("parses and re-escapes sheet names containing apostrophes (Excel '')", () => {
      const parsed = parseChartSourceRange("Sheet1", "'O''Brien'!A1:B2");
      expect(parsed.sourceSheetName).toBe("O'Brien");
      expect(parsed.bareA1).toBe("A1:B2");
      expect(parsed.displaySourceRange).toBe("'O''Brien'!A1:B2");
    });

    it("rejects external, 3D, multi-area, structured, and invalid A1", () => {
      expect(() => parseChartSourceRange("Sheet1", "[Book.xlsx]Sheet1!A1")).toThrow(/external/i);
      expect(() => parseChartSourceRange("Sheet1", "'[Book.xlsx]Data'!A1")).toThrow(/external/i);
      expect(() => parseChartSourceRange("Sheet1", "Sheet1:Sheet3!A1")).toThrow(/3D/i);
      expect(() => parseChartSourceRange("Sheet1", "A1:B2,C3:D4")).toThrow(/multi-area/i);
      expect(() => parseChartSourceRange("Sheet1", "Table1[Col]")).toThrow(/structured/i);
      expect(() => parseChartSourceRange("Sheet1", "   ")).toThrow();
      expect(() => parseChartSourceRange("Sheet1", "not-a-range")).toThrow(/A1/i);
      expect(() => parseChartSourceRange("Sheet1", "A0")).toThrow(/row must be >= 1/);
    });
  });

  describe("normalizeSameSheetSourceRange (series helpers)", () => {
    it("strips matching sheet prefix and uppercases", () => {
      expect(normalizeSameSheetSourceRange("Sheet1", "Sheet1!a1:b3")).toBe("A1:B3");
      expect(normalizeSameSheetSourceRange("Sheet1", "A1:C2")).toBe("A1:C2");
    });
    it("still rejects cross-sheet for same-sheet helpers", () => {
      expect(() => normalizeSameSheetSourceRange("Sheet1", "Other!A1")).toThrow(/same worksheet/);
      expect(() => normalizeSameSheetSourceRange("Sheet1", "A1:B2,C3:D4")).toThrow(/multi-area/);
    });
  });

  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSourceExcel>;
    beforeEach(() => {
      fake = installChartSourceExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { window?: unknown }).window;
    });

    it("setData is pending until first sync; series only after load+sync", async () => {
      const adapter = new OfficeJsAdapter();
      expect(fake.getPending()).toBeUndefined();
      expect(fake.getItemsVisible()).toBe(false);
      const result = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "A1:C3",
        seriesBy: "columns",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sourceRange).toBe("A1:C3");
        expect(result.data.seriesBy).toBe("columns");
        expect(result.data.series.length).toBe(2);
        expect(result.data.series[0]?.index).toBe(1);
        expect(result.data.series[0]?.name).toBe("C1");
        expect(result.data.series[1]?.name).toBe("C2");
        expect(result.data.chartName).toBe("C1");
      }
      expect(fake.getPending()).toBeUndefined();
      expect(fake.getCommitted().sourceRange).toBe("A1:C3");
      expect(fake.getCommitted().sourceSheet).toBe("Sheet1");
      expect(fake.getItemsVisible()).toBe(true);
    });

    it("supports seriesBy rows and auto default", async () => {
      const adapter = new OfficeJsAdapter();
      const rows = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "A1:B4",
        seriesBy: "rows",
      });
      expect(rows.ok).toBe(true);
      if (rows.ok) {
        expect(rows.data.seriesBy).toBe("rows");
        expect(rows.data.series.map((s) => s.name)).toEqual(["R1", "R2", "R3"]);
      }
      const auto = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "A1:D2",
      });
      expect(auto.ok).toBe(true);
      if (auto.ok) {
        expect(auto.data.seriesBy).toBe("auto");
        expect(auto.data.series).toHaveLength(3);
      }
    });

    it("accepts same-sheet Sheet! prefix", async () => {
      const adapter = new OfficeJsAdapter();
      const ok = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "Sheet1!B2:D5",
        seriesBy: "auto",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.sourceRange).toBe("B2:D5");
    });

    it("supports cross-sheet source via setData Range from source worksheet", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "Sheet2!A1:C3",
        seriesBy: "columns",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sourceRange).toBe("Sheet2!A1:C3");
        expect(result.data.series.map((s) => s.name)).toEqual(["C1", "C2"]);
      }
      expect(fake.getCommitted().sourceSheet).toBe("Sheet2");
      expect(fake.getCommitted().sourceRange).toBe("A1:C3");
    });

    it("supports quoted sheet names with spaces", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "'Sheet 2'!A1:B4",
        seriesBy: "rows",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sourceRange).toBe("'Sheet 2'!A1:B4");
        expect(result.data.series.map((s) => s.name)).toEqual(["R1", "R2", "R3"]);
      }
      expect(fake.getCommitted().sourceSheet).toBe("Sheet 2");
    });

    it("supports apostrophe sheet names via Excel '' escaping", async () => {
      fake = installChartSourceExcel({ sheets: ["Sheet1", "O'Brien"] });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "'O''Brien'!A1:B4",
        seriesBy: "columns",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sourceRange).toBe("'O''Brien'!A1:B4");
      }
      expect(fake.getCommitted().sourceSheet).toBe("O'Brien");
    });

    it("fails clearly when source sheet is missing", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "Missing!A1:B2",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.reason).toMatch(/not found|ItemNotFound|Missing/i);
      }
    });

    it("rejects illegal external workbook references", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "[Other.xlsx]Sheet1!A1:B2",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.reason).toMatch(/external/i);
    });

    it("fails when Chart.setData is missing on host", async () => {
      fake = installChartSourceExcel({ withSetData: false });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "A1:B2",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.reason).toMatch(/setData/i);
    });

    it("skipping first sync after setData yields old series snapshot", async () => {
      const broken = await fake.brokenUpdateSkipFirstSync("A1:C3");
      expect(broken.map((s) => s.name)).toEqual(["Series1", "Series2"]);
      expect(fake.getCommitted().sourceRange).toBe("A1:C3");
      expect(fake.getCommitted().series.map((s) => s.name)).toEqual(["C1", "C2"]);

      const adapter = new OfficeJsAdapter();
      const okPath = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "A1:D2",
        seriesBy: "auto",
      });
      expect(okPath.ok).toBe(true);
      if (okPath.ok) {
        expect(okPath.data.series.map((s) => s.name)).toEqual(["C1", "C2", "C3"]);
      }
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity for columns", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const result = await new ToolExecutor(host).execute({
        name: "chart.source.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          sourceRange: "A1:C4",
          seriesBy: "columns",
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as {
          seriesBy: string;
          series: { index: number; name: string }[];
        };
        expect(data.seriesBy).toBe("columns");
        expect(data.series.map((s) => s.index)).toEqual([1, 2]);
        expect(data.series.map((s) => s.name)).toEqual(["C1", "C2"]);
      }
    });

    it("MockHost parity for rows", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const result = await new ToolExecutor(host).execute({
        name: "chart.source.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          sourceRange: "A1:B4",
          seriesBy: "rows",
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as { seriesBy: string; series: { name: string }[] };
        expect(data.seriesBy).toBe("rows");
        expect(data.series.map((s) => s.name)).toEqual(["R1", "R2", "R3"]);
      }
    });

    it("MockHost parity for auto default", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const result = await new ToolExecutor(host).execute({
        name: "chart.source.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          sourceRange: "A1:D2",
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as { seriesBy: string; series: { name: string }[] };
        expect(data.seriesBy).toBe("auto");
        expect(data.series.map((s) => s.name)).toEqual(["C1", "C2", "C3"]);
      }
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1" },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: "" },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: "   " },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: null },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: undefined },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: "A1", seriesBy: "diagonal" },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: "A1", seriesBy: null },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: "A1", seriesBy: undefined },
        { sheetName: "Sheet1", chartName: "C1", sourceRange: "A1", extra: true },
        { sheetName: "", chartName: "C1", sourceRange: "A1" },
        { sheetName: null, chartName: "C1", sourceRange: "A1" },
        { sheetName: "Sheet1", chartName: null, sourceRange: "A1" },
      ]) {
        const result = await executor.execute({
          name: "chart.source.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.source.update");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
      expect(def?.description).toMatch(/跨表|Sheet2/i);
    });
  });

  it("WPS chart.source.update is typed unsupported with setData evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.source.update",
      arguments: { sheetName: "Sheet1", chartName: "C1", sourceRange: "A1:B2" },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.source.update");
      expect(detail.evidence).toMatch(/Chart\.setData/);
    }
  });
});
