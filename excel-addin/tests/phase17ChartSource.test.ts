import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { normalizeSameSheetSourceRange } from "../shared/host/officeJsChartSource";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSourceExcel } from "./fakes/officeJsChartSourceFake";
import { MockHostAdapter } from "./mockHost";

describe("phase17 chart source update", () => {
  describe("normalizeSameSheetSourceRange", () => {
    it("strips matching sheet prefix and uppercases", () => {
      expect(normalizeSameSheetSourceRange("Sheet1", "Sheet1!a1:b3")).toBe("A1:B3");
      expect(normalizeSameSheetSourceRange("Sheet1", "A1:C2")).toBe("A1:C2");
    });
    it("rejects cross-sheet, empty, multi-area, and row 0", () => {
      expect(() => normalizeSameSheetSourceRange("Sheet1", "Other!A1")).toThrow(/same worksheet/);
      expect(() => normalizeSameSheetSourceRange("Sheet1", "   ")).toThrow();
      expect(() => normalizeSameSheetSourceRange("Sheet1", "not-a-range")).toThrow(/A1/);
      expect(() => normalizeSameSheetSourceRange("Sheet1", "A1:B2,C3:D4")).toThrow(/multi-area/);
      expect(() => normalizeSameSheetSourceRange("Sheet1", "A0")).toThrow(/row must be >= 1/);
      expect(() => normalizeSameSheetSourceRange("Sheet1", "A0:B1")).toThrow(/row must be >= 1/);
    });
  });

  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSourceExcel>;
    beforeEach(() => {
      fake = installChartSourceExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
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

    it("accepts same-sheet Sheet! prefix and rejects cross-sheet", async () => {
      const adapter = new OfficeJsAdapter();
      const ok = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "Sheet1!B2:D5",
        seriesBy: "auto",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.sourceRange).toBe("B2:D5");

      const bad = await adapter.updateChartSource({
        sheetName: "Sheet1",
        chartName: "C1",
        sourceRange: "Other!A1:B2",
      });
      expect(bad.ok).toBe(false);
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
