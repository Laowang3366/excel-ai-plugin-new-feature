import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSeriesAddExcel } from "./fakes/officeJsChartSeriesAddFake";
import { MockHostAdapter } from "./mockHost";

describe("phase22 chart series add", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesAddExcel>;
    beforeEach(() => {
      fake = installChartSeriesAddExcel([
        { name: "S1", chartType: "ColumnClustered", smooth: false },
      ]);
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("appends named empty series with continuous 1-based index", async () => {
      const result = await new OfficeJsAdapter().addChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        name: "NewS",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          addedSeries: { index: 2, name: "NewS", chartType: "column", smooth: false },
          dataBound: false,
        });
      }
      expect(fake.getCommittedNames()).toEqual(["S1", "NewS"]);
    });

    it("appends after two existing series as index 3", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      fake = installChartSeriesAddExcel([
        { name: "A", chartType: "ColumnClustered", smooth: false },
        { name: "B", chartType: "Line", smooth: true },
      ]);
      const result = await new OfficeJsAdapter().addChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        name: "Third",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.addedSeries.index).toBe(3);
        expect(result.data.addedSeries.name).toBe("Third");
        expect(result.data.dataBound).toBe(false);
      }
      expect(fake.getCommittedNames()).toEqual(["A", "B", "Third"]);
    });

    it("preserves unknown host chartType on add readback", async () => {
      fake.setNextAddChartType("CustomUnknownType");
      const result = await new OfficeJsAdapter().addChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        name: "U",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.addedSeries.chartType).toBe("CustomUnknownType");
        expect(result.data.addedSeries.index).toBe(2);
      }
    });

    it("omitted name uses host default and empty start works", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      fake = installChartSeriesAddExcel([]);
      const result = await new OfficeJsAdapter().addChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.dataBound).toBe(false);
        expect(result.data.addedSeries.index).toBe(1);
        expect(result.data.addedSeries.name).toBe("Series1");
        expect(result.data.addedSeries.chartType).toBe("column");
      }
    });

    it("skipping add sync keeps old items snapshot", async () => {
      const names = await fake.brokenAddSkipFirstSync("Ghost");
      expect(names).toEqual(["S1"]);
      expect(fake.getCommittedNames()).toEqual(["S1", "Ghost"]);

      const ok = await new OfficeJsAdapter().addChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        name: "After",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) {
        expect(ok.data.addedSeries).toEqual({
          index: 3,
          name: "After",
          chartType: "column",
          smooth: false,
        });
      }
    });

    it("chartName comes from host load not input", async () => {
      fake.setLoadedChartName("HostChart");
      const result = await new OfficeJsAdapter().addChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        name: "X",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.chartName).toBe("HostChart");
        expect(result.data.chartName).not.toBe("C1");
      }
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity with and without name", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      const named = await executor.execute({
        name: "chart.series.add",
        arguments: { sheetName: "Sheet1", chartName: "C1", name: "Extra" },
      });
      expect(named.ok).toBe(true);
      if (named.ok) {
        expect(named.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          addedSeries: { index: 3, name: "Extra", chartType: "column", smooth: false },
          dataBound: false,
        });
      }

      const def = await executor.execute({
        name: "chart.series.add",
        arguments: { sheetName: "Sheet1", chartName: "C1" },
      });
      expect(def.ok).toBe(true);
      if (def.ok) {
        expect(def.data).toMatchObject({
          dataBound: false,
          addedSeries: { index: 4, name: "Series4", chartType: "column", smooth: false },
        });
      }
    });

    it("trims sheetName/chartName/name before host call", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const result = await new ToolExecutor(host).execute({
        name: "chart.series.add",
        arguments: {
          sheetName: "  Sheet1  ",
          chartName: "  C1  ",
          name: "  Trimmed  ",
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          addedSeries: { index: 3, name: "Trimmed", chartType: "column", smooth: false },
          dataBound: false,
        });
      }
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1" },
        { chartName: "C1" },
        { sheetName: "", chartName: "C1" },
        { sheetName: "   ", chartName: "C1" },
        { sheetName: "Sheet1", chartName: "" },
        { sheetName: "Sheet1", chartName: "  " },
        { sheetName: "Sheet1", chartName: "C1", name: "" },
        { sheetName: "Sheet1", chartName: "C1", name: "  " },
        { sheetName: "Sheet1", chartName: "C1", name: null },
        { sheetName: "Sheet1", chartName: "C1", name: undefined },
        { sheetName: "Sheet1", chartName: "C1", extra: 1 },
        { sheetName: null, chartName: "C1" },
        { sheetName: "Sheet1", chartName: null },
        { sheetName: undefined, chartName: "C1" },
        { sheetName: "Sheet1", chartName: undefined },
        { sheetName: 1, chartName: "C1" },
        { sheetName: true, chartName: "C1" },
        { sheetName: { a: 1 }, chartName: "C1" },
        { sheetName: ["Sheet1"], chartName: "C1" },
        { sheetName: "Sheet1", chartName: 2 },
        { sheetName: "Sheet1", chartName: false },
        { sheetName: "Sheet1", chartName: { c: 1 } },
        { sheetName: "Sheet1", chartName: ["C1"] },
        { sheetName: "Sheet1", chartName: "C1", name: 3 },
        { sheetName: "Sheet1", chartName: "C1", name: true },
        { sheetName: "Sheet1", chartName: "C1", name: { n: 1 } },
        { sheetName: "Sheet1", chartName: "C1", name: ["N"] },
      ]) {
        const result = await executor.execute({
          name: "chart.series.add",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.add");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS is typed unsupported with ChartSeriesCollection.add evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.add",
      arguments: { sheetName: "Sheet1", chartName: "C1", name: "N" },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.add");
      expect(detail.evidence).toMatch(/ChartSeriesCollection\.add/);
    }
  });
});
