import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSeriesDeleteExcel } from "./fakes/officeJsChartSeriesDeleteFake";
import { MockHostAdapter } from "./mockHost";

describe("phase21 chart series delete", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesDeleteExcel>;
    beforeEach(() => {
      fake = installChartSeriesDeleteExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("deletes first, middle, and last with continuous 1-based readback", async () => {
      const mid = await new OfficeJsAdapter().deleteChartSeries("Sheet1", "C1", 2);
      expect(mid.ok).toBe(true);
      if (mid.ok) {
        expect(mid.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          deletedSeriesIndex: 2,
          remainingSeries: [
            { index: 1, name: "S1", chartType: "column", smooth: false },
            { index: 2, name: "S3", chartType: "line", smooth: true },
          ],
        });
      }
      expect(fake.getCommittedNames()).toEqual(["S1", "S3"]);

      const first = await new OfficeJsAdapter().deleteChartSeries("Sheet1", "C1", 1);
      expect(first.ok).toBe(true);
      if (first.ok) {
        expect(first.data.deletedSeriesIndex).toBe(1);
        expect(first.data.remainingSeries).toEqual([
          { index: 1, name: "S3", chartType: "line", smooth: true },
        ]);
      }

      delete (globalThis as { Excel?: unknown }).Excel;
      fake = installChartSeriesDeleteExcel();
      const last = await new OfficeJsAdapter().deleteChartSeries("Sheet1", "C1", 3);
      expect(last.ok).toBe(true);
      if (last.ok) {
        expect(last.data.deletedSeriesIndex).toBe(3);
        expect(last.data.remainingSeries.map((s) => s.name)).toEqual(["S1", "S2"]);
        expect(last.data.remainingSeries.map((s) => s.index)).toEqual([1, 2]);
      }
    });

    it("skipping delete sync keeps old items snapshot", async () => {
      const names = await fake.brokenDeleteSkipFirstSync(2);
      expect(names).toEqual(["S1", "S2", "S3"]);
      expect(fake.getCommittedNames()).toEqual(["S1", "S3"]);

      const ok = await new OfficeJsAdapter().deleteChartSeries("Sheet1", "C1", 1);
      expect(ok.ok).toBe(true);
      if (ok.ok) {
        expect(ok.data.remainingSeries.map((s) => s.name)).toEqual(["S3"]);
        expect(ok.data.remainingSeries.map((s) => s.index)).toEqual([1]);
      }
    });

    it("out of range fails", async () => {
      const result = await new OfficeJsAdapter().deleteChartSeries("Sheet1", "C1", 99);
      expect(result.ok).toBe(false);
    });

    it("chartName comes from host load, not input echo", async () => {
      fake.setLoadedChartName("HostChartName");
      const result = await new OfficeJsAdapter().deleteChartSeries("Sheet1", "C1", 1);
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.chartName).toBe("HostChartName");
        expect(result.data.chartName).not.toBe("C1");
      }
    });

    it("non-string host chart.name fails", async () => {
      fake.setLoadedChartName(42);
      const result = await new OfficeJsAdapter().deleteChartSeries("Sheet1", "C1", 1);
      expect(result.ok).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity remaining series", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "chart.series.delete",
        arguments: { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          deletedSeriesIndex: 1,
          remainingSeries: [{ index: 1, name: "Series2", chartType: "column", smooth: false }],
        });
      }
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: -1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1.5 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: Number.NaN },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: Number.POSITIVE_INFINITY },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: undefined },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, extra: 1 },
        { sheetName: "", chartName: "C1", seriesIndex: 1 },
        { sheetName: "   ", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "  \t  ", seriesIndex: 1 },
        { sheetName: undefined, chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: undefined, seriesIndex: 1 },
        { sheetName: null, chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: null, seriesIndex: 1 },
      ]) {
        const result = await executor.execute({
          name: "chart.series.delete",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.delete");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS is typed unsupported with ChartSeries.delete evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.delete",
      arguments: { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.delete");
      expect(detail.evidence).toMatch(/ChartSeries\.delete/);
    }
  });
});
