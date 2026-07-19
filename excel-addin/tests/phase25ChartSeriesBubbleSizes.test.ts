import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { normalizeSameSheetSourceRange } from "../shared/host/officeJsChartSource";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSeriesBubbleSizesExcel } from "./fakes/officeJsChartSeriesBubbleSizesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase25 chart series bubble sizes", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesBubbleSizesExcel>;
    beforeEach(() => {
      fake = installChartSeriesBubbleSizesExcel({ hostSourcePrefix: "HostSheet!" });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("binds bubble sizes with host sheet/chart names and source string (not input echo)", async () => {
      fake.setLoadedSheetName("HostSheet");
      fake.setLoadedChartName("HostChart");
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "C2:C5",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sheetName).toBe("HostSheet");
        expect(result.data.sheetName).not.toBe("Sheet1");
        expect(result.data.chartName).toBe("HostChart");
        expect(result.data.chartName).not.toBe("C1");
        expect(result.data.bubbleSizesSource).toBe("HostSheet!C2:C5");
        expect(result.data.bubbleSizesSource).not.toBe("C2:C5");
        expect(result.data.dataBound).toBe(true);
        expect(result.data.seriesIndex).toBe(1);
      }
      expect(fake.getCommitted(0)?.bubbleSizesSource).toBe("HostSheet!C2:C5");
      expect(fake.getCommitted(1)?.bubbleSizesSource).toBeNull();
    });

    it("seriesIndex 2 writes only second series; first unchanged", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        bubbleSizesRange: "Z2:Z4",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.seriesIndex).toBe(2);
        expect(result.data.bubbleSizesSource).toBe("HostSheet!Z2:Z4");
      }
      expect(fake.getCommitted(0)?.bubbleSizesSource).toBeNull();
      expect(fake.getCommitted(1)?.bubbleSizesSource).toBe("HostSheet!Z2:Z4");
    });

    it("skipping first sync cannot read new source string", async () => {
      const broken = await fake.brokenSkipFirstSync(1, "E2:E3");
      expect(broken).toBeNull();
      expect(fake.getCommitted(0)?.bubbleSizesSource).toBe("HostSheet!E2:E3");

      const ok = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "F2:F3",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.bubbleSizesSource).toBe("HostSheet!F2:F3");
    });

    it("stale A→B without write-first-sync still reads A not B", async () => {
      const boundA = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "A2:A4",
      });
      expect(boundA.ok).toBe(true);
      if (boundA.ok) expect(boundA.data.bubbleSizesSource).toBe("HostSheet!A2:A4");
      expect(fake.getCommitted(0)?.bubbleSizesSource).toBe("HostSheet!A2:A4");

      const stale = await fake.brokenSkipFirstSync(1, "B2:B4");
      expect(stale).toBe("HostSheet!A2:A4");
      expect(stale).not.toBe("HostSheet!B2:B4");
      expect(fake.getCommitted(0)?.bubbleSizesSource).toBe("HostSheet!B2:B4");
    });

    it("ExcelApi 1.15 precheck false is unsupported and never calls setBubbleSizes", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installChartSeriesBubbleSizesExcel({ excelApi115: false });
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "B2:B3",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.evidence).toMatch(/ExcelApi 1\.15/);
        expect(result.reason).toMatch(/isSetSupported|1\.15/);
      }
      expect(f.getSetterCallCounts()).toEqual({ setBubbleSizesCalls: 0 });
      expect(f.getCommitted(0)?.bubbleSizesSource).toBeNull();
    });

    it("missing getDimensionDataSourceString is ordinary fail after precheck", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installChartSeriesBubbleSizesExcel({ supportReadback: false });
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "B2:B3",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/getDimensionDataSourceString missing/);
        expect(result.capability).toBe("chart.series.bubbleSizes.update");
        expect(result.host).toBe("office-js");
        expect(result.reason).toMatch(/getDimensionDataSourceString/i);
      }
      expect(f.getSetterCallCounts().setBubbleSizesCalls).toBe(0);
      expect(f.getCommitted(0)?.bubbleSizesSource).toBeNull();
    });

    it("readback sync rejection fails without fake requirement-set evidence", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installChartSeriesBubbleSizesExcel({ failReadbackSync: true });
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "B2:B3",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.capability).toBe("chart.series.bubbleSizes.update");
        expect(result.host).toBe("office-js");
        expect(result.reason).toMatch(/readback sync rejected/);
      }
      expect(f.getSetterCallCounts().setBubbleSizesCalls).toBeGreaterThan(0);
      expect(f.getCommitted(0)?.bubbleSizesSource).toBeNull();
    });

    it("non-bubble host business error is not reclassified as ExcelApi 1.15 requirement-set", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartSeriesBubbleSizesExcel();
      const excel = (globalThis as unknown as { Excel: { run: Function } }).Excel;
      const originalRun = excel.run;
      excel.run = async <T>(fn: (ctx: unknown) => Promise<T>) => {
        void fn;
        throw new Error("setBubbleSizes is only valid for Bubble chart series");
      };
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "B2:B3",
      });
      excel.run = originalRun;
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.capability).toBe("chart.series.bubbleSizes.update");
        expect(result.host).toBe("office-js");
        expect(result.reason).toMatch(/only valid for Bubble/);
        expect(result.evidence).toBeUndefined();
      }
    });

    it("empty ClientResult source fails", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartSeriesBubbleSizesExcel({ readbackValue: "" });
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "B2:B3",
      });
      expect(result.ok).toBe(false);
    });

    it("non-string ClientResult source fails", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartSeriesBubbleSizesExcel({ readbackValue: 42 });
      const result = await new OfficeJsAdapter().updateChartSeriesBubbleSizes({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "B2:B3",
      });
      expect(result.ok).toBe(false);
    });
  });

  describe("range validation", () => {
    it("accepts same-sheet A1 and rejects cross-sheet for bubbleSizes", () => {
      expect(normalizeSameSheetSourceRange("Sheet1", "C2:C5")).toBe("C2:C5");
      expect(normalizeSameSheetSourceRange("Sheet1", "Sheet1!c2:c5")).toBe("C2:C5");
      expect(() => normalizeSameSheetSourceRange("Sheet1", "Other!C2:C5")).toThrow(
        /same worksheet/,
      );
    });

    it("host rejects cross-sheet bubbleSizesRange", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const result = await new ToolExecutor(host).execute({
        name: "chart.series.bubbleSizes.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          bubbleSizesRange: "Other!C2:C5",
        },
      });
      expect(result.ok).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity trim and persisted bubble source state", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const result = await new ToolExecutor(host).execute({
        name: "chart.series.bubbleSizes.update",
        arguments: {
          sheetName: "  Sheet1  ",
          chartName: "  C1  ",
          seriesIndex: 1,
          bubbleSizesRange: "  c2:c5  ",
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          bubbleSizesSource: "Sheet1!C2:C5",
          dataBound: true,
        });
      }
      const series = host.chartSeries.get("Sheet1\0C1");
      expect(series?.[0]?.bubbleSizesSource).toBe("Sheet1!C2:C5");
      expect(series?.[1]?.bubbleSizesSource).toBeNull();

      const second = await new ToolExecutor(host).execute({
        name: "chart.series.bubbleSizes.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          bubbleSizesRange: "D2:D5",
        },
      });
      expect(second.ok).toBe(true);
      if (second.ok) expect(second.data).toMatchObject({ bubbleSizesSource: "Sheet1!D2:D5" });
      expect(series?.[0]?.bubbleSizesSource).toBe("Sheet1!D2:D5");
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, bubbleSizesRange: "C2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1.5, bubbleSizesRange: "C2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: -1, bubbleSizesRange: "C2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: Number.NaN, bubbleSizesRange: "C2" },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: Number.POSITIVE_INFINITY,
          bubbleSizesRange: "C2",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: Number.NEGATIVE_INFINITY,
          bubbleSizesRange: "C2",
        },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, bubbleSizesRange: "" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, bubbleSizesRange: "  " },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, bubbleSizesRange: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, bubbleSizesRange: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, bubbleSizesRange: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, bubbleSizesRange: ["C2"] },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, bubbleSizesRange: "C2", extra: 1 },
        { sheetName: 1, chartName: "C1", seriesIndex: 1, bubbleSizesRange: "C2" },
        { sheetName: "Sheet1", chartName: false, seriesIndex: 1, bubbleSizesRange: "C2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: null, bubbleSizesRange: "C2" },
        { sheetName: "", chartName: "C1", seriesIndex: 1, bubbleSizesRange: "C2" },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          bubbleSizesRange: undefined,
        },
      ]) {
        const result = await executor.execute({
          name: "chart.series.bubbleSizes.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.bubbleSizes.update");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
      expect(
        (def?.parameters as { required?: string[] }).required,
      ).toEqual(["sheetName", "chartName", "seriesIndex", "bubbleSizesRange"]);
    });
  });

  it("WPS is typed unsupported with BubbleSizes evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.bubbleSizes.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        bubbleSizesRange: "C2:C5",
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.bubbleSizes.update");
      expect(detail.evidence).toMatch(/setBubbleSizes|BubbleSizes/);
    }
  });
});
