import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { normalizeSameSheetSourceRange } from "../shared/host/officeJsChartSource";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSeriesValuesExcel } from "./fakes/officeJsChartSeriesValuesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase23 chart series values", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesValuesExcel>;
    beforeEach(() => {
      fake = installChartSeriesValuesExcel({ hostSourcePrefix: "HostSheet!" });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("binds values-only with host source string readback (not input echo)", async () => {
      fake.setLoadedChartName("HostChart");
      const result = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        valuesRange: "B2:B5",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.chartName).toBe("HostChart");
        expect(result.data.chartName).not.toBe("C1");
        expect(result.data.valuesSource).toBe("HostSheet!B2:B5");
        expect(result.data.valuesSource).not.toBe("B2:B5");
        expect(result.data.dataBound).toBe(true);
      }
      expect(fake.getCommitted(0)?.valuesSource).toBe("HostSheet!B2:B5");
      expect(fake.getCommitted(1)?.valuesSource).toBeNull();
    });

    it("seriesIndex 2 writes only second series; first unchanged", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        valuesRange: "Z2:Z4",
        xValuesRange: "Y2:Y4",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.seriesIndex).toBe(2);
        expect(result.data.valuesSource).toBe("HostSheet!Z2:Z4");
        expect(result.data.xValuesSource).toBe("HostSheet!Y2:Y4");
        expect(result.data.xValuesSource).not.toBe("Y2:Y4");
      }
      expect(fake.getCommitted(0)?.valuesSource).toBeNull();
      expect(fake.getCommitted(0)?.xValuesSource).toBeNull();
      expect(fake.getCommitted(1)?.valuesSource).toBe("HostSheet!Z2:Z4");
      expect(fake.getCommitted(1)?.xValuesSource).toBe("HostSheet!Y2:Y4");
    });

    it("binds xValues-only and dual dimensions", async () => {
      const xOnly = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        xValuesRange: "A2:A5",
      });
      expect(xOnly.ok).toBe(true);
      if (xOnly.ok) {
        expect(xOnly.data.xValuesSource).toBe("HostSheet!A2:A5");
        expect(xOnly.data.valuesSource).toBeUndefined();
      }

      const both = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        valuesRange: "C2:C4",
        xValuesRange: "D2:D4",
      });
      expect(both.ok).toBe(true);
      if (both.ok) {
        expect(both.data.valuesSource).toBe("HostSheet!C2:C4");
        expect(both.data.xValuesSource).toBe("HostSheet!D2:D4");
      }
    });

    it("skipping first sync cannot read new source string", async () => {
      const broken = await fake.brokenSkipFirstSync(1, "E2:E3");
      expect(broken).toBeNull();
      expect(fake.getCommitted(0)?.valuesSource).toBe("HostSheet!E2:E3");

      const ok = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        valuesRange: "F2:F3",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.valuesSource).toBe("HostSheet!F2:F3");
    });

    it("ExcelApi 1.15 precheck false is unsupported and never calls setters", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installChartSeriesValuesExcel({ excelApi115: false });
      const result = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        valuesRange: "B2:B3",
        xValuesRange: "A2:A3",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.evidence).toMatch(/ExcelApi 1\.15/);
        expect(result.reason).toMatch(/isSetSupported|1\.15/);
      }
      expect(f.getSetterCallCounts()).toEqual({ setValuesCalls: 0, setXAxisValuesCalls: 0 });
      expect(f.getCommitted(0)?.valuesSource).toBeNull();
      expect(f.getCommitted(0)?.xValuesSource).toBeNull();
    });

    it("missing getDimensionDataSourceString is ordinary fail after precheck", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installChartSeriesValuesExcel({ supportReadback: false });
      const result = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        valuesRange: "B2:B3",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.capability).toBe("chart.series.values.update");
        expect(result.host).toBe("office-js");
        expect(result.reason).toMatch(/getDimensionDataSourceString/i);
      }
      expect(f.getCommitted(0)?.valuesSource).toBeNull();
    });

    it("readback sync rejection is ordinary fail and leaves no unverified write", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installChartSeriesValuesExcel({ failReadbackSync: true });
      const result = await new OfficeJsAdapter().updateChartSeriesValues({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        valuesRange: "B2:B3",
        xValuesRange: "A2:A3",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.capability).toBe("chart.series.values.update");
        expect(result.host).toBe("office-js");
        expect(result.reason).toMatch(/getDimensionDataSourceString|sync|reject/i);
      }
      expect(f.getSetterCallCounts().setValuesCalls).toBeGreaterThan(0);
      expect(f.getCommitted(0)?.valuesSource).toBeNull();
      expect(f.getCommitted(0)?.xValuesSource).toBeNull();
    });
  });

  describe("range validation", () => {
    it("accepts same-sheet A1 and rejects cross-sheet for values/xValues", () => {
      expect(normalizeSameSheetSourceRange("Sheet1", "B2:B5")).toBe("B2:B5");
      expect(normalizeSameSheetSourceRange("Sheet1", "Sheet1!a2:a5")).toBe("A2:A5");
      expect(() => normalizeSameSheetSourceRange("Sheet1", "Other!B2:B5")).toThrow(
        /same worksheet/,
      );
      expect(() => normalizeSameSheetSourceRange("Sheet1", "Other!A2:A5")).toThrow(
        /same worksheet/,
      );
    });

    it("host rejects cross-sheet valuesRange/xValuesRange", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const valuesCross = await executor.execute({
        name: "chart.series.values.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          valuesRange: "Other!B2:B5",
        },
      });
      expect(valuesCross.ok).toBe(false);

      const xCross = await executor.execute({
        name: "chart.series.values.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          xValuesRange: "Other!A2:A5",
        },
      });
      expect(xCross.ok).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity trim and dual ranges", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const result = await new ToolExecutor(host).execute({
        name: "chart.series.values.update",
        arguments: {
          sheetName: "  Sheet1  ",
          chartName: "  C1  ",
          seriesIndex: 1,
          valuesRange: "  b2:b5  ",
          xValuesRange: " a2:a5 ",
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          valuesSource: "Sheet1!B2:B5",
          xValuesSource: "Sheet1!A2:A5",
          dataBound: true,
        });
      }
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, valuesRange: "B2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1.5, valuesRange: "B2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, valuesRange: "" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, valuesRange: "  " },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, valuesRange: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, valuesRange: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, valuesRange: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, valuesRange: ["B2"] },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, valuesRange: "B2", extra: 1 },
        { sheetName: 1, chartName: "C1", seriesIndex: 1, valuesRange: "B2" },
        { sheetName: "Sheet1", chartName: false, seriesIndex: 1, valuesRange: "B2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: null, valuesRange: "B2" },
        { sheetName: "", chartName: "C1", seriesIndex: 1, valuesRange: "B2" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, xValuesRange: undefined },
      ]) {
        const result = await executor.execute({
          name: "chart.series.values.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.values.update");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS is typed unsupported with values/xValues evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.values.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        valuesRange: "B2:B5",
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.values.update");
      expect(detail.evidence).toMatch(/values\/xValues/);
    }
  });
});
