import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartDataLabelsExcel } from "./fakes/officeJsChartDataLabelsFake";
import { MockHostAdapter } from "./mockHost";

describe("phase26 chart series dataLabels.enabled", () => {
  describe("Office.js sync-gated hasDataLabels", () => {
    let fake: ReturnType<typeof installChartDataLabelsExcel>;
    beforeEach(() => {
      fake = installChartDataLabelsExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("enabled true alone with real host readback (1.7 path, no show* fields)", async () => {
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          enabled: true,
        });
        expect(result.data.showValue).toBeUndefined();
        expect(result.data.numberFormat).toBeUndefined();
      }
      expect(fake.getCommitted(0)?.enabled).toBe(true);
      expect(fake.getWriteCallCounts().dataLabelsWriteCalls).toBe(0);
    });

    it("enabled false alone with real host readback", async () => {
      await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: false,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.enabled).toBe(false);
      }
      expect(fake.getCommitted(0)?.enabled).toBe(false);
    });

    it("enabled true combined with show fields on seriesIndex 2", async () => {
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        enabled: true,
        showValue: true,
        showCategoryName: true,
        showSeriesName: false,
        numberFormat: "0.00",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          seriesIndex: 2,
          enabled: true,
          showValue: true,
          showCategoryName: true,
          showSeriesName: false,
          numberFormat: "0.00",
        });
      }
      expect(fake.getCommitted(1)?.enabled).toBe(true);
      expect(fake.getCommitted(1)?.showValue).toBe(true);
    });

    it("skipping first sync after write yields old snapshot", async () => {
      const broken = await fake.brokenUpdateSkipFirstSync();
      expect(broken.enabled).toBe(false);
      expect(broken.showValue).toBe(false);
      expect(fake.getCommitted(0)?.enabled).toBe(true);
      expect(fake.getCommitted(0)?.showValue).toBe(true);

      const ok = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        showCategoryName: true,
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) {
        expect(ok.data.enabled).toBe(true);
        expect(ok.data.showValue).toBe(true);
        expect(ok.data.showCategoryName).toBe(true);
      }
    });

    it("ExcelApi 1.7 precheck false: hasDataLabels writes 0 and no workbook locate", async () => {
      const f = installChartDataLabelsExcel({ excelApi17: false, excelApi18: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.7/);
        expect(result.evidence).toMatch(/hasDataLabels.*1\.7|enabled-only/i);
      }
      const counts = f.getWriteCallCounts();
      expect(counts.hasDataLabelsWriteCalls).toBe(0);
      expect(counts.dataLabelsWriteCalls).toBe(0);
      expect(counts.excelRunCalls).toBe(0);
      expect(counts.worksheetGetItemCalls).toBe(0);
    });

    it("ExcelApi 1.7=true 1.8=false: enabled-only succeeds without dataLabels touch", async () => {
      const f = installChartDataLabelsExcel({ excelApi17: true, excelApi18: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          enabled: true,
        });
      }
      expect(f.getWriteCallCounts().hasDataLabelsWriteCalls).toBe(1);
      expect(f.getWriteCallCounts().dataLabelsWriteCalls).toBe(0);
    });

    it("ExcelApi 1.8 precheck false for dataLabels fields: writes 0 and no workbook locate", async () => {
      const f = installChartDataLabelsExcel({ excelApi17: true, excelApi18: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        showValue: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.8/);
        expect(result.evidence).toMatch(/1\.8/);
      }
      const counts = f.getWriteCallCounts();
      expect(counts.hasDataLabelsWriteCalls).toBe(0);
      expect(counts.dataLabelsWriteCalls).toBe(0);
      expect(counts.excelRunCalls).toBe(0);
      expect(counts.worksheetGetItemCalls).toBe(0);
    });

    it("ExcelApi 1.8 precheck false for enabled+show combo: all writes 0", async () => {
      const f = installChartDataLabelsExcel({ excelApi17: true, excelApi18: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
        showValue: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.8/);
      }
      const counts = f.getWriteCallCounts();
      expect(counts.hasDataLabelsWriteCalls).toBe(0);
      expect(counts.dataLabelsWriteCalls).toBe(0);
      expect(counts.excelRunCalls).toBe(0);
      expect(counts.worksheetGetItemCalls).toBe(0);
    });

    it("missing hasDataLabels fails without requirement-set evidence", async () => {
      installChartDataLabelsExcel({ supportHasDataLabels: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.reason).toMatch(/hasDataLabels missing/i);
        expect(result.evidence ?? "").not.toMatch(/ExcelApi 1\.7/);
        expect(result.reason ?? "").not.toMatch(/isSetSupported/);
      }
    });

    it("enabled-only succeeds when dataLabels property absent (1.7 path)", async () => {
      const f = installChartDataLabelsExcel({ supportDataLabels: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.data.enabled).toBe(true);
      expect(f.getWriteCallCounts().dataLabelsWriteCalls).toBe(0);
    });

    it("missing dataLabels on 1.8 path fails without requirement-set evidence", async () => {
      installChartDataLabelsExcel({ supportDataLabels: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        showValue: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.reason).toMatch(/dataLabels missing/i);
        expect(result.evidence ?? "").not.toMatch(/ExcelApi 1\.8/);
      }
    });

    it("rejects null/undefined/non-boolean enabled readback and non-string names", async () => {
      fake.poisonCommitted(0, { enabled: null as unknown as boolean });
      expect(
        (
          await new OfficeJsAdapter().updateChartDataLabels({
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 1,
            showValue: true,
          })
        ).ok,
      ).toBe(false);

      const f2 = installChartDataLabelsExcel();
      f2.poisonCommitted(0, { showValue: null as unknown as boolean });
      expect(
        (
          await new OfficeJsAdapter().updateChartDataLabels({
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 1,
            showCategoryName: true,
          })
        ).ok,
      ).toBe(false);

      const f3 = installChartDataLabelsExcel();
      f3.setChartName(42);
      expect(
        (
          await new OfficeJsAdapter().updateChartDataLabels({
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 1,
            enabled: true,
          })
        ).ok,
      ).toBe(false);

      const f4 = installChartDataLabelsExcel();
      f4.setSheetName(null);
      expect(
        (
          await new OfficeJsAdapter().updateChartDataLabels({
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 1,
            enabled: true,
          })
        ).ok,
      ).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity for enabled true/false and combined", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      const on = await executor.execute({
        name: "chart.series.dataLabels.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          enabled: true,
        },
      });
      expect(on.ok).toBe(true);
      if (on.ok) {
        expect(on.data).toMatchObject({ enabled: true, seriesIndex: 1 });
      }

      const combo = await executor.execute({
        name: "chart.series.dataLabels.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          enabled: true,
          showValue: true,
          numberFormat: "0",
        },
      });
      expect(combo.ok).toBe(true);
      if (combo.ok) {
        expect(combo.data).toMatchObject({
          enabled: true,
          showValue: true,
          numberFormat: "0",
        });
      }

      const off = await executor.execute({
        name: "chart.series.dataLabels.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          enabled: false,
        },
      });
      expect(off.ok).toBe(true);
      if (off.ok) {
        expect(off.data).toMatchObject({ enabled: false });
      }
    });

    it("rejects invalid args including enabled=false with other fields", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1.5, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: NaN, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: Infinity, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: -Infinity, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: -1, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: null, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: undefined, enabled: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, enabled: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, enabled: "true" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, enabled: false, showValue: true },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          enabled: false,
          numberFormat: "0",
        },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, enabled: true, extra: 1 },
        { sheetName: "", chartName: "C1", seriesIndex: 1, enabled: true },
      ]) {
        const result = await executor.execute({
          name: "chart.series.dataLabels.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema includes enabled and additionalProperties false", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.dataLabels.update");
      expect(def?.riskLevel).toBe("moderate");
      const params = def?.parameters as {
        additionalProperties?: boolean;
        properties?: Record<string, unknown>;
      };
      expect(params.additionalProperties).toBe(false);
      expect(params.properties?.enabled).toEqual({ type: "boolean" });
    });
  });

  it("WPS dataLabels update is typed unsupported with hasDataLabels evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.dataLabels.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.dataLabels.update");
      expect(detail.evidence).toMatch(/hasDataLabels|dataLabels\.enabled/);
    }
  });
});
