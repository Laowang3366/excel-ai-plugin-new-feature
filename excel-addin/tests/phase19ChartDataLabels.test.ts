import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartDataLabelsExcel } from "./fakes/officeJsChartDataLabelsFake";
import { MockHostAdapter } from "./mockHost";

describe("phase19 chart series dataLabels", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartDataLabelsExcel>;
    beforeEach(() => {
      fake = installChartDataLabelsExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("updates showValue alone with real readback", async () => {
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        showValue: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          enabled: false,
          showValue: true,
          showCategoryName: false,
          showSeriesName: false,
          numberFormat: "General",
        });
      }
      expect(fake.getCommitted(0)?.showValue).toBe(true);
    });

    it("combined update all fields", async () => {
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        showValue: true,
        showCategoryName: true,
        showSeriesName: false,
        numberFormat: "0.0%",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          seriesIndex: 2,
          showValue: true,
          showCategoryName: true,
          showSeriesName: false,
          numberFormat: "0.0%",
        });
      }
    });

    it("skipping first sync after write yields old snapshot", async () => {
      const broken = await fake.brokenUpdateSkipFirstSync();
      expect(broken.showValue).toBe(false);
      expect(broken.enabled).toBe(false);
      expect(fake.getCommitted(0)?.showValue).toBe(true);

      const ok = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        showCategoryName: true,
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) {
        expect(ok.data.showValue).toBe(true);
        expect(ok.data.showCategoryName).toBe(true);
      }
    });

    it("seriesIndex out of range fails", async () => {
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 99,
        showValue: true,
      });
      expect(result.ok).toBe(false);
    });

    it("rejects null/undefined loaded fields and non-string chart.name", async () => {
      fake.poisonCommitted(0, { showValue: null });
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

      const f2 = installChartDataLabelsExcel();
      f2.poisonCommitted(0, { showValue: undefined });
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
      f3.poisonCommitted(0, { numberFormat: null });
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

      const f4 = installChartDataLabelsExcel();
      f4.setChartName(42);
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
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity for each field", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      const v = await executor.execute({
        name: "chart.series.dataLabels.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          showValue: true,
        },
      });
      expect(v.ok).toBe(true);
      if (v.ok) {
        expect(v.data).toMatchObject({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          showValue: true,
        });
      }

      const fmt = await executor.execute({
        name: "chart.series.dataLabels.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          numberFormat: "0",
          showSeriesName: true,
        },
      });
      expect(fmt.ok).toBe(true);
      if (fmt.ok) {
        expect(fmt.data).toMatchObject({
          numberFormat: "0",
          showSeriesName: true,
          showValue: true,
        });
      }
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, showValue: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1.5, showValue: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: null, showValue: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: undefined, showValue: true },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, showValue: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, numberFormat: "" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, numberFormat: "   " },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, numberFormat: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, showValue: true, extra: 1 },
        { sheetName: "", chartName: "C1", seriesIndex: 1, showValue: true },
      ]) {
        const result = await executor.execute({
          name: "chart.series.dataLabels.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and moderate risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.dataLabels.update");
      expect(def?.riskLevel).toBe("moderate");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS dataLabels update is typed unsupported with ChartDataLabels evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.dataLabels.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        showValue: true,
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.dataLabels.update");
      expect(detail.evidence).toMatch(/ChartDataLabels|hasDataLabels|dataLabels\.enabled/);
    }
  });
});
