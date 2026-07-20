import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import {
  CHART_DATA_LABEL_POSITIONS,
  type ChartDataLabelPosition,
} from "../shared/host/chartDataLabelsTypes";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { installChartDataLabelsExcel } from "./fakes/officeJsChartDataLabelsFake";
import { MockHostAdapter } from "./mockHost";

const POS_HOST: Record<ChartDataLabelPosition, string> = {
  none: "None",
  center: "Center",
  insideEnd: "InsideEnd",
  insideBase: "InsideBase",
  outsideEnd: "OutsideEnd",
  left: "Left",
  right: "Right",
  top: "Top",
  bottom: "Bottom",
  bestFit: "BestFit",
  callout: "Callout",
};

describe("phase52 chart.series.dataLabels extended fields", () => {
  describe("Office.js", () => {
    let fake: ReturnType<typeof installChartDataLabelsExcel>;
    beforeEach(() => {
      fake = installChartDataLabelsExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("round-trips new fields with old fields in one write", async () => {
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
        showValue: true,
        showCategoryName: true,
        showSeriesName: false,
        numberFormat: "0%",
        showPercentage: true,
        showBubbleSize: false,
        showLegendKey: true,
        separator: " | ",
        position: "outsideEnd",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          enabled: true,
          showValue: true,
          showCategoryName: true,
          showSeriesName: false,
          numberFormat: "0%",
          showPercentage: true,
          showBubbleSize: false,
          showLegendKey: true,
          separator: " | ",
          position: "outsideEnd",
        });
      }
      expect(fake.getCommitted(0)?.position).toBe("OutsideEnd");
      expect(fake.getCommitted(0)?.separator).toBe(" | ");
    });

    it("maps all positions to host and normalizes readback", async () => {
      for (const position of CHART_DATA_LABEL_POSITIONS) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        fake = installChartDataLabelsExcel();
        const result = await new OfficeJsAdapter().updateChartDataLabels({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          position,
        });
        expect(result.ok).toBe(true);
        if (result.ok) expect(result.data.position).toBe(position);
        expect(fake.getCommitted(0)?.position).toBe(POS_HOST[position]);
      }
    });

    it("preserves separator empty string and edge spaces", async () => {
      for (const separator of ["", "  x  ", "\t"]) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        fake = installChartDataLabelsExcel();
        const result = await new OfficeJsAdapter().updateChartDataLabels({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          separator,
        });
        expect(result.ok).toBe(true);
        if (result.ok) expect(result.data.separator).toBe(separator);
        expect(fake.getCommitted(0)?.separator).toBe(separator);
      }
    });

    it("enabled-only with 1.7 true 1.8 false succeeds without dataLabels access", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      fake = installChartDataLabelsExcel({ excelApi17: true, excelApi18: false });
      const result = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.enabled).toBe(true);
        expect(result.data.showPercentage).toBeUndefined();
        expect(result.data.position).toBeUndefined();
      }
      const counts = fake.getWriteCallCounts();
      expect(counts.hasDataLabelsWriteCalls).toBeGreaterThanOrEqual(1);
      expect(counts.dataLabelsWriteCalls).toBe(0);
    });

    it("new fields: 1.8 false/missing/throw typed unsupported, zero Excel.run", async () => {
      for (const opts of [
        { excelApi18: false },
        { missingIsSetSupported: true },
        { isSetSupportedThrows: true },
      ] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        fake = installChartDataLabelsExcel(opts);
        const result = await new OfficeJsAdapter().updateChartDataLabels({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          showPercentage: true,
        });
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).toBe(true);
          expect(result.reason).toMatch(/ExcelApi 1\.8|isSetSupported/i);
        }
        expect(fake.getWriteCallCounts().excelRunCalls).toBe(0);
        expect(fake.getWriteCallCounts().dataLabelsWriteCalls).toBe(0);
        expect(fake.getWriteCallCounts().worksheetGetItemCalls).toBe(0);
      }
    });

    it("1.8 true but extended members missing is ordinary failed with zero writes", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      fake = installChartDataLabelsExcel({ supportExtendedLabels: false });
      for (const input of [
        { showPercentage: true as const },
        { enabled: true as const, showPercentage: true as const },
      ]) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        fake = installChartDataLabelsExcel({ supportExtendedLabels: false });
        const result = await new OfficeJsAdapter().updateChartDataLabels({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          ...input,
        });
        expect(result.ok).toBe(false);
        if (!result.ok) {
          expect(result.unsupported).not.toBe(true);
          expect(result.reason).toMatch(/showPercentage|not available/i);
        }
        const counts = fake.getWriteCallCounts();
        expect(counts.dataLabelsWriteCalls).toBe(0);
        expect(counts.hasDataLabelsWriteCalls).toBe(0);
      }
    });

    it("bad/malformed position readback fails closed without echoing request", async () => {
      const adapter = new OfficeJsAdapter();
      for (const bad of ["Invalid", "Inside-End", "In side End"]) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        fake = installChartDataLabelsExcel();
        const ok = await adapter.updateChartDataLabels({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          position: "center",
        });
        expect(ok.ok).toBe(true);
        fake.poisonCommitted(0, { position: bad });
        const second = await adapter.updateChartDataLabels({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          showValue: true,
        });
        expect(second.ok).toBe(false);
        if (!second.ok) {
          expect(second.unsupported).not.toBe(true);
          expect(second.reason).toMatch(/position|unsupported host value/i);
          expect(second.reason).not.toMatch(/center/i);
        }
      }
    });

    it("skipping first sync after write yields stale snapshot (anti-stale contract)", async () => {
      const broken = await fake.brokenUpdateSkipFirstSync();
      expect(broken.enabled).toBe(false);
      expect(broken.showValue).toBe(false);
      const ok = await new OfficeJsAdapter().updateChartDataLabels({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        enabled: true,
        showPercentage: true,
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) {
        expect(ok.data.enabled).toBe(true);
        expect(ok.data.showPercentage).toBe(true);
      }
    });
  });

  describe("executor / schema / WPS", () => {
    it("rejects Invalid/unknown/null/undefined position and unknown fields before Host", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      let calls = 0;
      const original = host.updateChartDataLabels.bind(host);
      host.updateChartDataLabels = async (input) => {
        calls += 1;
        return original(input);
      };
      const ex = new ToolExecutor(host);
      for (const args of [
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          position: "Invalid",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          position: "diagonal",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          position: null,
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          position: undefined,
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          showPercentage: true,
          extra: 1,
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          separator: null,
        },
      ]) {
        const result = await ex.execute({
          name: "chart.series.dataLabels.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
      expect(calls).toBe(0);
    });

    it("rejects enabled=false with any new label field before Host", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      let calls = 0;
      const original = host.updateChartDataLabels.bind(host);
      host.updateChartDataLabels = async (input) => {
        calls += 1;
        return original(input);
      };
      const ex = new ToolExecutor(host);
      for (const field of [
        { showPercentage: true },
        { showBubbleSize: true },
        { showLegendKey: true },
        { separator: "" },
        { position: "left" },
      ]) {
        const result = await ex.execute({
          name: "chart.series.dataLabels.update",
          arguments: {
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 1,
            enabled: false,
            ...field,
          },
        });
        expect(result.ok).toBe(false);
      }
      expect(calls).toBe(0);
    });

    it("mock host accepts extended fields; tool count 98; WPS unsupported", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const ok = await new ToolExecutor(host).execute({
        name: "chart.series.dataLabels.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          showPercentage: true,
          separator: "",
          position: "callout",
        },
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) {
        expect(ok.data).toMatchObject({
          showPercentage: true,
          separator: "",
          position: "callout",
        });
      }

      const wps = await new ToolExecutor(new WpsJsaAdapter()).execute({
        name: "chart.series.dataLabels.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          showPercentage: true,
        },
      });
      expect(wps.ok).toBe(false);
      if (!wps.ok) expect(wps.unsupported).toBe(true);
      expect(TOOL_DEFINITIONS).toHaveLength(98);
    });
  });
});
