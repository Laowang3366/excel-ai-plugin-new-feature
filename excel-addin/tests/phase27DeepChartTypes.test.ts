import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import type { ChartType } from "../shared/host/types";
import { mapChartType, toChartTypeLabel } from "../shared/host/officeJsChartTypes";
import { CHART_TOOL_DEFINITIONS } from "../shared/tools/chartDefinitions";
import { ToolExecutor } from "../shared/tools";
import { installChartSeriesExcel } from "./fakes/officeJsChartSeriesFake";
import { installObjectUpdateExcel } from "./fakes/officeJsObjectUpdateFake";
import { installSyncGatedExcel } from "./fakes/officeJsSyncGated";
import { MockHostAdapter } from "./mockHost";

const DEEP: ChartType[] = ["doughnut", "bubble", "radar", "linemarkers"];
const BASIC: ChartType[] = ["column", "line", "bar", "area", "pie", "scatter"];

const DEEP_ENUM: Record<string, string> = {
  doughnut: "Doughnut",
  bubble: "Bubble",
  radar: "Radar",
  linemarkers: "LineMarkers",
};

describe("phase27 deep chart types", () => {
  describe("map / reverse-map", () => {
    it("maps four deep types to official Office.js enums", () => {
      for (const type of DEEP) {
        expect(mapChartType(type)).toBe(DEEP_ENUM[type]);
      }
      expect(mapChartType(undefined)).toBe("ColumnClustered");
    });

    it("reverse-maps host strings; LineMarkers≠line, Doughnut≠pie", () => {
      expect(toChartTypeLabel("Doughnut")).toBe("doughnut");
      expect(toChartTypeLabel("Bubble")).toBe("bubble");
      expect(toChartTypeLabel("Radar")).toBe("radar");
      expect(toChartTypeLabel("LineMarkers")).toBe("linemarkers");
      expect(toChartTypeLabel("LineMarkers")).not.toBe("line");
      expect(toChartTypeLabel("Doughnut")).not.toBe("pie");
      expect(toChartTypeLabel("Line")).toBe("line");
      expect(toChartTypeLabel("Pie")).toBe("pie");
      expect(toChartTypeLabel("Surface")).toBe("Surface");
      expect(toChartTypeLabel("StockHLC")).toBe("StockHLC");
    });

    it("deep types require exact official names; unknown raw stays raw", () => {
      expect(toChartTypeLabel("Dough-nut")).toBe("Dough-nut");
      expect(toChartTypeLabel("doughnut")).toBe("doughnut");
      expect(toChartTypeLabel("Line Markers")).toBe("Line Markers");
      expect(toChartTypeLabel("linemarkers")).toBe("linemarkers");
      expect(toChartTypeLabel("bubble")).toBe("bubble");
      expect(toChartTypeLabel("Radar ")).toBe("Radar ");
      expect(toChartTypeLabel("LineMarkersStacked")).toBe("line");
    });

    it("basic six still map", () => {
      expect(mapChartType("column")).toBe("ColumnClustered");
      expect(mapChartType("line")).toBe("Line");
      expect(mapChartType("bar")).toBe("BarClustered");
      expect(mapChartType("area")).toBe("Area");
      expect(mapChartType("pie")).toBe("Pie");
      expect(mapChartType("scatter")).toBe("XYScatter");
    });
  });

  describe("chart.create + list host readback", () => {
    beforeEach(() => {
      installSyncGatedExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it.each(DEEP)("create %s writes enum and readbacks label", async (chartType) => {
      const adapter = new OfficeJsAdapter();
      const created = await adapter.createChart({
        sheetName: "Sheet1",
        sourceRange: "A1:B2",
        chartType,
        name: `Deep_${chartType}`,
      });
      expect(created.ok).toBe(true);
      if (created.ok) expect(created.data.chartType).toBe(chartType);

      const listed = await adapter.listCharts("Sheet1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        const row = listed.data.find((c) => c.name === `Deep_${chartType}`);
        expect(row?.chartType).toBe(chartType);
      }
    });
  });

  describe("chart.update host readback", () => {
    beforeEach(() => {
      installObjectUpdateExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it.each(DEEP)("update to %s readbacks label", async (chartType) => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChart({
        sheetName: "Sheet1",
        chartName: "C1",
        chartType,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) expect(updated.data.chartType).toBe(chartType);
    });
  });

  describe("chart.series.update host readback", () => {
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it.each(DEEP)("seriesIndex 1-based update %s", async (chartType) => {
      installChartSeriesExcel();
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        chartType,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.index).toBe(1);
        expect(updated.data.chartType).toBe(chartType);
      }

      const second = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        chartType,
      });
      expect(second.ok).toBe(true);
      if (second.ok) {
        expect(second.data.index).toBe(2);
        expect(second.data.chartType).toBe(chartType);
      }
    });
  });

  describe("schema / executor strictness", () => {
    it("accepts ten types; rejects unknown/null/undefined/extra", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      for (const chartType of [...BASIC, ...DEEP]) {
        const okCreate = await executor.execute({
          name: "chart.create",
          arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", chartType },
        });
        expect(okCreate.ok).toBe(true);
      }

      const defaultCol = await executor.execute({
        name: "chart.create",
        arguments: { sheetName: "Sheet1", sourceRange: "A1:B2" },
      });
      expect(defaultCol.ok).toBe(true);
      if (defaultCol.ok) {
        expect((defaultCol.data as { chartType: string }).chartType).toBe("column");
      }

      await executor.execute({
        name: "chart.create",
        arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" },
      });

      for (const args of [
        { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "stock" },
        { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "funnel" },
        { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: null },
        { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: undefined },
        { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "" },
        { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "Doughnut" },
        { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "bubble", extra: true },
        { sheetName: "Sheet1", chartName: "C1", chartType: "radar", extra: true },
      ]) {
        const tool = "chartName" in args ? "chart.update" : "chart.create";
        const bad = await executor.execute({
          name: tool,
          arguments: args as Record<string, unknown>,
        });
        expect(bad.ok).toBe(false);
      }

      for (const badUpdate of [
        { sheetName: "Sheet1", chartName: "C1", title: "T", chartType: null },
        { sheetName: "Sheet1", chartName: "C1", title: "T", chartType: undefined },
        { sheetName: "Sheet1", chartName: "C1", title: "T", chartType: "" },
        { sheetName: "Sheet1", chartName: "C1", title: "T", chartType: "stock" },
        { sheetName: "Sheet1", chartName: "C1", title: "T", chartType: 1 },
      ]) {
        const r = await executor.execute({
          name: "chart.update",
          arguments: badUpdate as Record<string, unknown>,
        });
        expect(r.ok).toBe(false);
      }

      for (const chartType of DEEP) {
        const seriesOk = await executor.execute({
          name: "chart.series.update",
          arguments: {
            sheetName: "Sheet1",
            chartName: "C1",
            seriesIndex: 1,
            chartType,
          },
        });
        expect(seriesOk.ok).toBe(true);
      }

      for (const seriesArgs of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, chartType: "stock" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, chartType: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, chartType: undefined },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          chartType: "line",
          formula: "x",
        },
      ]) {
        const seriesBad = await executor.execute({
          name: "chart.series.update",
          arguments: seriesArgs as Record<string, unknown>,
        });
        expect(seriesBad.ok).toBe(false);
      }
    });

    it("chart.create schema has additionalProperties false", () => {
      const def = CHART_TOOL_DEFINITIONS.find((t) => t.name === "chart.create");
      expect(
        (def?.parameters as { additionalProperties?: boolean }).additionalProperties,
      ).toBe(false);
    });
  });

  describe("create/update skip-first-sync fence", () => {
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("create skip-first-sync reads stale column not deep type", async () => {
      const fake = installSyncGatedExcel();
      const stale = await fake.brokenCreateSkipFirstSync("Doughnut");
      expect(stale.chartType).toBe("ColumnClustered");
    });

    it("update skip-first-sync keeps old chartType", async () => {
      const fake = installObjectUpdateExcel();
      const stale = await fake.brokenUpdateSkipFirstSync("Bubble");
      expect(stale.chartType).toBe("ColumnClustered");
      expect(stale.title).toBe("Old");
    });
  });

  describe("WPS typed unsupported", () => {
    it("create/update/series deep types stay unsupported on WPS path", async () => {
      const { WpsJsaAdapter } = await import("../shared/host/wpsJsaAdapter");
      const host = new WpsJsaAdapter();
      for (const chartType of DEEP) {
        const created = await host.createChart({
          sheetName: "Sheet1",
          sourceRange: "A1:B2",
          chartType,
        });
        expect(created.ok).toBe(false);
        if (!created.ok) expect(created.unsupported).toBe(true);

        const updated = await host.updateChart({
          sheetName: "Sheet1",
          chartName: "C1",
          chartType,
        });
        expect(updated.ok).toBe(false);
        if (!updated.ok) expect(updated.unsupported).toBe(true);

        const series = await host.updateChartSeries({
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          chartType,
        });
        expect(series.ok).toBe(false);
        if (!series.ok) expect(series.unsupported).toBe(true);
      }
    });
  });
});
