import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import type { ChartType } from "../shared/host/types";
import { mapChartType, toChartTypeLabel } from "../shared/host/officeJsChartTypes";
import { ToolExecutor } from "../shared/tools";
import { installSyncGatedExcel } from "./fakes/officeJsSyncGated";
import { installObjectUpdateExcel } from "./fakes/officeJsObjectUpdateFake";
import { MockHostAdapter } from "./mockHost";

const BASIC: ChartType[] = ["column", "line", "bar", "area", "pie", "scatter"];

const OFFICE_ENUM: Record<string, string> = {
  column: "ColumnClustered",
  line: "Line",
  bar: "BarClustered",
  area: "Area",
  pie: "Pie",
  scatter: "XYScatter",
};

describe("phase13 chart types", () => {
  describe("map / label helpers", () => {
    it("maps six contract types to Office.js enums", () => {
      for (const type of BASIC) {
        expect(mapChartType(type)).toBe(OFFICE_ENUM[type]);
      }
      expect(mapChartType(undefined)).toBe("ColumnClustered");
    });

    it("labels only known Office enums; unknowns stay raw", () => {
      expect(toChartTypeLabel("ColumnClustered")).toBe("column");
      expect(toChartTypeLabel("Line")).toBe("line");
      expect(toChartTypeLabel("BarClustered")).toBe("bar");
      expect(toChartTypeLabel("Area")).toBe("area");
      expect(toChartTypeLabel("Pie")).toBe("pie");
      expect(toChartTypeLabel("XYScatter")).toBe("scatter");
      expect(toChartTypeLabel("XYScatterLines")).toBe("scatter");
      // not in known map → raw (no includes/prefix guessing)
      expect(toChartTypeLabel("CustomBar")).toBe("CustomBar");
      expect(toChartTypeLabel("BarCustom")).toBe("BarCustom");
      expect(toChartTypeLabel("AreaWhatever")).toBe("AreaWhatever");
      expect(toChartTypeLabel("AreaCustom")).toBe("AreaCustom");
      expect(toChartTypeLabel("Surface")).toBe("Surface");
    });
  });

  describe("Office.js create + list", () => {
    beforeEach(() => {
      installSyncGatedExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it.each(BASIC)("create %s uses Office enum and readbacks label", async (chartType) => {
      const adapter = new OfficeJsAdapter();
      const created = await adapter.createChart({
        sheetName: "Sheet1",
        sourceRange: "A1:B2",
        chartType,
        name: `C_${chartType}`,
      });
      expect(created.ok).toBe(true);
      if (created.ok) {
        expect(created.data.chartType).toBe(chartType);
      }

      const listed = await adapter.listCharts("Sheet1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        const row = listed.data.find((c) => c.name === `C_${chartType}`);
        expect(row?.chartType).toBe(chartType);
      }
    });

    it("defaults omitted chartType to column", async () => {
      const adapter = new OfficeJsAdapter();
      const created = await adapter.createChart({
        sheetName: "Sheet1",
        sourceRange: "A1:B2",
        name: "DefaultCol",
      });
      expect(created.ok).toBe(true);
      if (created.ok) expect(created.data.chartType).toBe("column");
    });
  });

  describe("Office.js update chartType", () => {
    beforeEach(() => {
      installObjectUpdateExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it.each(BASIC)("update to %s readbacks label", async (chartType) => {
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

  describe("executor", () => {
    it("accepts six types and rejects unknown", async () => {
      const executor = new ToolExecutor(new MockHostAdapter());
      for (const chartType of BASIC) {
        const result = await executor.execute({
          name: "chart.create",
          arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", chartType },
        });
        expect(result.ok).toBe(true);
      }
      const bad = await executor.execute({
        name: "chart.create",
        arguments: { sheetName: "Sheet1", sourceRange: "A1:B2", chartType: "stock" },
      });
      expect(bad.ok).toBe(false);
    });
  });
});
