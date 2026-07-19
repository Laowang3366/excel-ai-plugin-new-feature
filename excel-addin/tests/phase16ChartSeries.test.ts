import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartSeriesExcel } from "./fakes/officeJsChartSeriesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase16 chart series", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesExcel>;

    beforeEach(() => {
      fake = installChartSeriesExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
    });

    it("list only sees series.items after load+sync", async () => {
      const adapter = new OfficeJsAdapter();
      expect(fake.getItemsVisible()).toBe(false);
      const listed = await adapter.listChartSeries("Sheet1", "C1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect(listed.data).toHaveLength(2);
        expect(listed.data[0]).toEqual({
          index: 1,
          name: "Series1",
          chartType: "column",
          smooth: false,
        });
        expect(listed.data[1]?.index).toBe(2);
        expect(listed.data[1]?.name).toBe("Series2");
      }
      expect(fake.getItemsVisible()).toBe(true);
    });

    it("update newName writes pending then committed readback", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        newName: "Renamed",
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data).toEqual({
          index: 1,
          name: "Renamed",
          chartType: "column",
          smooth: false,
        });
      }
      expect(fake.getCommitted(0)?.name).toBe("Renamed");
      expect(fake.getPending(0)).toBeUndefined();
    });

    it("update chartType alone with real readback", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        chartType: "line",
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.index).toBe(2);
        expect(updated.data.chartType).toBe("line");
        expect(updated.data.name).toBe("Series2");
      }
    });

    it("update smooth alone with real readback", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        smooth: true,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.smooth).toBe(true);
        expect(updated.data.name).toBe("Series1");
      }
    });

    it("combined update newName+chartType+smooth", async () => {
      const adapter = new OfficeJsAdapter();
      const updated = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        newName: "Combo",
        chartType: "area",
        smooth: true,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data).toEqual({
          index: 1,
          name: "Combo",
          chartType: "area",
          smooth: true,
        });
      }
      const listed = await adapter.listChartSeries("Sheet1", "C1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect(listed.data[0]).toEqual({
          index: 1,
          name: "Combo",
          chartType: "area",
          smooth: true,
        });
      }
    });

    it("seriesIndex out of range fails (not input echo)", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 99,
        newName: "X",
      });
      expect(result.ok).toBe(false);
    });

    it("unknown host chartType remains raw string on list", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      installChartSeriesExcel({
        series: [{ name: "R1", chartType: "Surface", smooth: false }],
      });
      const adapter = new OfficeJsAdapter();
      const listed = await adapter.listChartSeries("Sheet1", "C1");
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect(listed.data[0]?.chartType).toBe("Surface");
      }
    });

    it("skipping first sync after write yields old snapshot (two-sync order)", async () => {
      const broken = await fake.brokenUpdateSkipFirstSync("ShouldNotAppear");
      expect(broken.name).toBe("Series1");
      expect(fake.getCommitted(0)?.name).toBe("ShouldNotAppear");
      const adapter = new OfficeJsAdapter();
      const okPath = await adapter.updateChartSeries({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        newName: "AfterProperSync",
      });
      expect(okPath.ok).toBe(true);
      if (okPath.ok) expect(okPath.data.name).toBe("AfterProperSync");
    });
  });

  describe("executor + schema", () => {
    it("list and update via tools with MockHost parity", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      const listed = await executor.execute({
        name: "chart.series.list",
        arguments: { sheetName: "Sheet1", chartName: "C1" },
      });
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect((listed.data as { index: number }[]).map((s) => s.index)).toEqual([1, 2]);
      }

      const updated = await executor.execute({
        name: "chart.series.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          newName: "Alpha",
          chartType: "line",
          smooth: true,
        },
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data).toEqual({
          index: 1,
          name: "Alpha",
          chartType: "line",
          smooth: true,
        });
      }

      const again = await host.listChartSeries("Sheet1", "C1");
      expect(again.ok).toBe(true);
      if (again.ok) {
        expect(again.data[0]?.name).toBe("Alpha");
        expect(again.data[0]?.smooth).toBe(true);
      }
    });

    it("rejects null/unknown/empty/non-positive/missing update fields", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);

      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, newName: "A" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1.5, newName: "A" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: -1, newName: "A" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: null, newName: "A" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: undefined, newName: "A" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, newName: "" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, newName: "   " },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, newName: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, newName: undefined },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, chartType: "stock" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, chartType: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, chartType: undefined },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, smooth: null },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, smooth: undefined },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, smooth: "yes" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, newName: "A", formula: "x" },
        { sheetName: "", chartName: "C1", seriesIndex: 1, newName: "A" },
        { sheetName: "Sheet1", chartName: "", seriesIndex: 1, newName: "A" },
        { sheetName: null, chartName: "C1", seriesIndex: 1, newName: "A" },
        { sheetName: undefined, chartName: "C1", seriesIndex: 1, newName: "A" },
      ]) {
        const result = await executor.execute({
          name: "chart.series.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }

      const listBad = await executor.execute({
        name: "chart.series.list",
        arguments: { sheetName: "Sheet1", chartName: "C1", extra: true },
      });
      expect(listBad.ok).toBe(false);

      for (const args of [
        { sheetName: "", chartName: "C1" },
        { sheetName: "Sheet1", chartName: "" },
        { sheetName: null, chartName: "C1" },
        { sheetName: "Sheet1", chartName: null },
        { sheetName: undefined, chartName: "C1" },
        { sheetName: "Sheet1", chartName: undefined },
        { sheetName: "Sheet1" },
        { chartName: "C1" },
        {},
      ]) {
        const result = await executor.execute({
          name: "chart.series.list",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema has additionalProperties false for both tools", () => {
      for (const name of ["chart.series.list", "chart.series.update"] as const) {
        const def = TOOL_DEFINITIONS.find((t) => t.name === name);
        expect(def).toBeTruthy();
        expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
          false,
        );
      }
      const update = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.update");
      expect(update?.riskLevel).toBe("moderate");
      const list = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.list");
      expect(list?.riskLevel).toBe("safe");
    });
  });

  it("WPS chart.series.* are typed unsupported with ChartSeries evidence", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    for (const call of [
      {
        name: "chart.series.list" as const,
        arguments: { sheetName: "Sheet1", chartName: "C1" },
      },
      {
        name: "chart.series.update" as const,
        arguments: { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, newName: "X" },
      },
    ]) {
      const result = await executor.execute(call);
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        const detail = result.detail as { evidence?: string; capability?: string };
        expect(detail.evidence).toMatch(/ChartSeries/);
        expect(detail.capability).toBe(call.name);
      }
    }
  });
});
