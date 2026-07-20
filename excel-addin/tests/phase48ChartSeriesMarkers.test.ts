import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { installChartSeriesMarkersExcel } from "./fakes/officeJsChartSeriesMarkersFake";
import { MockHostAdapter } from "./mockHost";

describe("phase48 chart series markers", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesMarkersExcel>;
    beforeEach(() => {
      fake = installChartSeriesMarkersExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("updates markers with host readback (not input echo)", async () => {
      fake.setChartName("HostChart");
      const result = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerStyle: "diamond",
        markerSize: 12,
        markerBackgroundColor: "#FF0000",
        markerForegroundColor: "#00FF00",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.chartName).toBe("HostChart");
        expect(result.data.chartName).not.toBe("C1");
        expect(result.data.markerStyle).toBe("diamond");
        expect(result.data.markerSize).toBe(12);
        expect(result.data.markerBackgroundColor).toBe("#FF0000");
        expect(result.data.markerForegroundColor).toBe("#00FF00");
      }
      expect(fake.getCommitted(0)?.markerStyle).toBe("Diamond");
      expect(fake.getCommitted(0)?.markerSize).toBe(12);
      expect(fake.getCommitted(1)?.markerStyle).toBe("Automatic");
    });

    it("seriesIndex 2 writes only second series", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 2,
        markerStyle: "circle",
        markerSize: 9,
      });
      expect(result.ok).toBe(true);
      expect(fake.getCommitted(0)?.markerStyle).toBe("Automatic");
      expect(fake.getCommitted(1)?.markerStyle).toBe("Circle");
      expect(fake.getCommitted(1)?.markerSize).toBe(9);
    });


    it("normalizes lowercase and bare hex colors on direct adapter path", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerBackgroundColor: "#ff00aa",
        markerForegroundColor: "00bbcc",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.markerBackgroundColor).toBe("#FF00AA");
        expect(result.data.markerForegroundColor).toBe("#00BBCC");
      }
      expect(fake.getCommitted(0)?.markerBackgroundColor).toBe("#FF00AA");
      expect(fake.getCommitted(0)?.markerForegroundColor).toBe("#00BBCC");
    });

    it("accepts picture markerStyle and maps host Picture readback", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerStyle: "picture",
        markerSize: 20,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.markerStyle).toBe("picture");
        expect(result.data.markerSize).toBe(20);
      }
      expect(fake.getCommitted(0)?.markerStyle).toBe("Picture");
    });

    it("tolerates host color case/alpha form on readback after normalized write", async () => {
      // Write upper-normalized colors, then host returns lowercase/no-hash until load snapshot —
      // poison committed after write-sync would break same-run; use coerce by writing then
      // re-read path via second update that only sets size while colors remain mixed case.
      const first = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerBackgroundColor: "#AABBCC",
        markerForegroundColor: "#112233",
      });
      expect(first.ok).toBe(true);
      // Host may store mixed case; poison to lowercase without # to simulate Office readback shapes.
      fake.poison(0, {
        markerBackgroundColor: "aabbcc",
        markerForegroundColor: "#112233",
      });
      const second = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerSize: 14,
      });
      // size-only update must succeed; color fields still load/normalize for result snapshot
      expect(second.ok).toBe(true);
      if (second.ok) {
        expect(second.data.markerSize).toBe(14);
        expect(second.data.markerBackgroundColor).toBe("#AABBCC");
        expect(second.data.markerForegroundColor).toBe("#112233");
      }
    });

    it("fails closed when host coerces style after write", async () => {
      fake.coerceStyleAfterWrite("Circle");
      const result = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerStyle: "square",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/markerStyle readback mismatch|failed/i);
      }
    });

    it("returns unsupported when ExcelApi 1.7 missing", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartSeriesMarkersExcel({ excelApi17: false });
      const result = await new OfficeJsAdapter().updateChartSeriesMarkers({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerStyle: "dot",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.7/);
      }
    });
  });

  describe("executor + schema", () => {
    it("MockHost update succeeds and normalizes hex", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "chart.series.markers.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          markerStyle: "triangle",
          markerSize: 15,
          markerBackgroundColor: "aabbcc",
          markerForegroundColor: "#DDEEFF",
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as {
          markerStyle: string;
          markerSize: number;
          markerBackgroundColor: string;
          markerForegroundColor: string;
        };
        expect(data.markerStyle).toBe("triangle");
        expect(data.markerSize).toBe(15);
        expect(data.markerBackgroundColor).toBe("#AABBCC");
        expect(data.markerForegroundColor).toBe("#DDEEFF");
      }
    });

    it("rejects invalid args before host", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      let hostCalls = 0;
      const original = host.updateChartSeriesMarkers.bind(host);
      host.updateChartSeriesMarkers = async (input) => {
        hostCalls += 1;
        return original(input);
      };
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, markerStyle: "invalid" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, markerSize: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, markerSize: 73 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, markerBackgroundColor: "" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, markerBackgroundColor: "red" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, markerStyle: "diamond", extra: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, markerStyle: "dot" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, markerStyle: null },
      ]) {
        const result = await executor.execute({
          name: "chart.series.markers.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
      expect(hostCalls).toBe(0);
    });

    it("registers tool and keeps total count 88", () => {
      expect(TOOL_DEFINITIONS.map((t) => t.name)).toContain("chart.series.markers.update");
      expect(TOOL_DEFINITIONS).toHaveLength(98);
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.markers.update");
      expect(def?.riskLevel).toBe("moderate");
      expect(
        (def?.parameters as { additionalProperties?: boolean }).additionalProperties,
      ).toBe(false);
    });
  });

  it("WPS markers update is typed unsupported", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.markers.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        markerStyle: "circle",
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.markers.update");
      expect(detail.evidence).toMatch(/marker/i);
    }
  });
});
