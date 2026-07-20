import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { installChartSeriesTrendlineFormatExcel } from "./fakes/officeJsChartSeriesTrendlineFormatFake";
import { MockHostAdapter } from "./mockHost";

describe("phase49 chart series trendline format", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesTrendlineFormatExcel>;
    beforeEach(() => {
      fake = installChartSeriesTrendlineFormatExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("updates line format with host readback (not input echo)", async () => {
      fake.setChartName("HostChart");
      const result = await new OfficeJsAdapter().updateChartSeriesTrendlineFormat({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        trendlineIndex: 1,
        color: "#FF0000",
        lineStyle: "dash",
        weight: 2.5,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.chartName).toBe("HostChart");
        expect(result.data.chartName).not.toBe("C1");
        expect(result.data.color).toBe("#FF0000");
        expect(result.data.lineStyle).toBe("dash");
        expect(result.data.weight).toBe(2.5);
        expect(result.data.seriesIndex).toBe(1);
        expect(result.data.trendlineIndex).toBe(1);
      }
      expect(fake.getCommitted(0, 0)?.lineStyle).toBe("Dash");
      expect(fake.getCommitted(0, 0)?.color).toBe("#FF0000");
      expect(fake.getCommitted(0, 1)?.lineStyle).toBe("Continuous");
    });

    it("normalizes lowercase color on direct adapter path", async () => {
      const result = await new OfficeJsAdapter().updateChartSeriesTrendlineFormat({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        trendlineIndex: 2,
        color: "aabbcc",
        lineStyle: "dot",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.trendlineIndex).toBe(2);
        expect(result.data.color).toBe("#AABBCC");
        expect(result.data.lineStyle).toBe("dot");
      }
      expect(fake.getCommitted(0, 1)?.color).toBe("#AABBCC");
      expect(fake.getCommitted(0, 1)?.lineStyle).toBe("Dot");
    });

    it("fails closed when host coerces lineStyle after write", async () => {
      fake.coerceStyleAfterWrite("Continuous");
      const result = await new OfficeJsAdapter().updateChartSeriesTrendlineFormat({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        trendlineIndex: 1,
        lineStyle: "dashDot",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/lineStyle readback mismatch|failed/i);
      }
    });

    it("returns unsupported when ExcelApi 1.7 missing", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartSeriesTrendlineFormatExcel({ excelApi17: false });
      const result = await new OfficeJsAdapter().updateChartSeriesTrendlineFormat({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        trendlineIndex: 1,
        color: "#112233",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.7/);
      }
    });
  });

  describe("executor + schema", () => {
    it("MockHost update succeeds", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      await host.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
      });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "chart.series.trendlines.format.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
          color: "ff00aa",
          lineStyle: "roundDot",
          weight: 3,
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        const data = result.data as {
          color: string;
          lineStyle: string;
          weight: number;
        };
        expect(data.color).toBe("#FF00AA");
        expect(data.lineStyle).toBe("roundDot");
        expect(data.weight).toBe(3);
      }
    });

    it("rejects invalid args before host", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      await host.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
      });
      let hostCalls = 0;
      const original = host.updateChartSeriesTrendlineFormat.bind(host);
      host.updateChartSeriesTrendlineFormat = async (input) => {
        hostCalls += 1;
        return original(input);
      };
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, trendlineIndex: 1 },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
          lineStyle: "solid",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
          weight: 0,
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
          color: "",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
          color: "red",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
          color: "#FF0000",
          extra: true,
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 0,
          trendlineIndex: 1,
          color: "#FF0000",
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: null,
          color: "#FF0000",
        },
      ]) {
        const result = await executor.execute({
          name: "chart.series.trendlines.format.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
      expect(hostCalls).toBe(0);
    });

    it("registers tool and keeps total count 89", () => {
      expect(TOOL_DEFINITIONS.map((t) => t.name)).toContain(
        "chart.series.trendlines.format.update",
      );
      expect(TOOL_DEFINITIONS).toHaveLength(89);
      const def = TOOL_DEFINITIONS.find(
        (t) => t.name === "chart.series.trendlines.format.update",
      );
      expect(def?.riskLevel).toBe("moderate");
      expect(
        (def?.parameters as { additionalProperties?: boolean }).additionalProperties,
      ).toBe(false);
    });
  });

  it("WPS trendline format is typed unsupported", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.series.trendlines.format.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        trendlineIndex: 1,
        color: "#FF0000",
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.series.trendlines.format.update");
      expect(detail.evidence).toMatch(/trendline|format|line/i);
    }
  });
});
