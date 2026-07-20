import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { listChatReadOnlyTools } from "../shared/agentChat/chatReadOnlyTools";
import { installChartSeriesTrendlinesExcel } from "./fakes/officeJsChartSeriesTrendlinesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase47 chart series trendlines", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartSeriesTrendlinesExcel>;
    beforeEach(() => {
      fake = installChartSeriesTrendlinesExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("adds linear trendline and lists with host readback", async () => {
      const adapter = new OfficeJsAdapter();
      const added = await adapter.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
        name: "TL1",
        showEquation: true,
      });
      expect(added.ok).toBe(true);
      if (added.ok) {
        expect(added.data).toMatchObject({
          seriesIndex: 1,
          trendlineIndex: 1,
          type: "linear",
          name: "TL1",
          showEquation: true,
        });
      }
      const listed = await adapter.listChartSeriesTrendlines("Sheet1", "C1", 1);
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect(listed.data.trendlines).toHaveLength(1);
        expect(listed.data.trendlines[0]?.type).toBe("linear");
      }
      expect(fake.getCommitted()[0]?.type).toBe("Linear");
    });

    it("updates polynomial order and deletes with remaining readback", async () => {
      const adapter = new OfficeJsAdapter();
      const a = await adapter.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "polynomial",
        polynomialOrder: 3,
      });
      expect(a.ok).toBe(true);
      const b = await adapter.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
      });
      expect(b.ok).toBe(true);
      const updated = await adapter.updateChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        trendlineIndex: 1,
        polynomialOrder: 4,
        showRSquared: true,
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(updated.data.polynomialOrder).toBe(4);
        expect(updated.data.showRSquared).toBe(true);
      }
      const deleted = await adapter.deleteChartSeriesTrendline("Sheet1", "C1", 1, 1);
      expect(deleted.ok).toBe(true);
      if (deleted.ok) {
        expect(deleted.data.deletedTrendlineIndex).toBe(1);
        expect(deleted.data.remainingTrendlines).toHaveLength(1);
        expect(deleted.data.remainingTrendlines[0]?.type).toBe("linear");
        expect(deleted.data.remainingTrendlines[0]?.trendlineIndex).toBe(1);
      }
    });

    it("returns unsupported when ExcelApi 1.7 missing", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartSeriesTrendlinesExcel({ excelApi17: false });
      const result = await new OfficeJsAdapter().listChartSeriesTrendlines("Sheet1", "C1", 1);
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.7/);
      }
    });

    it("returns unsupported for 1.8 fields when ExcelApi 1.8 missing", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartSeriesTrendlinesExcel({ excelApi17: true, excelApi18: false });
      const result = await new OfficeJsAdapter().addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
        showEquation: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/ExcelApi 1\.8/);
      }
    });


    it("assigns stable 1-based indexes for multi-trendline add via returned object/getCount", async () => {
      const adapter = new OfficeJsAdapter();
      const first = await adapter.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
        name: "First",
      });
      const second = await adapter.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "exponential",
        name: "Second",
      });
      expect(first.ok).toBe(true);
      expect(second.ok).toBe(true);
      if (first.ok && second.ok) {
        expect(first.data.trendlineIndex).toBe(1);
        expect(second.data.trendlineIndex).toBe(2);
        expect(first.data.name).toBe("First");
        expect(second.data.name).toBe("Second");
        expect(first.data.type).toBe("linear");
        expect(second.data.type).toBe("exponential");
      }
      const listed = await adapter.listChartSeriesTrendlines("Sheet1", "C1", 1);
      expect(listed.ok).toBe(true);
      if (listed.ok) {
        expect(listed.data.trendlines.map((t) => t.trendlineIndex)).toEqual([1, 2]);
        expect(listed.data.trendlines.map((t) => t.name)).toEqual(["First", "Second"]);
      }
    });

    it("writes intercept empty string as automatic and readbacks a number", async () => {
      const adapter = new OfficeJsAdapter();
      const added = await adapter.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
        intercept: "",
      });
      expect(added.ok).toBe(true);
      if (added.ok) {
        expect(typeof added.data.intercept).toBe("number");
        expect(Number.isFinite(added.data.intercept as number)).toBe(true);
      }
      const updated = await adapter.updateChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        trendlineIndex: 1,
        intercept: "",
      });
      expect(updated.ok).toBe(true);
      if (updated.ok) {
        expect(typeof updated.data.intercept).toBe("number");
      }
    });

    it("fails closed when host tampers type after write", async () => {
      const adapter = new OfficeJsAdapter();
      const okAdd = await adapter.addChartSeriesTrendline({
        sheetName: "Sheet1",
        chartName: "C1",
        seriesIndex: 1,
        type: "linear",
      });
      expect(okAdd.ok).toBe(true);
      fake.poisonLast({ type: { bad: true } as unknown as string });
      const listed = await adapter.listChartSeriesTrendlines("Sheet1", "C1", 1);
      expect(listed.ok).toBe(false);
    });
  });

  describe("executor + schema", () => {
    it("MockHost full lifecycle", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const add = await executor.execute({
        name: "chart.series.trendlines.add",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          type: "movingAverage",
          movingAveragePeriod: 3,
        },
      });
      expect(add.ok).toBe(true);
      const list = await executor.execute({
        name: "chart.series.trendlines.list",
        arguments: { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
      });
      expect(list.ok).toBe(true);
      if (list.ok) {
        const data = list.data as { trendlines: Array<{ type: string }> };
        expect(data.trendlines[0]?.type).toBe("movingAverage");
      }
    });


    it("MockHost intercept empty string stores numeric automatic readback", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const add = await executor.execute({
        name: "chart.series.trendlines.add",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          type: "linear",
          intercept: "",
        },
      });
      expect(add.ok).toBe(true);
      if (add.ok) {
        const data = add.data as { intercept: number | null };
        expect(typeof data.intercept).toBe("number");
      }
    });

    it("rejects invalid args before host", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      let hostCalls = 0;
      const original = host.addChartSeriesTrendline.bind(host);
      host.addChartSeriesTrendline = async (input) => {
        hostCalls += 1;
        return original(input);
      };
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, type: "spline" },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          type: "linear",
          polynomialOrder: 3,
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          type: "polynomial",
          polynomialOrder: 1,
        },
        {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          type: "linear",
          extra: true,
        },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 0, type: "linear" },
        { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1, type: null },
      ]) {
        const result = await executor.execute({
          name: "chart.series.trendlines.add",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
      expect(hostCalls).toBe(0);
    });

    it("registers four tools; list is read-only safe", () => {
      const names = TOOL_DEFINITIONS.map((t) => t.name);
      expect(names).toContain("chart.series.trendlines.list");
      expect(names).toContain("chart.series.trendlines.add");
      expect(names).toContain("chart.series.trendlines.update");
      expect(names).toContain("chart.series.trendlines.delete");
      expect(TOOL_DEFINITIONS).toHaveLength(96);
      expect(listChatReadOnlyTools().map((d) => d.name)).toContain(
        "chart.series.trendlines.list",
      );
      const listDef = TOOL_DEFINITIONS.find((t) => t.name === "chart.series.trendlines.list");
      expect(listDef?.riskLevel).toBe("safe");
      for (const name of [
        "chart.series.trendlines.add",
        "chart.series.trendlines.update",
        "chart.series.trendlines.delete",
      ]) {
        expect(TOOL_DEFINITIONS.find((t) => t.name === name)?.riskLevel).toBe("moderate");
        expect(
          (TOOL_DEFINITIONS.find((t) => t.name === name)?.parameters as {
            additionalProperties?: boolean;
          }).additionalProperties,
        ).toBe(false);
      }
    });
  });

  it("WPS all four tools are typed unsupported", async () => {
    const executor = new ToolExecutor(new WpsJsaAdapter());
    for (const call of [
      {
        name: "chart.series.trendlines.list" as const,
        arguments: { sheetName: "Sheet1", chartName: "C1", seriesIndex: 1 },
      },
      {
        name: "chart.series.trendlines.add" as const,
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          type: "linear",
        },
      },
      {
        name: "chart.series.trendlines.update" as const,
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
          name: "X",
        },
      },
      {
        name: "chart.series.trendlines.delete" as const,
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          seriesIndex: 1,
          trendlineIndex: 1,
        },
      },
    ]) {
      const result = await executor.execute(call);
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        const detail = result.detail as { evidence?: string; capability?: string };
        expect(detail.capability).toBe(call.name);
        expect(detail.evidence).toMatch(/trendline/i);
      }
    }
  });
});
