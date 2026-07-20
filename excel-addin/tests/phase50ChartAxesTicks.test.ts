import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { ToolExecutor } from "../shared/tools/executor";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { installChartAxesExcel } from "./fakes/officeJsChartAxesFake";
import { MockHostAdapter } from "./mockHost";

describe("phase50 chart axes ticks/position/linkNumberFormat", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartAxesExcel>;
    beforeEach(() => {
      fake = installChartAxesExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("updates minorUnit and tick marks with host readback", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        minorUnit: 5,
        majorTickMark: "inside",
        minorTickMark: "cross",
        tickLabelPosition: "high",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.minorUnit).toBe(5);
        expect(result.data.majorTickMark).toBe("inside");
        expect(result.data.minorTickMark).toBe("cross");
        expect(result.data.tickLabelPosition).toBe("high");
      }
      expect(fake.getCommitted("Value", "Primary")?.minorUnit).toBe(5);
      expect(fake.getCommitted("Value", "Primary")?.majorTickMark).toBe("Inside");
      expect(fake.getCommitted("Value", "Primary")?.tickLabelPosition).toBe("High");
    });

    it("sets majorUnit/minorUnit empty string to automatic (readback number)", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        majorUnit: "",
        minorUnit: "",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(typeof result.data.majorUnit).toBe("number");
        expect(typeof result.data.minorUnit).toBe("number");
      }
    });

    it("sets position custom via setPositionAt with readback", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        position: "custom",
        positionAt: 42,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.position).toBe("custom");
        expect(result.data.positionAt).toBe(42);
      }
      expect(fake.getCommitted("Value", "Primary")?.position).toBe("Custom");
      expect(fake.getCommitted("Value", "Primary")?.positionAt).toBe(42);
    });

    it("allows positionAt alone and position automatic", async () => {
      const alone = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        positionAt: 7,
      });
      expect(alone.ok).toBe(true);
      if (alone.ok) {
        expect(alone.data.position).toBe("custom");
        expect(alone.data.positionAt).toBe(7);
      }
      const auto = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        position: "maximum",
      });
      expect(auto.ok).toBe(true);
      if (auto.ok) expect(auto.data.position).toBe("maximum");
    });

    it("updates linkNumberFormat with ExcelApi 1.9 readback", async () => {
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        linkNumberFormat: false,
      });
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.data.linkNumberFormat).toBe(false);
      expect(fake.getCommitted("Value", "Primary")?.linkNumberFormat).toBe(false);
    });

    it("returns null advanced fields when ExcelApi 1.7 missing for base-only write", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      fake = installChartAxesExcel({ excelApi17: false, excelApi18: false, excelApi19: false });
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        reverse: true,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.reverse).toBe(true);
        expect(result.data.majorTickMark).toBeNull();
        expect(result.data.position).toBeNull();
        expect(result.data.linkNumberFormat).toBeNull();
      }
    });

    it("gates tick fields on ExcelApi 1.7", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartAxesExcel({ excelApi17: false });
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        majorTickMark: "none",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/1\.7/);
      }
    });

    it("gates position on ExcelApi 1.8", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartAxesExcel({ excelApi18: false });
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        position: "minimum",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/1\.8/);
      }
    });

    it("gates linkNumberFormat on ExcelApi 1.9", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartAxesExcel({ excelApi19: false });
      const result = await new OfficeJsAdapter().updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        linkNumberFormat: true,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.reason).toMatch(/1\.9/);
      }
    });

    it("does not succeed when host tampers tick mark after write", async () => {
      const adapter = new OfficeJsAdapter();
      const first = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        majorTickMark: "outside",
      });
      expect(first.ok).toBe(true);
      fake.poisonCommitted("Value", "Primary", { majorTickMark: "None" });
      const second = await adapter.updateChartAxis({
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        reverse: true,
      });
      expect(second.ok).toBe(true);
      if (second.ok) {
        expect(second.data.majorTickMark).toBe("none");
        expect(second.data.majorTickMark).not.toBe("outside");
      }
    });
  });

  describe("executor/schema/mock", () => {
    it("rejects position=custom without positionAt and invalid enums", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1", chartName: "C1", kind: "value", position: "custom" },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", majorTickMark: "diagonal" },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", minorUnit: null },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", linkNumberFormat: "yes" },
        { sheetName: "Sheet1", chartName: "C1", kind: "value", positionAt: null },
      ]) {
        const result = await executor.execute({
          name: "chart.axes.update",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("accepts minorUnit empty string and new fields via mock host", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const result = await executor.execute({
        name: "chart.axes.update",
        arguments: {
          sheetName: "Sheet1",
          chartName: "C1",
          kind: "value",
          minorUnit: "",
          majorTickMark: "cross",
          position: "custom",
          positionAt: 3,
          linkNumberFormat: false,
        },
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data).toMatchObject({
          majorTickMark: "cross",
          position: "custom",
          positionAt: 3,
          linkNumberFormat: false,
          minorUnit: 2,
        });
      }
    });

    it("schema lists new enums without adding a tool", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.axes.update");
      expect(def).toBeTruthy();
      const props = (def?.parameters as { properties: Record<string, unknown> }).properties;
      expect(props.minorUnit).toBeTruthy();
      expect(props.majorTickMark).toBeTruthy();
      expect(props.position).toBeTruthy();
      expect(props.linkNumberFormat).toBeTruthy();
      expect(TOOL_DEFINITIONS).toHaveLength(89);
    });
  });

  it("WPS remains typed unsupported for new fields", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.axes.update",
      arguments: {
        sheetName: "Sheet1",
        chartName: "C1",
        kind: "value",
        majorTickMark: "inside",
      },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.axes.update");
      expect(detail.evidence).toMatch(/ChartAxis/);
    }
  });
});
