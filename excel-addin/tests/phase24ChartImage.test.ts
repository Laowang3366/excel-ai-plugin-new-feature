import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installChartImageExcel } from "./fakes/officeJsChartImageFake";
import { MockHostAdapter } from "./mockHost";

describe("phase24 chart image get", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installChartImageExcel>;
    beforeEach(() => {
      fake = installChartImageExcel({
        imagePayload: "aG9zdC1iYXNlNjQ=",
        hostSheetName: "HostSheet",
        chartName: "HostChart",
      });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("returns host sheetName/chartName/Base64, not input echo", async () => {
      const result = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sheetName).toBe("HostSheet");
        expect(result.data.sheetName).not.toBe("Sheet1");
        expect(result.data.chartName).toBe("HostChart");
        expect(result.data.chartName).not.toBe("C1");
        expect(result.data.imageBase64).toBe("aG9zdC1iYXNlNjQ=");
      }
      expect(fake.getImageCalls()).toBe(1);
    });

    it("passes optional width/height into host image payload", async () => {
      const result = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
        width: 200,
        height: 100,
      });
      expect(result.ok).toBe(true);
      if (result.ok) expect(result.data.imageBase64).toBe("aG9zdC1iYXNlNjQ=:200x100");
    });

    it("skip-sync cannot read image value", async () => {
      const stale = await fake.brokenSkipSync();
      expect(stale).toBeNull();
      expect(fake.getImageCalls()).toBe(1);
      const ok = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.imageBase64).toBe("aG9zdC1iYXNlNjQ=");
    });

    it("stale snapshot: after payload A then B without sync still reads A", async () => {
      const snap = await fake.staleAfterPayloadChange();
      expect(snap.first).toBe("aG9zdC1iYXNlNjQ=");
      expect(snap.stale).toBe("aG9zdC1iYXNlNjQ=");
      expect(snap.pending).toBe("cGF5bG9hZEI=");
      expect(fake.getImageCalls()).toBe(2);

      const ok = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.imageBase64).toBe("cGF5bG9hZEI=");
    });

    it("ExcelApi 1.2 precheck false never calls getImage", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installChartImageExcel({ excelApi12: false });
      const result = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.evidence).toMatch(/ExcelApi 1\.2/);
      }
      expect(f.getImageCalls()).toBe(0);
    });

    it("empty image payload fails", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartImageExcel({ imagePayload: "" });
      const result = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).not.toBe(true);
    });

    it("1.2 true but missing getImage is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartImageExcel({ missingGetImage: true });
      const result = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.capability).toBe("chart.image.get");
        expect(result.host).toBe("office-js");
        expect(result.reason).toMatch(/getImage/i);
      }
    });

    it("sync failure is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installChartImageExcel({ syncFails: true });
      const result = await new OfficeJsAdapter().getChartImage({
        sheetName: "Sheet1",
        chartName: "C1",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/sync/i);
      }
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity with trim and dimensions", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      const def = await executor.execute({
        name: "chart.image.get",
        arguments: { sheetName: "  Sheet1  ", chartName: "  C1  " },
      });
      expect(def.ok).toBe(true);
      if (def.ok) {
        expect(def.data).toEqual({
          sheetName: "Sheet1",
          chartName: "C1",
          imageBase64: "bW9ja2ltYWdl",
        });
      }
      const sized = await executor.execute({
        name: "chart.image.get",
        arguments: { sheetName: "Sheet1", chartName: "C1", width: 320, height: 240 },
      });
      expect(sized.ok).toBe(true);
      if (sized.ok) expect(sized.data).toMatchObject({ imageBase64: "bW9ja2ltYWdl:320x240" });
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      await host.createChart({ sheetName: "Sheet1", sourceRange: "A1:B2", name: "C1" });
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1" },
        { chartName: "C1" },
        { sheetName: "", chartName: "C1" },
        { sheetName: "   ", chartName: "C1" },
        { sheetName: "Sheet1", chartName: "" },
        { sheetName: "Sheet1", chartName: "C1", width: 0 },
        { sheetName: "Sheet1", chartName: "C1", width: 4097 },
        { sheetName: "Sheet1", chartName: "C1", height: -1 },
        { sheetName: "Sheet1", chartName: "C1", width: 1.5 },
        { sheetName: "Sheet1", chartName: "C1", width: null },
        { sheetName: "Sheet1", chartName: "C1", width: undefined },
        { sheetName: "Sheet1", chartName: "C1", extra: 1 },
        { sheetName: 1, chartName: "C1" },
        { sheetName: "Sheet1", chartName: false },
        { sheetName: null, chartName: "C1" },
      ]) {
        const result = await executor.execute({
          name: "chart.image.get",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema additionalProperties false and safe risk", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "chart.image.get");
      expect(def?.riskLevel).toBe("safe");
      expect((def?.parameters as { additionalProperties?: boolean }).additionalProperties).toBe(
        false,
      );
    });
  });

  it("WPS is typed unsupported with Chart.getImage evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "chart.image.get",
      arguments: { sheetName: "Sheet1", chartName: "C1" },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as { evidence?: string; capability?: string };
      expect(detail.capability).toBe("chart.image.get");
      expect(detail.evidence).toMatch(/Chart\.getImage|export/);
    }
  });
});
