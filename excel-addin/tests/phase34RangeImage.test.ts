import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installRangeImageExcel } from "./fakes/officeJsRangeImageFake";
import { MockHostAdapter } from "./mockHost";
import { buildAdvancedExcelBoundary } from "../shared/prompts/advancedExcelBoundary";

describe("phase34 range image get", () => {
  describe("Office.js sync-gated", () => {
    let fake: ReturnType<typeof installRangeImageExcel>;
    beforeEach(() => {
      fake = installRangeImageExcel({
        imagePayload: "aG9zdC1yYW5nZS1iYXNlNjQ=",
        hostSheetName: "HostSheet",
        hostAddress: "HostSheet!$A$1:$B$2",
      });
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("returns host sheetName/address/Base64, not input echo", async () => {
      const result = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.sheetName).toBe("HostSheet");
        expect(result.data.sheetName).not.toBe("Sheet1");
        expect(result.data.address).toBe("HostSheet!$A$1:$B$2");
        expect(result.data.address).not.toBe("A1:B2");
        expect(result.data.imageBase64).toBe("aG9zdC1yYW5nZS1iYXNlNjQ=");
      }
      expect(fake.getImageCalls()).toBe(1);
    });

    it("skip-sync cannot read image value", async () => {
      const stale = await fake.brokenSkipSync();
      expect(stale).toBeNull();
      expect(fake.getImageCalls()).toBe(1);
      const ok = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.imageBase64).toBe("aG9zdC1yYW5nZS1iYXNlNjQ=");
    });

    it("stale snapshot: after payload A then B without sync still reads A", async () => {
      const snap = await fake.staleAfterPayloadChange();
      expect(snap.first).toBe("aG9zdC1yYW5nZS1iYXNlNjQ=");
      expect(snap.stale).toBe("aG9zdC1yYW5nZS1iYXNlNjQ=");
      expect(snap.pending).toBe("cGF5bG9hZEI=");
      expect(fake.getImageCalls()).toBe(2);

      const ok = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(ok.ok).toBe(true);
      if (ok.ok) expect(ok.data.imageBase64).toBe("cGF5bG9hZEI=");
    });

    it("ExcelApi 1.7 precheck false never calls getImage", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const f = installRangeImageExcel({ excelApi17: false });
      const result = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(result.evidence).toMatch(/ExcelApi 1\.7|Range\.getImage/);
      }
      expect(f.getImageCalls()).toBe(0);
    });

    it("1.7 true but missing getImage is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installRangeImageExcel({ missingGetImage: true });
      const result = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.capability).toBe("range.image.get");
        expect(result.host).toBe("office-js");
        expect(result.reason).toMatch(/getImage/i);
      }
    });

    it("empty image payload is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installRangeImageExcel({ imagePayload: "" });
      const result = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).not.toBe(true);
    });

    it("non-string host name/address is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installRangeImageExcel({ hostSheetName: 12 });
      const badName = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(badName.ok).toBe(false);
      if (!badName.ok) expect(badName.unsupported).not.toBe(true);

      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installRangeImageExcel({ hostAddress: null });
      const badAddr = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(badAddr.ok).toBe(false);
      if (!badAddr.ok) expect(badAddr.unsupported).not.toBe(true);
    });

    it("sync failure is ordinary fail", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installRangeImageExcel({ syncFails: true });
      const result = await new OfficeJsAdapter().getRangeImage({
        sheetName: "Sheet1",
        range: "A1:B2",
      });
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).not.toBe(true);
        expect(result.reason).toMatch(/sync/i);
      }
    });
  });

  describe("executor + schema", () => {
    it("MockHost parity with trim", async () => {
      const host = new MockHostAdapter();
      const executor = new ToolExecutor(host);
      const def = await executor.execute({
        name: "range.image.get",
        arguments: { sheetName: "  Sheet1  ", range: "  A1:B2  " },
      });
      expect(def.ok).toBe(true);
      if (def.ok) {
        expect(def.data).toEqual({
          sheetName: "Sheet1",
          address: "Sheet1!A1:B2",
          imageBase64: "bW9ja3JhbmdlaW1hZ2U=",
        });
      }
    });

    it("rejects invalid args", async () => {
      const host = new MockHostAdapter();
      const executor = new ToolExecutor(host);
      for (const args of [
        { sheetName: "Sheet1" },
        { range: "A1" },
        { sheetName: "", range: "A1" },
        { sheetName: "   ", range: "A1" },
        { sheetName: "Sheet1", range: "" },
        { sheetName: "Sheet1", range: "   " },
        { sheetName: "Sheet1", range: "A1", extra: 1 },
        { sheetName: 1, range: "A1" },
        { sheetName: "Sheet1", range: false },
        { sheetName: null, range: "A1" },
        { sheetName: undefined, range: "A1" },
        { sheetName: "Sheet1", range: null },
        { sheetName: "Sheet1", range: undefined },
      ]) {
        const result = await executor.execute({
          name: "range.image.get",
          arguments: args as Record<string, unknown>,
        });
        expect(result.ok).toBe(false);
      }
    });

    it("schema only sheetName/range, minLength1, additionalProperties false, safe", () => {
      const def = TOOL_DEFINITIONS.find((t) => t.name === "range.image.get");
      expect(def?.riskLevel).toBe("safe");
      const params = def?.parameters as {
        additionalProperties?: boolean;
        properties?: Record<string, { type?: string; minLength?: number }>;
        required?: string[];
      };
      expect(params.additionalProperties).toBe(false);
      expect(Object.keys(params.properties ?? {}).sort()).toEqual(["range", "sheetName"]);
      expect(params.properties?.sheetName?.minLength).toBe(1);
      expect(params.properties?.range?.minLength).toBe(1);
      expect(params.required).toEqual(["sheetName", "range"]);
    });
  });

  it("WPS is typed unsupported with Range.getImage evidence", async () => {
    const result = await new ToolExecutor(new WpsJsaAdapter()).execute({
      name: "range.image.get",
      arguments: { sheetName: "Sheet1", range: "A1:B2" },
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      const detail = result.detail as {
        capability?: string;
        evidence?: string;
        host?: string;
      };
      expect(detail.capability).toBe("range.image.get");
      expect(detail.host).toBe("wps-jsa");
      expect(detail.evidence).toMatch(/Range\.getImage|export|Base64/i);
    }
  });

  it("chart.image.get still registered and not confused", () => {
    expect(TOOL_DEFINITIONS.some((t) => t.name === "chart.image.get")).toBe(true);
    expect(TOOL_DEFINITIONS.some((t) => t.name === "range.image.get")).toBe(true);
  });

  it("prompt boundary mentions range.image.get ExcelApi 1.7 memory Base64", () => {
    const boundary = buildAdvancedExcelBoundary({});
    expect(boundary).toContain("`range.image.get`");
    expect(boundary).toMatch(/ExcelApi 1\.7|Range\.getImage|Base64/);
    expect(boundary).toMatch(/WPS/);
    expect(boundary).toContain("`chart.image.get`");
  });
});
