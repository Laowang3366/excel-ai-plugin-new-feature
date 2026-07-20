import { afterEach, describe, expect, it, vi } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { unsupported } from "../shared/host/types";

describe("host adapters without host runtime", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("Office.js adapter reports disconnected status when Excel.run is missing", async () => {
    const adapter = new OfficeJsAdapter();
    const status = await adapter.getStatus();
    expect(status.ok).toBe(true);
    if (status.ok) {
      expect(status.data.connected).toBe(false);
      expect(status.data.kind).toBe("office-js");
    }
  });

  it("Office.js range ops return typed unsupported without Excel.run", async () => {
    const adapter = new OfficeJsAdapter();
    const result = await adapter.readRange("Sheet1", "A1");
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      expect(result.capability).toBe("range.read");
    }
  });

  it("Office.js detects dynamic arrays from the ExcelApi 1.12 requirement set", () => {
    const isSetSupported = vi.fn(
      (name: string, version: string) =>
        name === "ExcelApi" && version === "1.12",
    );
    vi.stubGlobal("window", {
      Office: { context: { requirements: { isSetSupported } } },
    });

    expect(new OfficeJsAdapter().getRuntimeCapabilities()).toEqual({
      dynamicArrayFunctionsEnabled: true,
    });
    expect(isSetSupported).toHaveBeenCalledWith("ExcelApi", "1.12");
  });

  it("WPS adapter returns typed unsupported without Application", async () => {
    const adapter = new WpsJsaAdapter();
    const result = await adapter.getSelection();
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      expect(result.host).toBe("wps-jsa");
    }
  });

  it("WPS keeps dynamic arrays disabled until a real JSA contract is verified", () => {
    expect(new WpsJsaAdapter().getRuntimeCapabilities()).toEqual({
      dynamicArrayFunctionsEnabled: false,
    });
  });

  it("unsupported helper keeps contract shape", () => {
    const result = unsupported("macro.run", "wps-jsa", "not in first batch");
    expect(result).toEqual({
      ok: false,
      unsupported: true,
      capability: "macro.run",
      host: "wps-jsa",
      reason: "not in first batch",
      evidence: undefined,
    });
  });
});
