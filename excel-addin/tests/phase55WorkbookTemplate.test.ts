import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { CHAT_READONLY_TOOL_ALLOWLIST } from "../shared/agentChat/chatReadOnlyTools";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { WORKBOOK_TEMPLATE_PRESET_STYLES } from "../shared/host/workbookTemplateTypes";
import { TOOL_DEFINITIONS, ToolExecutor } from "../shared/tools";
import { installTemplateExcel } from "./fakes/officeJsTemplateFake";
import { MockHostAdapter } from "./mockHost";

describe("phase55 workbook.template", () => {
  it("registers apply+capture; total 98 unique tools", () => {
    const names = TOOL_DEFINITIONS.map((d) => d.name);
    expect(TOOL_DEFINITIONS).toHaveLength(98);
    expect(new Set(names).size).toBe(98);
    expect(names).toContain("workbook.template.apply");
    expect(names).toContain("workbook.template.capture");
  });

  it("risk and readonly allowlist", () => {
    const apply = TOOL_DEFINITIONS.find((d) => d.name === "workbook.template.apply")!;
    const capture = TOOL_DEFINITIONS.find((d) => d.name === "workbook.template.capture")!;
    expect(apply.riskLevel).toBe("dangerous");
    expect(capture.riskLevel).toBe("safe");
    expect(CHAT_READONLY_TOOL_ALLOWLIST).toContain("workbook.template.capture");
    expect(CHAT_READONLY_TOOL_ALLOWLIST).not.toContain("workbook.template.apply");
    expect((apply.parameters as { additionalProperties: boolean }).additionalProperties).toBe(
      false,
    );
    expect((capture.parameters as { additionalProperties: boolean }).additionalProperties).toBe(
      false,
    );
  });

  it("executor rejects null/unknown before Host", async () => {
    const host = new MockHostAdapter();
    const ex = new ToolExecutor(host);
    let hostCalls = 0;
    const orig = host.applyWorkbookTemplate.bind(host);
    host.applyWorkbookTemplate = async (input) => {
      hostCalls += 1;
      return orig(input);
    };
    const unknown = await ex.execute({
      name: "workbook.template.apply",
      arguments: { extra: true },
    });
    expect(unknown.ok).toBe(false);
    expect(hostCalls).toBe(0);

    const nullPreset = await ex.execute({
      name: "workbook.template.apply",
      arguments: { preset: null },
    });
    expect(nullPreset.ok).toBe(false);
    expect(hostCalls).toBe(0);

    const dup = await ex.execute({
      name: "workbook.template.apply",
      arguments: { sheetNames: ["Sheet1", "sheet1"] },
    });
    expect(dup.ok).toBe(false);
    expect(hostCalls).toBe(0);

    const blank = await ex.execute({
      name: "workbook.template.apply",
      arguments: { sheetNames: ["  "] },
    });
    expect(blank.ok).toBe(false);
    expect(hostCalls).toBe(0);

    const captureUnknown = await ex.execute({
      name: "workbook.template.capture",
      arguments: { x: 1 },
    });
    expect(captureUnknown.ok).toBe(false);
  });

  it("MockHost apply defaults and capture", async () => {
    const host = new MockHostAdapter();
    const ex = new ToolExecutor(host);
    const applied = await ex.execute({ name: "workbook.template.apply", arguments: {} });
    expect(applied.ok).toBe(true);
    if (applied.ok) {
      const data = applied.data as {
        preset: string;
        appliedSheetCount: number;
        appliedSheets: { readback: { fontName: string; freezeRowCount: number } }[];
      };
      expect(data.preset).toBe("professional");
      expect(data.appliedSheetCount).toBe(1);
      expect(data.appliedSheets[0]!.readback.fontName).toBe("微软雅黑");
      expect(data.appliedSheets[0]!.readback.freezeRowCount).toBe(1);
    }
    const cap = await ex.execute({ name: "workbook.template.capture", arguments: {} });
    expect(cap.ok).toBe(true);
  });

  describe("Office.js apply", () => {
    let gates: ReturnType<typeof installTemplateExcel>;
    beforeEach(() => {
      gates = installTemplateExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("applies all four presets with host readback", async () => {
      const adapter = new OfficeJsAdapter();
      for (const preset of ["professional", "financial", "dashboard", "minimal"] as const) {
        gates = installTemplateExcel();
        const result = await adapter.applyWorkbookTemplate({
          preset,
          allSheets: false,
          fontName: "微软雅黑",
          fontSize: 10.5,
          autoFit: true,
          showGridlines: false,
          freezeRows: 1,
        });
        expect(result.ok).toBe(true);
        if (result.ok) {
          expect(result.data.preset).toBe(preset);
          expect(result.data.appliedSheetCount).toBe(1);
          const rb = result.data.appliedSheets[0]!.readback;
          expect(rb.headerFill).toBe(WORKBOOK_TEMPLATE_PRESET_STYLES[preset].headerFill);
          expect(rb.headerFontColor).toBe(
            WORKBOOK_TEMPLATE_PRESET_STYLES[preset].headerFontColor,
          );
          expect(rb.headerBold).toBe(true);
          expect(rb.headerWrapText).toBe(true);
          expect(rb.headerRowHeight).toBe(24);
          expect(rb.autoFitVerified).toBe(false);
          expect(rb.freezeRowCount).toBe(1);
        }
      }
    });

    it("sheetNames missing fails before write; empty skipped", async () => {
      const adapter = new OfficeJsAdapter();
      const missing = await adapter.applyWorkbookTemplate({
        preset: "professional",
        sheetNames: ["NoSuch"],
        allSheets: true,
        fontName: "微软雅黑",
        fontSize: 10.5,
        autoFit: false,
        showGridlines: false,
        freezeRows: 0,
      });
      expect(missing.ok).toBe(false);
      expect(gates.writeCalls()).toBe(0);

      const withEmpty = await adapter.applyWorkbookTemplate({
        preset: "professional",
        sheetNames: ["Sheet1", "Empty"],
        allSheets: true,
        fontName: "Arial",
        fontSize: 12,
        autoFit: false,
        showGridlines: true,
        freezeRows: 0,
      });
      expect(withEmpty.ok).toBe(true);
      if (withEmpty.ok) {
        expect(withEmpty.data.appliedSheetCount).toBe(1);
        expect(withEmpty.data.skippedSheets.some((s) => s.name === "Empty")).toBe(true);
        expect(withEmpty.data.appliedSheets[0]!.readback.freezeRowCount).toBe(0);
        expect(withEmpty.data.appliedSheets[0]!.readback.showGridlines).toBe(true);
      }
    });

    it("requirement 1.8 false/missing/throw → typed unsupported + zero Excel.run", async () => {
      for (const opts of [
        { excelApi18: false },
        { missingIsSetSupported: true },
        { isSetSupportedThrows: true },
      ] as const) {
        delete (globalThis as { Excel?: unknown }).Excel;
        delete (globalThis as { Office?: unknown }).Office;
        const g = installTemplateExcel(opts);
        const adapter = new OfficeJsAdapter();
        const result = await adapter.applyWorkbookTemplate({
          preset: "professional",
          allSheets: true,
          fontName: "微软雅黑",
          fontSize: 10.5,
          autoFit: true,
          showGridlines: false,
          freezeRows: 1,
        });
        expect(result.ok).toBe(false);
        if (!result.ok) expect(result.unsupported).toBe(true);
        expect(g.excelRunCalls()).toBe(0);
        expect(g.writeCalls()).toBe(0);
      }
      gates = installTemplateExcel();
    });

    it("member missing after requirement → ordinary failed", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const g = installTemplateExcel({ missingFreeze: true });
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate({
        preset: "professional",
        allSheets: false,
        fontName: "微软雅黑",
        fontSize: 10.5,
        autoFit: true,
        showGridlines: false,
        freezeRows: 1,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).not.toBe(true);
      expect(g.writeCalls()).toBe(0);
      gates = installTemplateExcel();
    });

    it("poisoned readback fails (no coercion success)", async () => {
      const adapter = new OfficeJsAdapter();
      gates.setPoison("Sheet1", { headerFill: 123 });
      const result = await adapter.applyWorkbookTemplate({
        preset: "professional",
        allSheets: false,
        fontName: "微软雅黑",
        fontSize: 10.5,
        autoFit: false,
        showGridlines: false,
        freezeRows: 1,
      });
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).not.toBe(true);
    });

    it("freezeRows 0 clears; header rowHeight 24 with autoFit", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.applyWorkbookTemplate({
        preset: "financial",
        sheetNames: ["Sheet1"],
        allSheets: true,
        fontName: "Segoe UI",
        fontSize: 11,
        autoFit: true,
        showGridlines: false,
        freezeRows: 0,
      });
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.appliedSheets[0]!.readback.freezeRowCount).toBe(0);
        expect(result.data.appliedSheets[0]!.readback.headerRowHeight).toBe(24);
      }
    });
  });

  describe("Office.js capture", () => {
    beforeEach(() => {
      installTemplateExcel();
    });
    afterEach(() => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
    });

    it("captures multi-sheet shallow snapshot including empty", async () => {
      const adapter = new OfficeJsAdapter();
      const result = await adapter.captureWorkbookTemplate();
      expect(result.ok).toBe(true);
      if (result.ok) {
        expect(result.data.template.version).toBe(1);
        expect(result.data.sheetCount).toBeGreaterThanOrEqual(3);
        const empty = result.data.template.sheets.find((s) => s.name === "Empty");
        expect(empty?.usedRange).toBeNull();
        expect(empty?.baseStyle).toBeNull();
        const s1 = result.data.template.sheets.find((s) => s.name === "Sheet1");
        expect(s1?.usedRange).toBeTruthy();
        expect(s1?.print.orientation).toBe("portrait");
        expect(result.data.limitations.some((l) => l.includes("Shallow"))).toBe(true);
      }
    });

    it("ExcelApi 1.9 unsupported zero run; >500 resource-limit", async () => {
      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      const g = installTemplateExcel({ excelApi19: false });
      const adapter = new OfficeJsAdapter();
      const unsup = await adapter.captureWorkbookTemplate();
      expect(unsup.ok).toBe(false);
      if (!unsup.ok) expect(unsup.unsupported).toBe(true);
      expect(g.excelRunCalls()).toBe(0);

      delete (globalThis as { Excel?: unknown }).Excel;
      delete (globalThis as { Office?: unknown }).Office;
      installTemplateExcel({ sheetCount: 501 });
      const limited = await adapter.captureWorkbookTemplate();
      expect(limited.ok).toBe(false);
      if (!limited.ok) {
        expect(limited.unsupported).not.toBe(true);
        expect(limited.reason).toMatch(/resource-limit/);
      }
      installTemplateExcel();
    });
  });

  it("WPS typed unsupported for both tools", async () => {
    const wps = new WpsJsaAdapter();
    const apply = await wps.applyWorkbookTemplate({
      preset: "professional",
      allSheets: true,
      fontName: "微软雅黑",
      fontSize: 10.5,
      autoFit: true,
      showGridlines: false,
      freezeRows: 1,
    });
    expect(apply.ok).toBe(false);
    if (!apply.ok) expect(apply.unsupported).toBe(true);
    const capture = await wps.captureWorkbookTemplate();
    expect(capture.ok).toBe(false);
    if (!capture.ok) expect(capture.unsupported).toBe(true);
  });
});
