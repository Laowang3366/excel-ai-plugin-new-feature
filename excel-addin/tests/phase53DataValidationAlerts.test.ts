/**
 * Phase53.1: DataValidation errorAlert/prompt ClientObject contract
 * (load top-level → sync → whole-object assign → sync → reload).
 */
import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { DV_FULL_LOAD_PROPS } from "../shared/host/officeJsValidationAlerts";
import {
  MAX_DV_ALERT_MESSAGE_CHARS,
  MAX_DV_ALERT_TITLE_CHARS,
  mapAlertStyleToHost,
  unmapAlertStyle,
} from "../shared/host/officeJsValidationAlerts";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { ToolExecutor } from "../shared/tools/executor";
import { installValidationExcel } from "./fakes/officeJsValidationFake";
import {
  installWpsValidationFake,
  uninstallWpsValidationFake,
} from "./fakes/wpsJsaValidationFake";
import { MockHostAdapter } from "./mockHost";

describe("phase53.1 dataValidation alerts ClientObject contract", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
    delete (globalThis as { Application?: unknown }).Application;
    uninstallWpsValidationFake();
  });

  it("maps alert styles bidirectionally (Stop/Warning/Information)", () => {
    expect(mapAlertStyleToHost("stop")).toBe("Stop");
    expect(mapAlertStyleToHost("warning")).toBe("Warning");
    expect(mapAlertStyleToHost("information")).toBe("Information");
    expect(unmapAlertStyle("Stop")).toBe("stop");
    expect(unmapAlertStyle("WARNING")).toBe("warning");
    expect(unmapAlertStyle("Stop-Alert")).toBeUndefined();
  });

  it("uses top-level load props only (no nested errorAlert/... paths)", async () => {
    const ctrl = installValidationExcel();
    ctrl.resetLoadPropsLog();
    await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "H1",
      rule: { type: "wholeNumber", operator: "equalTo", formula1: "1" },
      errorAlert: { showAlert: true, style: "stop", title: "T", message: "M" },
    });
    const logs = ctrl.getLoadPropsLog();
    expect(logs.length).toBeGreaterThan(0);
    for (const p of logs) {
      expect(p).toBe(DV_FULL_LOAD_PROPS);
      expect(p).not.toMatch(/errorAlert\//);
      expect(p).not.toMatch(/prompt\//);
    }
  });

  it("round-trips full errorAlert+prompt+allowBlank; sync order pre→write→reload", async () => {
    const ctrl = installValidationExcel();
    ctrl.resetSyncCount();
    ctrl.resetDvWriteCounts();
    const adapter = new OfficeJsAdapter();
    const written = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "H1:H3",
      rule: {
        type: "wholeNumber",
        operator: "between",
        formula1: "1",
        formula2: "10",
        allowBlank: false,
      },
      errorAlert: {
        showAlert: true,
        style: "warning",
        title: "Bad",
        message: "Enter 1-10",
      },
      prompt: { showPrompt: true, title: "Hint", message: "Whole number" },
    });
    expect(written.ok).toBe(true);
    if (!written.ok) return;
    expect(written.data.rule?.allowBlank).toBe(false);
    expect(written.data.errorAlert).toEqual({
      showAlert: true,
      style: "warning",
      title: "Bad",
      message: "Enter 1-10",
    });
    expect(written.data.prompt).toEqual({
      showPrompt: true,
      title: "Hint",
      message: "Whole number",
    });
    // pre-read sync + write sync + reload sync
    expect(ctrl.getSyncCount()).toBeGreaterThanOrEqual(3);
    const counts = ctrl.getDvWriteCounts();
    expect(counts.rule).toBe(1);
    expect(counts.ignoreBlanks).toBe(1);
    expect(counts.errorAlert).toBe(1);
    expect(counts.prompt).toBe(1);

    const read = await adapter.readDataValidation("Sheet1", "H1:H3");
    expect(read.ok).toBe(true);
    if (read.ok) {
      expect(read.data.errorAlert?.style).toBe("warning");
      expect(read.data.prompt?.title).toBe("Hint");
    }
  });

  it("partial alert/prompt update preserves unspecified host fields", async () => {
    installValidationExcel();
    const adapter = new OfficeJsAdapter();
    const setup = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "I1",
      rule: { type: "list", listValues: ["A", "B"] },
      errorAlert: {
        showAlert: true,
        style: "stop",
        title: "Keep",
        message: "Me",
      },
      prompt: { showPrompt: true, title: "Stay", message: "Here" },
    });
    expect(setup.ok).toBe(true);

    const partial = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "I1",
      rule: { type: "list", listValues: ["X"] },
      errorAlert: { message: "OnlyMsg" },
      prompt: { title: "OnlyTitle" },
    });
    expect(partial.ok).toBe(true);
    if (!partial.ok) return;
    expect(partial.data.errorAlert).toEqual({
      showAlert: true,
      style: "stop",
      title: "Keep",
      message: "OnlyMsg",
    });
    expect(partial.data.prompt).toEqual({
      showPrompt: true,
      title: "OnlyTitle",
      message: "Here",
    });
  });

  it("omitting metadata preserves host values and zero metadata write calls", async () => {
    const ctrl = installValidationExcel();
    const adapter = new OfficeJsAdapter();
    const setup = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "I2",
      rule: { type: "list", listValues: ["A"] },
      errorAlert: { showAlert: true, style: "stop", title: "Keep", message: "Me" },
      prompt: { showPrompt: true, title: "Stay", message: "Here" },
    });
    expect(setup.ok).toBe(true);
    ctrl.resetDvWriteCounts();

    const rewritten = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "I2",
      rule: { type: "list", listValues: ["X", "Y"], allowBlank: true },
    });
    expect(rewritten.ok).toBe(true);
    if (!rewritten.ok) return;
    expect(rewritten.data.errorAlert?.title).toBe("Keep");
    expect(rewritten.data.prompt?.message).toBe("Here");
    const counts = ctrl.getDvWriteCounts();
    expect(counts.errorAlert).toBe(0);
    expect(counts.prompt).toBe(0);
    expect(counts.rule).toBe(1);
  });

  it("pre-read missing errorAlert → ordinary failed, all write counts 0", async () => {
    const ctrl = installValidationExcel({ missingDvErrorAlert: true });
    ctrl.resetDvWriteCounts();
    ctrl.resetSyncCount();
    const failed = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "K1",
      rule: { type: "decimal", operator: "greaterThan", formula1: "0" },
      errorAlert: { showAlert: true, style: "information", message: "x" },
    });
    expect(failed.ok).toBe(false);
    if (!failed.ok) {
      expect(failed.unsupported).not.toBe(true);
      expect(failed.reason).toMatch(/errorAlert/i);
    }
    const counts = ctrl.getDvWriteCounts();
    expect(counts.rule).toBe(0);
    expect(counts.ignoreBlanks).toBe(0);
    expect(counts.errorAlert).toBe(0);
    expect(counts.prompt).toBe(0);
  });

  it("pre-read incomplete errorAlert fields → ordinary failed, zero writes", async () => {
    const ctrl = installValidationExcel({ missingDvErrorAlertFields: true });
    ctrl.resetDvWriteCounts();
    const failed = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "K2",
      rule: { type: "list", listValues: ["a"] },
      errorAlert: { style: "stop" },
    });
    expect(failed.ok).toBe(false);
    if (!failed.ok) {
      expect(failed.unsupported).not.toBe(true);
      expect(failed.reason).toMatch(/errorAlert/i);
    }
    expect(ctrl.getDvWriteCounts()).toEqual({
      rule: 0,
      ignoreBlanks: 0,
      errorAlert: 0,
      prompt: 0,
    });
  });

  it("bad style readback after write → ordinary failed", async () => {
    installValidationExcel({
      tamperDvReadback: { errorAlertStyle: "Not-A-Style" },
    });
    const bad = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "J1",
      rule: { type: "wholeNumber", operator: "equalTo", formula1: "1" },
      errorAlert: { style: "stop", showAlert: true, title: "t", message: "m" },
    });
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(bad.unsupported).not.toBe(true);
      expect(bad.reason).toMatch(/errorAlert\.style|unknown/i);
    }
  });

  it("ExcelApi 1.8 missing → typed unsupported; no Excel.run side effects", async () => {
    const ctrl = installValidationExcel({ excelApi18: false });
    ctrl.resetSyncCount();
    ctrl.resetDvWriteCounts();
    const r = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "L1",
      rule: { type: "list", listValues: ["a"] },
      prompt: { showPrompt: true, title: "t", message: "m" },
    });
    expect(r.ok).toBe(false);
    if (!r.ok) {
      expect(r.unsupported).toBe(true);
      expect(r.reason).toMatch(/ExcelApi 1\.8/);
    }
    expect(ctrl.getSyncCount()).toBe(0);
    expect(ctrl.getDvWriteCounts()).toEqual({
      rule: 0,
      ignoreBlanks: 0,
      errorAlert: 0,
      prompt: 0,
    });
  });

  it("executor rejects empty objects, illegal style, null, length before Host", async () => {
    const host = new MockHostAdapter();
    let calls = 0;
    const orig = host.writeDataValidation.bind(host);
    host.writeDataValidation = async (input) => {
      calls += 1;
      return orig(input);
    };
    const ex = new ToolExecutor(host);

    const empty = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        errorAlert: {},
      },
    });
    expect(empty.ok).toBe(false);
    if (!empty.ok) expect(empty.error).toMatch(/empty/);

    const emptyPrompt = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        prompt: {},
      },
    });
    expect(emptyPrompt.ok).toBe(false);

    const badStyle = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        errorAlert: { style: "critical" },
      },
    });
    expect(badStyle.ok).toBe(false);

    const long = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        prompt: { title: "t".repeat(MAX_DV_ALERT_TITLE_CHARS + 1) },
      },
    });
    expect(long.ok).toBe(false);

    const longMsg = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        errorAlert: { message: "m".repeat(MAX_DV_ALERT_MESSAGE_CHARS + 1) },
      },
    });
    expect(longMsg.ok).toBe(false);
    expect(calls).toBe(0);

    const ok = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        prompt: { title: "", message: "  spaced  " },
      },
    });
    expect(ok.ok).toBe(true);
    expect(calls).toBe(1);
  });

  it("WPS rejects errorAlert/prompt typed unsupported with zero write", async () => {
    installWpsValidationFake();
    const adapter = new WpsJsaAdapter();
    const plain = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "A1",
      rule: { type: "list", listValues: ["ok"] },
    });
    expect(plain.ok).toBe(true);

    const blocked = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "A2",
      rule: { type: "list", listValues: ["x"] },
      errorAlert: { showAlert: true, style: "stop", title: "E", message: "m" },
    });
    expect(blocked.ok).toBe(false);
    if (!blocked.ok) {
      expect(blocked.unsupported).toBe(true);
      expect(blocked.reason).toMatch(/errorAlert\/prompt/);
    }
    const read = await adapter.readDataValidation("Sheet1", "A2");
    expect(read.ok).toBe(true);
    if (read.ok) expect(read.data.rule).toBeNull();
  });

  it("tool count stays 89; schema minProperties on alert objects", () => {
    expect(TOOL_DEFINITIONS).toHaveLength(89);
    const write = TOOL_DEFINITIONS.find((t) => t.name === "dataValidation.write");
    expect(write).toBeTruthy();
    const props = write!.parameters.properties as Record<string, { minProperties?: number }>;
    expect(props.errorAlert?.minProperties).toBe(1);
    expect(props.prompt?.minProperties).toBe(1);
  });
});
