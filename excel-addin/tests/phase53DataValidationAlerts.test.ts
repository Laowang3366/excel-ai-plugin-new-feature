/**
 * Phase53: dataValidation errorAlert / prompt / ignoreBlanks(allowBlank) Office.js contract.
 */
import { afterEach, describe, expect, it } from "vitest";
import { OfficeJsAdapter } from "../shared/host/officeJsAdapter";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { TOOL_DEFINITIONS } from "../shared/tools/definitions";
import { ToolExecutor } from "../shared/tools/executor";
import {
  MAX_DV_ALERT_MESSAGE_CHARS,
  MAX_DV_ALERT_TITLE_CHARS,
  mapAlertStyleToHost,
  unmapAlertStyle,
} from "../shared/host/officeJsValidationAlerts";
import { installValidationExcel } from "./fakes/officeJsValidationFake";
import { installWpsValidationFake } from "./fakes/wpsJsaValidationFake";
import { MockHostAdapter } from "./mockHost";

describe("phase53 dataValidation alerts/prompt", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
    delete (globalThis as { Office?: unknown }).Office;
    delete (globalThis as { Application?: unknown }).Application;
  });

  it("maps alert styles bidirectionally (official Stop/Warning/Information)", () => {
    expect(mapAlertStyleToHost("stop")).toBe("Stop");
    expect(mapAlertStyleToHost("warning")).toBe("Warning");
    expect(mapAlertStyleToHost("information")).toBe("Information");
    expect(unmapAlertStyle("Stop")).toBe("stop");
    expect(unmapAlertStyle("WARNING")).toBe("warning");
    expect(unmapAlertStyle("Information")).toBe("information");
    expect(unmapAlertStyle("Stop-Alert")).toBeUndefined();
    expect(unmapAlertStyle("In formation")).toBeUndefined();
  });

  it("round-trips errorAlert + prompt + allowBlank with real host readback", async () => {
    installValidationExcel();
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
      prompt: {
        showPrompt: true,
        title: "Hint",
        message: "Whole number",
      },
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

    const read = await adapter.readDataValidation("Sheet1", "H1:H3");
    expect(read.ok).toBe(true);
    if (!read.ok) return;
    expect(read.data.errorAlert?.style).toBe("warning");
    expect(read.data.prompt?.title).toBe("Hint");
    expect(read.data.rule?.allowBlank).toBe(false);
  });

  it("preserves host prompt/errorAlert when write omits them", async () => {
    installValidationExcel();
    const adapter = new OfficeJsAdapter();
    const setup = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "I1",
      rule: { type: "list", listValues: ["A", "B"] },
      errorAlert: { showAlert: true, style: "stop", title: "Keep", message: "Me" },
      prompt: { showPrompt: true, title: "Stay", message: "Here" },
    });
    expect(setup.ok).toBe(true);

    const rewritten = await adapter.writeDataValidation({
      sheetName: "Sheet1",
      range: "I1",
      rule: { type: "list", listValues: ["X", "Y"], allowBlank: true },
      // intentionally omit errorAlert/prompt — must not clear host metadata
    });
    expect(rewritten.ok).toBe(true);
    if (!rewritten.ok) return;
    expect(rewritten.data.rule?.listValues).toEqual(["X", "Y"]);
    expect(rewritten.data.errorAlert?.title).toBe("Keep");
    expect(rewritten.data.prompt?.message).toBe("Here");
  });

  it("fails closed on bad errorAlert style readback without echoing request", async () => {
    installValidationExcel({
      tamperDvReadback: { errorAlertStyle: "Not-A-Style" },
    });
    const bad = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "J1",
      rule: { type: "wholeNumber", operator: "equalTo", formula1: "1" },
      errorAlert: { style: "stop", showAlert: true },
    });
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(bad.unsupported).not.toBe(true);
      expect(bad.reason).toMatch(/errorAlert\.style|unknown/i);
    }
  });

  it("member precheck: missing errorAlert is ordinary failed with zero partial rule when wanted", async () => {
    const ctrl = installValidationExcel({ missingDvErrorAlert: true });
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
      expect(failed.reason).toMatch(/errorAlert is missing/);
    }
    // Excel.run still entered but precheck throws before rule write — sync may be 0
    expect(ctrl.getSyncCount()).toBe(0);

    // Without errorAlert, rule-only write still works when errorAlert member absent
    const okRule = await new OfficeJsAdapter().writeDataValidation({
      sheetName: "Sheet1",
      range: "K2",
      rule: { type: "decimal", operator: "greaterThan", formula1: "0" },
    });
    expect(okRule.ok).toBe(true);
  });

  it("ExcelApi 1.8 missing → typed unsupported for write with alerts", async () => {
    installValidationExcel({ excelApi18: false });
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
  });

  it("executor rejects illegal style/null/unknown/length before Host", async () => {
    const host = new MockHostAdapter();
    const spy = viSpy(host);
    const ex = new ToolExecutor(host);

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
    if (!badStyle.ok) expect(badStyle.error).toMatch(/stop\|warning\|information/);

    const nullAlert = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        errorAlert: null,
      },
    });
    expect(nullAlert.ok).toBe(false);

    const unknown = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        errorAlert: { boom: true },
      },
    });
    expect(unknown.ok).toBe(false);
    if (!unknown.ok) expect(unknown.error).toMatch(/unknown errorAlert field/);

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
    if (!long.ok) expect(long.error).toMatch(/maxLength/);

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

    expect(spy.calls).toBe(0);

    const okEmpty = await ex.execute({
      name: "dataValidation.write",
      arguments: {
        sheetName: "Sheet1",
        range: "A1",
        rule: { type: "list", listValues: ["a"] },
        errorAlert: {},
        prompt: { title: "", message: "  spaced  " },
      },
    });
    expect(okEmpty.ok).toBe(true);
    if (okEmpty.ok) {
      expect((okEmpty.data as { prompt?: { message?: string } }).prompt?.message).toBe(
        "  spaced  ",
      );
    }
    expect(spy.calls).toBe(1);
  });

  it("WPS rejects errorAlert/prompt with typed unsupported and zero write side effects", async () => {
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
    if (read.ok) {
      // zero write side effects: no rule on A2
      expect(read.data.rule).toBeNull();
    }
  });

  it("keeps tool count at 89; schema documents errorAlert/prompt", () => {
    expect(TOOL_DEFINITIONS).toHaveLength(89);
    const write = TOOL_DEFINITIONS.find((t) => t.name === "dataValidation.write");
    expect(write).toBeTruthy();
    const props = write!.parameters.properties as Record<string, unknown>;
    expect(props.errorAlert).toBeTruthy();
    expect(props.prompt).toBeTruthy();
    expect(write!.description).toMatch(/errorAlert/);
    expect(write!.description).not.toMatch(/本批不实现/);
  });
});

function viSpy(host: MockHostAdapter) {
  let calls = 0;
  const orig = host.writeDataValidation.bind(host);
  host.writeDataValidation = async (input) => {
    calls += 1;
    return orig(input);
  };
  return {
    get calls() {
      return calls;
    },
  };
}
