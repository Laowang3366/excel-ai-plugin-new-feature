/**
 * Office.js DataValidation errorAlert / prompt helpers (ExcelApi 1.8).
 * Official:
 * - https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidation
 * - https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidationerroralert
 * - https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidationprompt
 * - https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidationalertstyle
 */
import type { ExcelDataValidation } from "./officeJsExcelTypes";
import type {
  DataValidationAlertStyle,
  DataValidationErrorAlert,
  DataValidationPrompt,
} from "./types";

/** Explicit schema/runtime cap (Excel UI historically ~255 for messages). */
export const MAX_DV_ALERT_TITLE_CHARS = 255;
export const MAX_DV_ALERT_MESSAGE_CHARS = 255;

export const DV_ALERT_STYLES: readonly DataValidationAlertStyle[] = [
  "stop",
  "warning",
  "information",
] as const;

const STYLE_TO_HOST: Record<DataValidationAlertStyle, string> = {
  stop: "Stop",
  warning: "Warning",
  information: "Information",
};

const HOST_STYLE_TO_PUBLIC: Record<string, DataValidationAlertStyle> = {
  stop: "stop",
  warning: "warning",
  information: "information",
};

/** Nested load paths for type/rule/ignoreBlanks + errorAlert + prompt. */
export const DV_FULL_LOAD_PROPS =
  "type,rule,ignoreBlanks," +
  "errorAlert/showAlert,errorAlert/style,errorAlert/title,errorAlert/message," +
  "prompt/showPrompt,prompt/title,prompt/message";

export function mapAlertStyleToHost(style: DataValidationAlertStyle): string {
  return STYLE_TO_HOST[style];
}

/** Case-insensitive exact official tokens only — no punctuation stripping. */
export function unmapAlertStyle(host: unknown): DataValidationAlertStyle | undefined {
  if (typeof host !== "string" || host.trim() === "") return undefined;
  return HOST_STYLE_TO_PUBLIC[host.trim().toLowerCase()];
}

export function isDataValidationAlertStyle(
  value: unknown,
): value is DataValidationAlertStyle {
  return typeof value === "string" && (DV_ALERT_STYLES as readonly string[]).includes(value);
}

type AlertBag = {
  showAlert?: boolean;
  style?: string;
  title?: string;
  message?: string;
};

type PromptBag = {
  showPrompt?: boolean;
  title?: string;
  message?: string;
};

export function parseErrorAlertFromHost(
  raw: unknown,
): { value: DataValidationErrorAlert | null; error?: string } {
  if (raw == null) return { value: null };
  if (typeof raw !== "object" || Array.isArray(raw)) {
    return { value: null, error: "errorAlert host value is not an object" };
  }
  const bag = raw as AlertBag;
  const out: DataValidationErrorAlert = {};
  if (bag.showAlert !== undefined) {
    if (typeof bag.showAlert !== "boolean") {
      return { value: null, error: "errorAlert.showAlert host readback is not boolean" };
    }
    out.showAlert = bag.showAlert;
  }
  if (bag.style !== undefined && bag.style !== null && bag.style !== "") {
    const mapped = unmapAlertStyle(bag.style);
    if (!mapped) {
      return {
        value: null,
        error: `errorAlert.style host readback unknown: ${String(bag.style)}`,
      };
    }
    out.style = mapped;
  }
  if (bag.title !== undefined && bag.title !== null) {
    if (typeof bag.title !== "string") {
      return { value: null, error: "errorAlert.title host readback is not string" };
    }
    out.title = bag.title;
  }
  if (bag.message !== undefined && bag.message !== null) {
    if (typeof bag.message !== "string") {
      return { value: null, error: "errorAlert.message host readback is not string" };
    }
    out.message = bag.message;
  }
  return { value: out };
}

export function parsePromptFromHost(
  raw: unknown,
): { value: DataValidationPrompt | null; error?: string } {
  if (raw == null) return { value: null };
  if (typeof raw !== "object" || Array.isArray(raw)) {
    return { value: null, error: "prompt host value is not an object" };
  }
  const bag = raw as PromptBag;
  const out: DataValidationPrompt = {};
  if (bag.showPrompt !== undefined) {
    if (typeof bag.showPrompt !== "boolean") {
      return { value: null, error: "prompt.showPrompt host readback is not boolean" };
    }
    out.showPrompt = bag.showPrompt;
  }
  if (bag.title !== undefined && bag.title !== null) {
    if (typeof bag.title !== "string") {
      return { value: null, error: "prompt.title host readback is not string" };
    }
    out.title = bag.title;
  }
  if (bag.message !== undefined && bag.message !== null) {
    if (typeof bag.message !== "string") {
      return { value: null, error: "prompt.message host readback is not string" };
    }
    out.message = bag.message;
  }
  return { value: out };
}

/** Fail before any write if required members for requested fields are missing. */
export function requireDvAlertMembers(
  dv: ExcelDataValidation,
  wantsErrorAlert: boolean,
  wantsPrompt: boolean,
): void {
  if (typeof dv.load !== "function") {
    throw new Error("dataValidation.load is missing");
  }
  // Always need ignoreBlanks + rule writable surface (present on real Office.js DV).
  try {
    // Accessors must exist; reading is fine even when unloaded for our fake/precheck.
    void dv.ignoreBlanks;
  } catch {
    throw new Error("dataValidation.ignoreBlanks is missing");
  }
  try {
    void dv.rule;
  } catch {
    throw new Error("dataValidation.rule is missing");
  }
  if (wantsErrorAlert) {
    if (!dv.errorAlert || typeof dv.errorAlert !== "object") {
      throw new Error("dataValidation.errorAlert is missing");
    }
    const ea = dv.errorAlert as AlertBag & object;
    for (const key of ["showAlert", "style", "title", "message"] as const) {
      if (!(key in ea)) {
        throw new Error(`dataValidation.errorAlert.${key} is missing`);
      }
    }
  }
  if (wantsPrompt) {
    if (!dv.prompt || typeof dv.prompt !== "object") {
      throw new Error("dataValidation.prompt is missing");
    }
    const p = dv.prompt as PromptBag & object;
    for (const key of ["showPrompt", "title", "message"] as const) {
      if (!(key in p)) {
        throw new Error(`dataValidation.prompt.${key} is missing`);
      }
    }
  }
}

export function applyErrorAlert(
  dv: ExcelDataValidation,
  alert: DataValidationErrorAlert,
): void {
  if (!dv.errorAlert || typeof dv.errorAlert !== "object") {
    throw new Error("dataValidation.errorAlert is missing");
  }
  const ea = dv.errorAlert;
  if (alert.showAlert !== undefined) ea.showAlert = alert.showAlert;
  if (alert.style !== undefined) ea.style = mapAlertStyleToHost(alert.style);
  if (alert.title !== undefined) ea.title = alert.title;
  if (alert.message !== undefined) ea.message = alert.message;
}

export function applyPrompt(dv: ExcelDataValidation, prompt: DataValidationPrompt): void {
  if (!dv.prompt || typeof dv.prompt !== "object") {
    throw new Error("dataValidation.prompt is missing");
  }
  const p = dv.prompt;
  if (prompt.showPrompt !== undefined) p.showPrompt = prompt.showPrompt;
  if (prompt.title !== undefined) p.title = prompt.title;
  if (prompt.message !== undefined) p.message = prompt.message;
}

function sameOptBool(a: boolean | undefined, b: boolean | undefined): boolean {
  if (a === undefined) return true;
  return a === b;
}

function sameOptStr(a: string | undefined, b: string | undefined): boolean {
  if (a === undefined) return true;
  return a === b;
}

/** Match only fields present on the request (partial write contract). */
export function errorAlertMatches(
  expected: DataValidationErrorAlert,
  actual: DataValidationErrorAlert | null | undefined,
): boolean {
  if (!actual) return false;
  if (!sameOptBool(expected.showAlert, actual.showAlert)) return false;
  if (expected.style !== undefined && expected.style !== actual.style) return false;
  if (!sameOptStr(expected.title, actual.title)) return false;
  if (!sameOptStr(expected.message, actual.message)) return false;
  return true;
}

export function promptMatches(
  expected: DataValidationPrompt,
  actual: DataValidationPrompt | null | undefined,
): boolean {
  if (!actual) return false;
  if (!sameOptBool(expected.showPrompt, actual.showPrompt)) return false;
  if (!sameOptStr(expected.title, actual.title)) return false;
  if (!sameOptStr(expected.message, actual.message)) return false;
  return true;
}
