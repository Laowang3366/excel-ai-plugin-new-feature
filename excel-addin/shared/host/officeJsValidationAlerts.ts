/**
 * Office.js DataValidation errorAlert / prompt helpers (ExcelApi 1.8).
 * Official ClientObject contract:
 * - load top-level: type, rule, ignoreBlanks, errorAlert, prompt
 * - assign whole errorAlert/prompt plain objects (not nested field mutation)
 * @see https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidation
 * @see https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidationerroralert
 * @see https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidationprompt
 * @see https://learn.microsoft.com/en-us/javascript/api/excel/excel.datavalidationalertstyle
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

/** Official DataValidationLoadOptions top-level only (no nested paths). */
export const DV_FULL_LOAD_PROPS = "type,rule,ignoreBlanks,errorAlert,prompt";

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

/** Complete public snapshot after load (all official fields present). */
export interface HostErrorAlertSnapshot {
  showAlert: boolean;
  style: DataValidationAlertStyle;
  title: string;
  message: string;
}

export interface HostPromptSnapshot {
  showPrompt: boolean;
  title: string;
  message: string;
}

/** Host-shaped whole-object assignment payload (PascalCase style tokens). */
export interface HostErrorAlertWrite {
  showAlert: boolean;
  style: string;
  title: string;
  message: string;
}

export interface HostPromptWrite {
  showPrompt: boolean;
  title: string;
  message: string;
}

export function mapAlertStyleToHost(style: DataValidationAlertStyle): string {
  return STYLE_TO_HOST[style];
}

/**
 * Case-insensitive exact official tokens only.
 * Lower-case only — no trim, no whitespace/punctuation stripping, no aliases.
 */
export function unmapAlertStyle(host: unknown): DataValidationAlertStyle | undefined {
  if (typeof host !== "string") return undefined;
  return HOST_STYLE_TO_PUBLIC[host.toLowerCase()];
}

export function isDataValidationAlertStyle(
  value: unknown,
): value is DataValidationAlertStyle {
  return typeof value === "string" && (DV_ALERT_STYLES as readonly string[]).includes(value);
}

/** Strict complete host errorAlert; null/empty/missing field → error. */
export function parseErrorAlertFromHost(
  raw: unknown,
): { value: HostErrorAlertSnapshot | null; error?: string } {
  if (raw == null) {
    return { value: null, error: "errorAlert host value is null/undefined" };
  }
  if (typeof raw !== "object" || Array.isArray(raw)) {
    return { value: null, error: "errorAlert host value is not an object" };
  }
  const bag = raw as Record<string, unknown>;
  if (typeof bag.showAlert !== "boolean") {
    return { value: null, error: "errorAlert.showAlert host readback is not boolean" };
  }
  if (typeof bag.style !== "string") {
    return { value: null, error: "errorAlert.style host readback is not string" };
  }
  const style = unmapAlertStyle(bag.style);
  if (!style) {
    return {
      value: null,
      error: `errorAlert.style host readback unknown: ${String(bag.style)}`,
    };
  }
  if (typeof bag.title !== "string") {
    return { value: null, error: "errorAlert.title host readback is not string" };
  }
  if (typeof bag.message !== "string") {
    return { value: null, error: "errorAlert.message host readback is not string" };
  }
  return {
    value: {
      showAlert: bag.showAlert,
      style,
      title: bag.title,
      message: bag.message,
    },
  };
}

/** Strict complete host prompt; null/empty/missing field → error. */
export function parsePromptFromHost(
  raw: unknown,
): { value: HostPromptSnapshot | null; error?: string } {
  if (raw == null) {
    return { value: null, error: "prompt host value is null/undefined" };
  }
  if (typeof raw !== "object" || Array.isArray(raw)) {
    return { value: null, error: "prompt host value is not an object" };
  }
  const bag = raw as Record<string, unknown>;
  if (typeof bag.showPrompt !== "boolean") {
    return { value: null, error: "prompt.showPrompt host readback is not boolean" };
  }
  if (typeof bag.title !== "string") {
    return { value: null, error: "prompt.title host readback is not string" };
  }
  if (typeof bag.message !== "string") {
    return { value: null, error: "prompt.message host readback is not string" };
  }
  return {
    value: {
      showPrompt: bag.showPrompt,
      title: bag.title,
      message: bag.message,
    },
  };
}


/** Official Excel.DataValidationType host tokens (ExcelApi 1.8). */
const OFFICIAL_DV_HOST_TYPES: Record<string, string> = {
  none: "None",
  wholenumber: "WholeNumber",
  decimal: "Decimal",
  list: "List",
  date: "Date",
  time: "Time",
  textlength: "TextLength",
  custom: "Custom",
  inconsistent: "Inconsistent",
  mixedcriteria: "MixedCriteria",
};

/**
 * Runtime type must be an official DataValidationType token.
 * Case-insensitive exact match only (lower-case); no trim / space collapse / aliases.
 * null/undefined/number/unknown/"Whole Number" → ordinary failed.
 */
export function assertOfficialDvHostType(raw: unknown): string {
  if (raw === null || raw === undefined) {
    throw new Error("dataValidation.type host readback is null/undefined");
  }
  if (typeof raw !== "string") {
    throw new Error(
      `dataValidation.type host readback is not string (got ${typeof raw})`,
    );
  }
  const canon = OFFICIAL_DV_HOST_TYPES[raw.toLowerCase()];
  if (!canon) {
    throw new Error(`dataValidation.type host readback unknown: ${raw}`);
  }
  return canon;
}

export type LoadedDvSurface = {
  type: string;
  rule: unknown;
  ignoreBlanks: boolean;
  errorAlert: HostErrorAlertSnapshot;
  prompt: HostPromptSnapshot;
};

/**
 * After load+sync only. Validates official surface; missing/null/bad → throw ordinary failed.
 * Must not run before first load/sync (PropertyNotLoaded).
 */
export function assertLoadedDvSurface(dv: ExcelDataValidation): LoadedDvSurface {
  if (typeof dv.load !== "function") {
    throw new Error("dataValidation.load is missing");
  }
  let rawType: unknown;
  let rule: unknown;
  let ignoreBlanks: boolean;
  try {
    rawType = dv.type;
  } catch (err) {
    throw new Error(
      `dataValidation.type PropertyNotLoaded: ${err instanceof Error ? err.message : String(err)}`,
    );
  }
  const type = assertOfficialDvHostType(rawType);
  try {
    rule = dv.rule;
  } catch (err) {
    throw new Error(
      `dataValidation.rule PropertyNotLoaded: ${err instanceof Error ? err.message : String(err)}`,
    );
  }
  try {
    ignoreBlanks = dv.ignoreBlanks;
  } catch (err) {
    throw new Error(
      `dataValidation.ignoreBlanks PropertyNotLoaded: ${err instanceof Error ? err.message : String(err)}`,
    );
  }
  if (typeof ignoreBlanks !== "boolean") {
    throw new Error("dataValidation.ignoreBlanks host readback is not boolean");
  }
  if (rule === null || typeof rule !== "object") {
    throw new Error("dataValidation.rule host readback is not an object");
  }

  let rawError: unknown;
  let rawPrompt: unknown;
  try {
    rawError = dv.errorAlert;
  } catch (err) {
    throw new Error(
      `dataValidation.errorAlert PropertyNotLoaded: ${err instanceof Error ? err.message : String(err)}`,
    );
  }
  try {
    rawPrompt = dv.prompt;
  } catch (err) {
    throw new Error(
      `dataValidation.prompt PropertyNotLoaded: ${err instanceof Error ? err.message : String(err)}`,
    );
  }

  const ea = parseErrorAlertFromHost(rawError);
  if (ea.error || !ea.value) {
    throw new Error(ea.error ?? "dataValidation.errorAlert incomplete");
  }
  const pr = parsePromptFromHost(rawPrompt);
  if (pr.error || !pr.value) {
    throw new Error(pr.error ?? "dataValidation.prompt incomplete");
  }
  return {
    type,
    rule,
    ignoreBlanks,
    errorAlert: ea.value,
    prompt: pr.value,
  };
}

/** Merge request partial onto loaded host snapshot → host write shape. */
export function mergeErrorAlertForWrite(
  host: HostErrorAlertSnapshot,
  partial: DataValidationErrorAlert,
): HostErrorAlertWrite {
  const style = partial.style ?? host.style;
  return {
    showAlert: partial.showAlert ?? host.showAlert,
    style: mapAlertStyleToHost(style),
    title: partial.title ?? host.title,
    message: partial.message ?? host.message,
  };
}

export function mergePromptForWrite(
  host: HostPromptSnapshot,
  partial: DataValidationPrompt,
): HostPromptWrite {
  return {
    showPrompt: partial.showPrompt ?? host.showPrompt,
    title: partial.title ?? host.title,
    message: partial.message ?? host.message,
  };
}

/** Whole-object assignment only (official scalar snapshot write). */
export function assignErrorAlert(dv: ExcelDataValidation, value: HostErrorAlertWrite): void {
  dv.errorAlert = {
    showAlert: value.showAlert,
    style: value.style,
    title: value.title,
    message: value.message,
  };
}

export function assignPrompt(dv: ExcelDataValidation, value: HostPromptWrite): void {
  dv.prompt = {
    showPrompt: value.showPrompt,
    title: value.title,
    message: value.message,
  };
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
