import { DotNetOfficeActionBridge } from "../electron/agent/officeWorker/dotNetOfficeActionBridge";
import { DotNetOfficeDocumentBridge } from "../electron/agent/officeWorker/dotNetOfficeDocumentBridge";
import { DotNetOpenXmlBridge } from "../electron/agent/officeWorker/dotNetOpenXmlBridge";
import {
  disposeOfficeWorkerClient,
  getOfficeWorkerClient,
} from "../electron/agent/officeWorker/officeWorkerClient";
import type {
  OfficeActionApp,
  OfficeActionInput,
  OfficeActionKind,
  OfficeActionResult,
} from "../electron/agent/tools/officeCore/types";

type FileActionInput = Omit<OfficeActionInput, "app" | "action"> & { action?: OfficeActionKind };

process.env.WENGGE_OFFICE_SMOKE = "1";

const openXmlBridge = new DotNetOpenXmlBridge();

export { DotNetOfficeActionBridge, DotNetOfficeDocumentBridge };

export function applyExcelAdvancedAction(input: FileActionInput): Promise<OfficeActionResult> {
  return executeFileAction("excel", input);
}

export function applyWordAdvancedAction(input: FileActionInput): Promise<OfficeActionResult> {
  return executeFileAction("word", input);
}

export function applyPresentationAdvancedAction(
  input: FileActionInput,
): Promise<OfficeActionResult> {
  return executeFileAction("presentation", input);
}

export function markWordBookmarkDirty(input: {
  filePath: string;
  instanceId: string;
  name: string;
  text: string;
}): Promise<{ saved: boolean; instanceId: string }> {
  return getOfficeWorkerClient().invoke("office.smoke.markWordBookmarkDirty", input);
}

export function openOfficeFixtures(input: {
  excelPaths: string[];
  wordPaths: string[];
  presentationPaths: string[];
}): Promise<{ excel: number[]; word: number[]; presentation: number[] }> {
  return getOfficeWorkerClient().invoke("office.smoke.openFixtures", input, 120_000);
}

export function closeOfficeFixtures(input: {
  excel: number[];
  word: number[];
  presentation: number[];
}): Promise<{ closed: number[] }> {
  return getOfficeWorkerClient().invoke("office.smoke.closeFixtures", input, 30_000);
}

export function listOfficeSmokeProcesses(): Promise<{ microsoft: number[]; wpsVisible: number[] }> {
  return getOfficeWorkerClient().invoke("office.smoke.listProcesses");
}

export function runningOfficeSmokeProcesses(ids: number[]): Promise<number[]> {
  return getOfficeWorkerClient().invoke("office.smoke.runningProcesses", { ids });
}

export function disposeOfficeWorker(): Promise<void> {
  return disposeOfficeWorkerClient();
}

function executeFileAction(
  app: OfficeActionApp,
  input: FileActionInput,
): Promise<OfficeActionResult> {
  return openXmlBridge.executeAction!({
    ...input,
    app,
    action: input.action || defaultAction(input.operation),
  });
}

function defaultAction(operation: string): OfficeActionKind {
  if (
    operation.startsWith("create") ||
    operation.startsWith("add") ||
    operation.startsWith("append")
  )
    return "insert";
  if (
    operation.startsWith("apply") ||
    operation.startsWith("style") ||
    operation === "setHeaderFooter"
  )
    return "style";
  return "edit";
}
