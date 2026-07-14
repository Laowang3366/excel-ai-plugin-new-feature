import { DotNetExcelBridge } from "../officeWorker/dotNetExcelBridge";
import { DotNetJsaBridge, DotNetUiBridge, DotNetVbaBridge } from "../officeWorker/dotNetMacroBridges";
import { DotNetPresentationBridge, DotNetWordBridge } from "../officeWorker/dotNetDocumentBridges";
import { DotNetOpenXmlBridge } from "../officeWorker/dotNetOpenXmlBridge";
import { disposeOfficeWorkerClient } from "../officeWorker/officeWorkerClient";
import { createLogger } from "../../shared/logger";

const bridgeRegistryLogger = createLogger("BridgeRegistry");

export interface OfficeBridgeRegistry {
  excelBridge: DotNetExcelBridge;
  vbaBridge: DotNetVbaBridge;
  jsaBridge: DotNetJsaBridge;
  uiBridge: DotNetUiBridge;
  wordBridge: DotNetWordBridge;
  presentationBridge: DotNetPresentationBridge;
  officeFileBridge: DotNetOpenXmlBridge;
}

let excelBridge: DotNetExcelBridge | null = null;
let vbaBridge: DotNetVbaBridge | null = null;
let jsaBridge: DotNetJsaBridge | null = null;
let uiBridge: DotNetUiBridge | null = null;
let wordBridge: DotNetWordBridge | null = null;
let presentationBridge: DotNetPresentationBridge | null = null;
let officeFileBridge: DotNetOpenXmlBridge | null = null;

/**
 * Office 桥接实例注册表。
 *
 * 关联模块：
 * - officeWorker/*: .NET 8 COM 与 Open XML Worker 的 TypeScript 薄桥。
 * - main-modules/ipcHandlers: Excel 连接状态和手动连接复用这里的实例。
 *
 * 主进程运行期 intentionally 复用同一组 bridge，避免每个 IPC 请求重复创建 COM 门面。
 * resetOfficeBridgeRegistry 用于会话/测试清理边界，不在工具执行中途调用。
 */
export function getOrCreateExcelBridge(): DotNetExcelBridge {
  if (!excelBridge) {
    excelBridge = new DotNetExcelBridge();
  }
  return excelBridge;
}

export function getOrCreateOfficeBridges(): OfficeBridgeRegistry {
  const activeExcelBridge = getOrCreateExcelBridge();
  if (!vbaBridge) {
    vbaBridge = new DotNetVbaBridge();
  }
  if (!jsaBridge) {
    jsaBridge = new DotNetJsaBridge();
  }
  if (!uiBridge) {
    uiBridge = new DotNetUiBridge();
  }
  if (!wordBridge) {
    wordBridge = new DotNetWordBridge();
  }
  if (!presentationBridge) {
    presentationBridge = new DotNetPresentationBridge();
  }
  if (!officeFileBridge) {
    officeFileBridge = new DotNetOpenXmlBridge();
  }

  return {
    excelBridge: activeExcelBridge,
    vbaBridge,
    jsaBridge,
    uiBridge,
    wordBridge,
    presentationBridge,
    officeFileBridge,
  };
}

export function getExcelBridge(): DotNetExcelBridge | null {
  return excelBridge;
}

export function getWordBridge(): DotNetWordBridge | null {
  return wordBridge;
}

export function getPresentationBridge(): DotNetPresentationBridge | null {
  return presentationBridge;
}

export function getVbaBridge(): DotNetVbaBridge | null {
  return vbaBridge;
}

export function setExcelBridgeInstance(bridge: DotNetExcelBridge | null): void {
  excelBridge = bridge;
  vbaBridge = null;
  jsaBridge = null;
  uiBridge = null;
}

export function resetOfficeBridgeRegistry(): void {
  excelBridge = null;
  vbaBridge = null;
  jsaBridge = null;
  uiBridge = null;
  wordBridge = null;
  presentationBridge = null;
  officeFileBridge = null;
}

export async function disconnectOfficeBridges(): Promise<void> {
  const cleanupTasks: Array<Promise<unknown>> = [];
  if (excelBridge) cleanupTasks.push(excelBridge.disconnect());
  if (wordBridge?.isConnected()) cleanupTasks.push(wordBridge.saveDocument());
  if (presentationBridge?.isConnected()) cleanupTasks.push(presentationBridge.savePresentation());

  const results = await Promise.allSettled(cleanupTasks);
  for (const result of results) {
    if (result.status === "rejected") {
      bridgeRegistryLogger.warn("Office bridge cleanup failed", result.reason instanceof Error
        ? { message: result.reason.message, stack: result.reason.stack }
        : { reason: String(result.reason) });
    }
  }
  await disposeOfficeWorkerClient();
}
