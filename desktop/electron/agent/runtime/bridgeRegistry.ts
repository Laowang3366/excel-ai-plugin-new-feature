import { ExcelComBridge } from "../tools/implementations/excel/excelComBridge";
import { ExcelScriptBridgeCom } from "../tools/implementations/excel/excelScriptBridgeCom";
import { ExcelUiComBridge } from "../tools/implementations/excel/excelUiComBridge";
import { ExcelVbaComBridge } from "../tools/implementations/excel/excelVbaComBridge";
import { OfficeScriptBridge } from "../tools/implementations/office/officeScriptBridge";
import { PresentationComBridge } from "../tools/implementations/office/presentationComBridge";
import { WordComBridge } from "../tools/implementations/office/wordComBridge";
import { OfficeOpenXmlFileBridge } from "../tools/implementations/officeOpenXml/officeOpenXmlFileBridge";
import type { ExcelConnectionBridge } from "../tools/contracts/excel";
import { createLogger } from "../../shared/logger";

const bridgeRegistryLogger = createLogger("BridgeRegistry");

export interface OfficeBridgeRegistry {
  excelBridge: ExcelComBridge;
  vbaBridge: ExcelVbaComBridge;
  scriptBridge: ExcelScriptBridgeCom;
  uiBridge: ExcelUiComBridge;
  wordBridge: WordComBridge;
  presentationBridge: PresentationComBridge;
  officeScriptBridge: OfficeScriptBridge;
  officeFileBridge: OfficeOpenXmlFileBridge;
}

let excelBridge: ExcelComBridge | null = null;
let vbaBridge: ExcelVbaComBridge | null = null;
let scriptBridge: ExcelScriptBridgeCom | null = null;
let uiBridge: ExcelUiComBridge | null = null;
let wordBridge: WordComBridge | null = null;
let presentationBridge: PresentationComBridge | null = null;
let officeScriptBridge: OfficeScriptBridge | null = null;
let officeFileBridge: OfficeOpenXmlFileBridge | null = null;

/**
 * Office 桥接实例注册表。
 *
 * 关联模块：
 * - tools/implementations/excel/*Bridge: Excel/WPS COM 操作桥接。
 * - tools/implementations/office/*Bridge: Word/PPT/Office 脚本桥接。
 * - tools/implementations/officeOpenXml/*: docx/pptx/xlsx 文件级编辑桥接。
 * - main-modules/ipcHandlers: Excel 连接状态和手动连接复用这里的实例。
 *
 * 主进程运行期 intentionally 复用同一组 bridge，避免每个 IPC 请求重复创建 COM 门面。
 * resetOfficeBridgeRegistry 用于会话/测试清理边界，不在工具执行中途调用。
 */
export function getOrCreateExcelBridge(): ExcelConnectionBridge {
  if (!excelBridge) {
    excelBridge = new ExcelComBridge();
  }
  return excelBridge;
}

export function getOrCreateOfficeBridges(): OfficeBridgeRegistry {
  const activeExcelBridge = getOrCreateExcelBridge() as ExcelComBridge;
  if (!vbaBridge) {
    vbaBridge = new ExcelVbaComBridge(activeExcelBridge);
  }
  if (!scriptBridge) {
    scriptBridge = new ExcelScriptBridgeCom(activeExcelBridge);
  }
  if (!uiBridge) {
    uiBridge = new ExcelUiComBridge(activeExcelBridge);
  }
  if (!wordBridge) {
    wordBridge = new WordComBridge();
  }
  if (!presentationBridge) {
    presentationBridge = new PresentationComBridge();
  }
  if (!officeScriptBridge) {
    officeScriptBridge = new OfficeScriptBridge();
  }
  if (!officeFileBridge) {
    officeFileBridge = new OfficeOpenXmlFileBridge();
  }

  return {
    excelBridge: activeExcelBridge,
    vbaBridge,
    scriptBridge,
    uiBridge,
    wordBridge,
    presentationBridge,
    officeScriptBridge,
    officeFileBridge,
  };
}

export function getExcelBridge(): ExcelComBridge | null {
  return excelBridge;
}

export function getWordBridge(): WordComBridge | null {
  return wordBridge;
}

export function getPresentationBridge(): PresentationComBridge | null {
  return presentationBridge;
}

export function getVbaBridge(): ExcelVbaComBridge | null {
  return vbaBridge;
}

export function setExcelBridgeInstance(bridge: ExcelComBridge | null): void {
  excelBridge = bridge;
  vbaBridge = null;
  scriptBridge = null;
  uiBridge = null;
}

export function resetOfficeBridgeRegistry(): void {
  excelBridge = null;
  vbaBridge = null;
  scriptBridge = null;
  uiBridge = null;
  wordBridge = null;
  presentationBridge = null;
  officeScriptBridge = null;
  officeFileBridge = null;
}

export async function disconnectOfficeBridges(): Promise<void> {
  const cleanupTasks: Array<Promise<unknown>> = [];
  if (excelBridge) cleanupTasks.push(excelBridge.disconnect());
  if (wordBridge) cleanupTasks.push(wordBridge.saveDocument());
  if (presentationBridge) cleanupTasks.push(presentationBridge.savePresentation());

  const results = await Promise.allSettled(cleanupTasks);
  for (const result of results) {
    if (result.status === "rejected") {
      bridgeRegistryLogger.warn("Office bridge cleanup failed", result.reason instanceof Error
        ? { message: result.reason.message, stack: result.reason.stack }
        : { reason: String(result.reason) });
    }
  }
}
