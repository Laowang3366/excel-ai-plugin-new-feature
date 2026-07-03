/**
 * Excel/WPS 工作表管理能力。
 *
 * 关联模块：
 * - excelComBridge.ts: 对外保留 sheetOperation 门面。
 * - tools/executors/excelExecutors.ts: sheet.operation 工具调用本能力。
 */

import { executePowerShell, psVar } from "../../../automation/powershell";
import type { SpreadsheetHost } from "./connectionMetadata";

export interface SheetOperationDeps {
  ensureConnected: () => Promise<SpreadsheetHost | null>;
  getProgId: () => string;
}

export async function sheetOperation(
  deps: SheetOperationDeps,
  operation: string,
  sheetName: string,
  options?: Record<string, unknown>
): Promise<unknown> {
  const host = await deps.ensureConnected();
  if (!host) throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");

  const progId = deps.getProgId();
  const varSheetName = psVar("_sheetName", sheetName);

  switch (operation) {
    case "rename": {
      const newName = (options?.newName as string) || "";
      if (!newName) throw new Error("重命名操作需要提供 newName 参数");
      await executePowerShell(`
        ${varSheetName}
        ${psVar("_newName", newName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        $ws = $wb.Sheets.Item($_sheetName)
        $ws.Name = $_newName
      `);
      return `工作表已重命名为 ${newName}`;
    }
    case "delete": {
      await executePowerShell(`
        ${varSheetName}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        $ws = $wb.Sheets.Item($_sheetName)
        $ws.Delete()
      `);
      return `工作表 ${sheetName} 已删除`;
    }
    case "add": {
      const addName = (options?.newName as string) || `Sheet${Date.now() % 10000}`;
      await executePowerShell(`
        ${psVar("_addName", addName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        $ws = $wb.Sheets.Add()
        $ws.Name = $_addName
      `);
      return `已添加工作表 ${addName}`;
    }
    case "copy": {
      const targetName = (options?.newName as string) || `${sheetName}_copy`;
      await executePowerShell(`
        ${varSheetName}
        ${psVar("_targetName", targetName)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        $source = $wb.Sheets.Item($_sheetName)
        $source.Copy($wb.Sheets.Item($wb.Sheets.Count))
        $wb.Sheets.Item($wb.Sheets.Count).Name = $_targetName
      `);
      return `已复制工作表为 ${targetName}`;
    }
    case "move": {
      const position = (options?.position as number) || 1;
      await executePowerShell(`
        ${varSheetName}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        $ws = $wb.Sheets.Item($_sheetName)
        $ws.Move($wb.Sheets.Item(${position}))
      `);
      return `已移动工作表 ${sheetName} 到位置 ${position}`;
    }
    default:
      throw new Error(`不支持的操作: ${operation}。支持 rename, delete, add, copy, move`);
  }
}
