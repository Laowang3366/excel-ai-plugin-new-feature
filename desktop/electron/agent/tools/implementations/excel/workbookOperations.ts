/**
 * Excel/WPS 工作簿能力。
 *
 * 关联模块：
 * - excelComBridge.ts: 对外保留 ExcelWorkbookBridge 门面，委托本模块执行工作簿操作。
 * - automation: 通过 Python 优先、PowerShell 兜底执行 COM 自动化。
 */

import {
  executePowerShell,
  psVar,
} from "../../../automation/powershell";
import { executeSmart } from "../../../automation/scriptEngine";
import { safeJsonParse } from "../../../automation/json";
import { normalizeWorkbookInspectMetadata, type SpreadsheetHost } from "./connectionMetadata";

export interface WorkbookOperationDeps {
  ensureConnected: (retries?: number) => Promise<SpreadsheetHost | null>;
  getProgId: () => string;
  getComVersion: () => string | undefined;
}

export async function inspectWorkbookOperation(deps: WorkbookOperationDeps): Promise<unknown> {
  const host = await deps.ensureConnected(0);
  if (!host) {
    throw new Error("未连接到 Excel/WPS，请先在侧边栏点击连接");
  }

  try {
    const progId = deps.getProgId();
    const pythonScript = `
info = {
    'name': app.name,
    'version': app.version,
    'workbooks': []
}
for b in app.books:
    wb_info = {
        'name': b.name,
        'path': b.fullpath if hasattr(b, 'fullpath') else '',
        'sheets': []
    }
    for s in b.sheets:
        used = s.used_range
        wb_info['sheets'].append({
            'name': s.name,
            'rows': used.rows.count if used else 0,
            'columns': used.columns.count if used else 0,
        })
    info['workbooks'].append(wb_info)
print(json.dumps(info, ensure_ascii=False))
`;

    const psScript = `
$excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
$info = @{
  name = $excel.Name
  version = $excel.Version
  workbooks = @()
}
foreach ($wb in $excel.Workbooks) {
  $sheets = @()
  foreach ($ws in $wb.Worksheets) {
    $sheets += @{
      name = $ws.Name
      rows = $ws.UsedRange.Rows.Count
      columns = $ws.UsedRange.Columns.Count
    }
  }
  $info.workbooks += @{
    name = $wb.Name
    path = $wb.Path
    sheets = $sheets
  }
}
$info | ConvertTo-Json -Depth 5
`;

    const { result, engine } = await executeSmart(pythonScript, psScript, 90000);
    const parsed = safeJsonParse(result, engine, "检查工作簿");
    return normalizeWorkbookInspectMetadata(parsed, host, deps.getComVersion());
  } catch (err: any) {
    throw new Error(`检查工作簿失败: ${err.message}`);
  }
}

export async function openWorkbookOperation(
  deps: WorkbookOperationDeps,
  filePath: string
): Promise<{ success: boolean; workbookName?: string; error?: string }> {
  const host = await deps.ensureConnected();
  if (!host) return { success: false, error: "未连接到 Excel/WPS，请先在侧边栏点击连接" };

  const progId = deps.getProgId();
  try {
    const result = await executePowerShell(`
      ${psVar("_filePath", filePath)}
      $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
      $wb = $excel.Workbooks.Open($_filePath)
      $wb.Activate()
      $excel.Visible = $true
      $wb.Name
    `);
    const workbookName = result.trim();
    return { success: true, workbookName };
  } catch (err: any) {
    return { success: false, error: `打开工作簿失败: ${err.message}` };
  }
}

export async function createWorkbookOperation(
  deps: WorkbookOperationDeps,
  filePath: string,
  sheetNames?: string[]
): Promise<{ success: boolean; workbookName?: string; error?: string }> {
  const host = await deps.ensureConnected();
  if (!host) return { success: false, error: "未连接到 Excel/WPS，请先在侧边栏点击连接" };

  const progId = deps.getProgId();
  try {
    let script = `
      ${psVar("_filePath", filePath)}
      $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
      $wb = $excel.Workbooks.Add()
      $wb.Activate()`;

    if (sheetNames && sheetNames.length > 0) {
      script += `
      ${psVar("_sheet0", sheetNames[0])}
      $wb.Sheets.Item(1).Name = $_sheet0`;

      for (let i = 1; i < sheetNames.length; i++) {
        script += `
        ${psVar(`_sheet${i}`, sheetNames[i])}
        $ws = $wb.Sheets.Add()
        $ws.Name = $_sheet${i}`;
      }

      if (sheetNames.length < 3) {
        script += `
        while ($wb.Sheets.Count -gt ${sheetNames.length}) {
          $wb.Sheets.Item($wb.Sheets.Count).Delete()
        }`;
      }
    }

    script += `
      $wb.SaveAs($_filePath)
      $excel.Visible = $true
      $wb.Name`;

    const result = await executePowerShell(script);
    const workbookName = result.trim();
    return { success: true, workbookName };
  } catch (err: any) {
    return { success: false, error: `创建工作簿失败: ${err.message}` };
  }
}

export async function saveWorkbookOperation(
  deps: WorkbookOperationDeps,
  saveAsPath?: string
): Promise<{ success: boolean; error?: string }> {
  const host = await deps.ensureConnected();
  if (!host) return { success: false, error: "未连接到 Excel/WPS，请先在侧边栏点击连接" };

  const progId = deps.getProgId();
  try {
    if (saveAsPath) {
      await executePowerShell(`
        ${psVar("_saveAsPath", saveAsPath)}
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        $wb.SaveAs($_saveAsPath)
      `);
    } else {
      await executePowerShell(`
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
        $wb = $excel.ActiveWorkbook
        $wb.Save()
      `);
    }
    return { success: true };
  } catch (err: any) {
    return { success: false, error: `保存工作簿失败: ${err.message}` };
  }
}

export async function switchWorkbookOperation(
  deps: WorkbookOperationDeps,
  workbookName: string
): Promise<{ success: boolean; workbookName?: string; error?: string }> {
  const host = await deps.ensureConnected();
  if (!host) return { success: false, error: "未连接到 Excel/WPS，请先在侧边栏点击连接" };

  const progId = deps.getProgId();
  try {
    const result = await executePowerShell(`
      ${psVar("_wbName", workbookName)}
      $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject('${progId}')
      $wb = $excel.Workbooks.Item($_wbName)
      $wb.Activate()
      $wb.Name
    `);
    const activeName = result.trim();
    return { success: true, workbookName: activeName };
  } catch (err: any) {
    return { success: false, error: `切换工作簿失败（可能未打开）: ${err.message}` };
  }
}
