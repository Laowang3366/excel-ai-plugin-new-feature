/**
 * ExcelComBridge — 工作簿 COM 桥接实现
 *
 * 通过 JScript/PowerShell COM 自动化与 Excel/WPS 交互，
 * 实现 ExcelWorkbookBridge 接口。
 */

import type { ExcelWorkbookBridge, RangeReadExpandMode, RangeReadResult } from "../../contracts/excel";
import {
  normalizeWorkbookInspectMetadata,
  resolveSpreadsheetHost,
  type SpreadsheetHost,
} from "./connectionMetadata";
import {
  detectExcelProcess,
  verifyExcelComAvailable,
} from "./excelConnectionProbe";
import {
  createWorkbookOperation,
  inspectWorkbookOperation,
  openWorkbookOperation,
  saveWorkbookOperation,
  switchWorkbookOperation,
} from "./workbookOperations";
import {
  clearRangeOperation,
  getSelectionAddressOperation,
  getSelectionOperation,
  readRangeOperation,
  writeRangeOperation,
} from "./rangeOperations";
import { getFormulaContextOperation } from "./formulaOperations";
import { sheetOperation as runSheetOperation } from "./sheetOperations";

export { normalizeWorkbookInspectMetadata, resolveSpreadsheetHost };

type ExcelConnectionStatus = {
  connected: boolean;
  host: SpreadsheetHost | "unknown";
  version?: string;
  workbookName?: string;
  availableHosts?: SpreadsheetHost[];
};

export class ExcelComBridge implements ExcelWorkbookBridge {
  private _connected = false;
  private _host: SpreadsheetHost | "unknown" = "unknown";
  /** 当 Office + WPS 同时存在时，用户手动选择的宿主 */
  private _selectedHost: SpreadsheetHost | null = null;
  private _comVersion?: string;

  async isConnected(): Promise<boolean> {
    return this._connected;
  }

  async getHostInfo(): Promise<{ host: "excel" | "wps"; version: string }> {
    return {
      host: (this._host === "unknown" ? "excel" : this._host) as "excel" | "wps",
      version: this._comVersion || "unknown",
    };
  }

  /** 获取当前 host（供 VBA Bridge 使用） */
  get host(): SpreadsheetHost | "unknown" {
    return this._host;
  }

  /**
   * 检测 Excel 连接状态（进程 + COM 可用性）
   *
   * 仅返回检测结果供前端展示，同时同步更新内部连接状态，
   * 确保前端显示与工具执行的状态一致。
   *
   * 当 Office + WPS 同时运行时返回 availableHosts，
   * 由 UI 弹窗让用户选择后再调用 selectHost 连接。
   */
  async detectStatus(): Promise<{
    connected: boolean;
    host: string;
    version?: string;
    workbookName?: string;
    availableHosts?: string[];
  }> {
    try {
      const proc = await detectExcelProcess();
      if (!proc.running) {
        // 进程不存在，同步断开内部状态
        this._connected = false;
        this._host = "unknown";
        this._selectedHost = null;
        return { connected: false, host: "unknown" };
      }

      // 两个宿主同时运行 → 优先检查 _selectedHost，
      // 如果尚未选择则返回 availableHosts 让前端弹窗
      if (proc.availableHosts.length > 1) {
        this._host = "unknown";
        this._connected = false;
        // 如果之前已经选择过，用选择的宿主尝试连接
        if (this._selectedHost) {
          const comResult = await verifyExcelComAvailable(this._selectedHost);
          if (comResult.available) {
            this._connected = true;
            this._host = this._selectedHost;
            this._comVersion = comResult.version;
            return {
              connected: true,
              host: this._selectedHost,
              version: comResult.version,
              workbookName: comResult.workbookName,
              availableHosts: proc.availableHosts,
            };
          }
          // 选中的宿主 COM 不可用，重新选择
          this._selectedHost = null;
        }
        return {
          connected: false,
          host: "unknown",
          availableHosts: proc.availableHosts,
        };
      }

      const singleHost = proc.availableHosts[0];
      if (!singleHost) {
        this._connected = false;
        this._host = "unknown";
        return { connected: false, host: "unknown" };
      }

      // 仅有一个宿主运行 → 正常流程
      const comResult = await verifyExcelComAvailable(singleHost);
      if (comResult.available) {
        this._connected = true;
        this._host = singleHost;
        this._selectedHost = singleHost;
        this._comVersion = comResult.version;
        return {
          connected: true,
          host: singleHost,
          version: comResult.version,
          workbookName: comResult.workbookName,
        };
      }

      // 进程存在但 COM 不可用
      this._connected = false;
      return { connected: false, host: singleHost };
    } catch {
      this._connected = false;
      return { connected: false, host: "unknown" };
    }
  }

  /**
   * 连接到 Excel/WPS 实例
   *
   * 当 Office + WPS 同时运行时，需要先调用 selectHost() 选择目标宿主，
   * 否则 connect() 返回 false。
   */
  async connect(): Promise<ExcelConnectionStatus> {
    return this.detectAndConnect(0, false);
  }

  /** 断开连接 */
  async disconnect(): Promise<void> {
    this._connected = false;
    this._selectedHost = null;
  }

  /**
   * 用户手动选择目标宿主（当 Office + WPS 同时运行时）
   *
   * 保存选择并尝试 COM 连接，返回连接结果。
   */
  async selectHost(host: "excel" | "wps"): Promise<{
    connected: boolean;
    host: string;
    version?: string;
    workbookName?: string;
  }> {
    this._selectedHost = host;
    const comResult = await verifyExcelComAvailable(host);
    if (comResult.available) {
      this._connected = true;
      this._host = host;
      this._comVersion = comResult.version;
      return {
        connected: true,
        host: host,
        version: comResult.version,
        workbookName: comResult.workbookName,
      };
    }
    this._connected = false;
    this._selectedHost = null;
    return { connected: false, host: host };
  }

  // ----------------------------------------------------------
  // ExcelWorkbookBridge 接口实现
  //
  // 优先通过 Python(xlwings) → JScript(cscript) → PowerShell 执行
  // Python/JScript 语法模型更熟悉，序列化可靠，出错率更低
  // ----------------------------------------------------------

  async inspectWorkbook(): Promise<unknown> {
    return inspectWorkbookOperation(this.getWorkbookOperationDeps());
  }

  async readRange(sheetName: string, range: string, expand?: RangeReadExpandMode): Promise<RangeReadResult> {
    return readRangeOperation(this.getRangeOperationDeps(), sheetName, range, expand);
  }

  async writeRange(sheetName: string, range: string, values: unknown[][]): Promise<void> {
    await writeRangeOperation(this.getRangeOperationDeps(), sheetName, range, values);
  }

  async clearRange(sheetName: string, range: string): Promise<void> {
    await clearRangeOperation(this.getRangeOperationDeps(), sheetName, range);
  }

  async getSelection(): Promise<{ address: string; values: unknown[][]; sheetName: string }> {
    return getSelectionOperation(
      this.getRangeOperationDeps(),
      async (sheetName, range) => (await this.readRange(sheetName, range)).values
    );
  }

  async getSelectionAddress(): Promise<{ address: string; sheetName: string }> {
    return getSelectionAddressOperation(this.getRangeOperationDeps());
  }

  async getFormulaContext(sheetName: string, range?: string): Promise<unknown> {
    return getFormulaContextOperation(this.getFormulaOperationDeps(), sheetName, range);
  }

  async sheetOperation(
    operation: string,
    sheetName: string,
    options?: Record<string, unknown>
  ): Promise<unknown> {
    return runSheetOperation(this.getSheetOperationDeps(), operation, sheetName, options);
  }

  // ----------------------------------------------------------
  // 工作簿管理
  // ----------------------------------------------------------

  /**
   * 打开已有工作簿
   *
   * 通过 COM Workbooks.Open 打开指定路径的文件，
   * 打开后自动成为活动工作簿。
   */
  async openWorkbook(filePath: string): Promise<{ success: boolean; workbookName?: string; error?: string }> {
    return openWorkbookOperation(this.getWorkbookOperationDeps(), filePath);
  }

  /**
   * 创建新工作簿
   *
   * 通过 COM Workbooks.Add 创建空白工作簿，
   * 可选指定初始工作表名称，然后 SaveAs 到指定路径。
   */
  async createWorkbook(filePath: string, sheetNames?: string[]): Promise<{ success: boolean; workbookName?: string; error?: string }> {
    return createWorkbookOperation(this.getWorkbookOperationDeps(), filePath, sheetNames);
  }

  /**
   * 保存工作簿
   *
   * 保存当前活动工作簿。如果指定 saveAsPath 则另存为新文件。
   */
  async saveWorkbook(saveAsPath?: string): Promise<{ success: boolean; error?: string }> {
    return saveWorkbookOperation(this.getWorkbookOperationDeps(), saveAsPath);
  }

  /**
   * 切换活动工作簿
   *
   * 通过名称切换到已打开的工作簿，使其成为活动工作簿。
   */
  async switchWorkbook(workbookName: string): Promise<{ success: boolean; workbookName?: string; error?: string }> {
    return switchWorkbookOperation(this.getWorkbookOperationDeps(), workbookName);
  }

  // ----------------------------------------------------------
  // 私有辅助方法
  // ----------------------------------------------------------

  private async getActiveHostOrDetect(retries = 1): Promise<SpreadsheetHost | null> {
    if (retries === 0 && this._connected && (this._selectedHost || this._host !== "unknown")) {
      return (this._selectedHost || this._host) as SpreadsheetHost;
    }
    return this.getDetectedHost(retries);
  }

  private getWorkbookOperationDeps() {
    return {
      ensureConnected: (retries?: number) => this.getActiveHostOrDetect(retries),
      getProgId: () => this.getProgId(),
      getComVersion: () => this._comVersion,
    };
  }

  private getRangeOperationDeps() {
    return {
      ensureConnected: (retries?: number) => this.getActiveHostOrDetect(retries),
      getProgId: () => this.getProgId(),
    };
  }

  private getFormulaOperationDeps() {
    return {
      ensureConnected: (retries?: number) => this.getActiveHostOrDetect(retries),
      getProgId: () => this.getProgId(),
    };
  }

  private getSheetOperationDeps() {
    return {
      ensureConnected: () => this.getDetectedHost(),
      getProgId: () => this.getProgId(),
    };
  }

  /** 获取 COM ProgID（Excel vs WPS） */
  private getProgId(): string {
    const target = this._selectedHost || this._host;
    return target === "wps" ? "Ket.Application" : "Excel.Application";
  }

  private async getDetectedHost(retries = 1): Promise<SpreadsheetHost | null> {
    const status = await this.detectAndConnect(retries);
    return status.connected && status.host !== "unknown" ? status.host : null;
  }

  /**
   * 检测并尝试连接，更新内部状态。
   *
   * 工具执行默认允许沿用当前宿主，前端手动 connect 则要求多宿主场景先 selectHost。
   */
  private async detectAndConnect(
    retries = 1,
    allowCurrentHostFallback = true
  ): Promise<ExcelConnectionStatus> {
    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        const proc = await detectExcelProcess();
        const currentHost = allowCurrentHostFallback ? this._host : "unknown";
        const targetHost = resolveSpreadsheetHost(proc.availableHosts, this._selectedHost, currentHost);
        if (!proc.running || !targetHost) {
          this._connected = false;
          this._host = "unknown";
          if (attempt < retries) {
            await new Promise(r => setTimeout(r, 1000));
            continue;
          }
          return {
            connected: false,
            host: "unknown",
            availableHosts: proc.availableHosts.length > 1 ? proc.availableHosts : undefined,
          };
        }

        // 更新 host
        this._host = targetHost;
        if (proc.availableHosts.length === 1) {
          this._selectedHost = targetHost;
        }

        // 验证 COM 可用性
        const comResult = await verifyExcelComAvailable(targetHost);
        if (comResult.available) {
          this._connected = true;
          this._comVersion = comResult.version;
          return {
            connected: true,
            host: targetHost,
            version: comResult.version,
            workbookName: comResult.workbookName,
          };
        }

        this._connected = false;
        if (attempt < retries) {
          await new Promise(r => setTimeout(r, 1000));
          continue;
        }
        return { connected: false, host: targetHost };
      } catch {
        this._connected = false;
        if (attempt < retries) {
          await new Promise(r => setTimeout(r, 1000));
          continue;
        }
        return { connected: false, host: "unknown" };
      }
    }
    return { connected: false, host: "unknown" };
  }
}
