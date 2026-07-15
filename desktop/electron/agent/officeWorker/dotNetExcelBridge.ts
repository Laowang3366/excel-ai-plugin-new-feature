import type {
  ExcelConnectionBridge,
  ExcelConnectionStatus,
  RangeReadExpandMode,
  RangeReadResult,
  RangeWriteResult,
} from "../tools/contracts/excel";
import { getOfficeWorkerClient, type OfficeWorkerClient } from "./officeWorkerClient";

export class DotNetExcelBridge implements ExcelConnectionBridge {
  private status: ExcelConnectionStatus = { connected: false, host: "unknown" };

  constructor(private readonly client: OfficeWorkerClient = getOfficeWorkerClient()) {}

  get host(): "excel" | "wps" | "unknown" {
    return this.status.host === "excel" || this.status.host === "wps"
      ? this.status.host
      : "unknown";
  }

  async isConnected(): Promise<boolean> {
    return this.status.connected;
  }

  async getHostInfo(): Promise<{ host: "excel" | "wps"; version: string }> {
    return {
      host: this.host === "wps" ? "wps" : "excel",
      version: this.status.version || "unknown",
    };
  }

  async detectStatus(): Promise<ExcelConnectionStatus> {
    this.status = await this.client.invoke<ExcelConnectionStatus>("excel.detectStatus");
    return this.status;
  }

  async connect(): Promise<ExcelConnectionStatus> {
    this.status = await this.client.invoke<ExcelConnectionStatus>("excel.connect");
    return this.status;
  }

  async selectHost(host: "excel" | "wps"): Promise<ExcelConnectionStatus> {
    this.status = await this.client.invoke<ExcelConnectionStatus>("excel.selectHost", { host });
    return this.status;
  }

  async disconnect(): Promise<void> {
    this.status = { connected: false, host: "unknown" };
  }

  inspectWorkbook(): Promise<unknown> {
    return this.client.invoke("excel.workbook.inspect");
  }

  readRange(
    sheetName: string,
    range: string,
    expand: RangeReadExpandMode = "none",
  ): Promise<RangeReadResult> {
    return this.client.invoke("excel.range.read", { sheetName, range, expand });
  }

  writeRange(
    sheetName: string,
    range: string,
    values: unknown[][],
    options?: { legacyCse?: boolean },
  ): Promise<RangeWriteResult> {
    return this.client.invoke("excel.range.write", {
      sheetName,
      range,
      values,
      legacyCse: options?.legacyCse === true,
    });
  }

  async clearRange(sheetName: string, range: string): Promise<void> {
    await this.client.invoke("excel.range.clear", { sheetName, range });
  }

  getSelection(): Promise<{ address: string; values: unknown[][]; sheetName: string }> {
    return this.client.invoke("excel.selection.read");
  }

  getSelectionAddress(): Promise<{ address: string; sheetName: string }> {
    return this.client.invoke("excel.selection.address");
  }

  getFormulaContext(sheetName: string, range?: string): Promise<unknown> {
    return this.client.invoke("excel.formula.context", { sheetName, range });
  }

  sheetOperation(
    operation: string,
    sheetName: string,
    options?: Record<string, unknown>,
  ): Promise<unknown> {
    return this.client.invoke("excel.sheet.operation", {
      operation,
      sheetName,
      options: options || {},
    });
  }

  openWorkbook(
    filePath: string,
  ): Promise<{ success: boolean; workbookName?: string; error?: string }> {
    return this.client.invoke("excel.workbook.open", { filePath });
  }

  createWorkbook(
    filePath: string,
    sheetNames?: string[],
  ): Promise<{ success: boolean; workbookName?: string; error?: string }> {
    return this.client.invoke("excel.workbook.create", { filePath, sheetNames: sheetNames || [] });
  }

  saveWorkbook(saveAsPath?: string): Promise<{ success: boolean; error?: string }> {
    return this.client.invoke("excel.workbook.save", { saveAsPath });
  }

  switchWorkbook(
    workbookName: string,
  ): Promise<{ success: boolean; workbookName?: string; error?: string }> {
    return this.client.invoke("excel.workbook.switch", { workbookName });
  }
}
