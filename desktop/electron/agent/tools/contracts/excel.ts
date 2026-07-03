/**
 * Excel 工具契约
 *
 * 被工具执行器和 Excel/WPS 具体实现共同依赖，不包含任何 COM、PowerShell 或 UI 细节。
 */

import type { ScriptEnvironment, ScriptResult } from "./scriptEnvironment";

export type RangeReadExpandMode = "none" | "spill" | "currentArray" | "currentRegion";

export interface RangeReadResult {
  values: unknown[][];
  address?: string;
  expanded?: boolean;
  expandMode?: RangeReadExpandMode;
}

export interface ExcelConnectionStatus {
  connected: boolean;
  host: string;
  version?: string;
  workbookName?: string;
  availableHosts?: string[];
  error?: string;
}

/**
 * Excel 工作簿桥接接口
 *
 * 这是关键抽象点：所有工具执行都通过这个接口，
 * 实际实现可以是 COM 自动化、WPS 自动化接口等。
 */
export interface ExcelWorkbookBridge {
  /** 检查连接状态 */
  isConnected(): Promise<boolean>;
  /** 获取当前宿主信息 */
  getHostInfo(): Promise<{ host: "excel" | "wps"; version: string }>;

  /** 工作簿检查 */
  inspectWorkbook(): Promise<unknown>;
  /** 读取范围 */
  readRange(sheetName: string, range: string, expand?: RangeReadExpandMode): Promise<RangeReadResult>;
  /** 写入范围 */
  writeRange(sheetName: string, range: string, values: unknown[][]): Promise<void>;
  /** 清除范围 */
  clearRange(sheetName: string, range: string): Promise<void>;
  /** 获取选区 */
  getSelection(): Promise<{ address: string; values: unknown[][]; sheetName: string }>;
  /** 快速获取选区地址（不读取单元格值） */
  getSelectionAddress(): Promise<{ address: string; sheetName: string }>;
  /** 获取公式上下文 */
  getFormulaContext(sheetName: string, range?: string): Promise<unknown>;
  /** 工作表操作 */
  sheetOperation(operation: string, sheetName: string, options?: Record<string, unknown>): Promise<unknown>;

  // ---- 工作簿管理 ----
  /** 打开已有工作簿 */
  openWorkbook(filePath: string): Promise<{ success: boolean; workbookName?: string; error?: string }>;
  /** 创建新工作簿 */
  createWorkbook(filePath: string, sheetNames?: string[]): Promise<{ success: boolean; workbookName?: string; error?: string }>;
  /** 保存工作簿 */
  saveWorkbook(saveAsPath?: string): Promise<{ success: boolean; error?: string }>;
  /** 切换活动工作簿 */
  switchWorkbook(workbookName: string): Promise<{ success: boolean; workbookName?: string; error?: string }>;
}

/**
 * Excel 连接状态桥接接口。
 *
 * 供 IPC/运行时查询和切换 Excel/WPS 连接状态；具体 COM 实现仍留在 implementations 层。
 */
export interface ExcelConnectionBridge extends ExcelWorkbookBridge {
  /** 检测当前 Excel/WPS 连接状态 */
  detectStatus(): Promise<ExcelConnectionStatus>;
  /** 主动连接 Excel/WPS */
  connect(): Promise<ExcelConnectionStatus>;
  /** 多宿主同时存在时选择目标宿主 */
  selectHost(host: "excel" | "wps"): Promise<ExcelConnectionStatus>;
}

/**
 * VBA 桥接接口
 */
export interface ExcelVbaBridge {
  /** 检测 VBA 能力 */
  detectCapabilities(): Promise<{ supported: boolean; version?: string }>;
  /** 运行宏 */
  runMacro(macroName: string, args?: unknown[]): Promise<unknown>;
  /** 写入模块 */
  writeModule(moduleName: string, code: string): Promise<void>;
  /** 执行代码 */
  executeCode(code: string): Promise<unknown>;
}

/**
 * 统一脚本执行桥接接口
 *
 * 根据环境自动选择最优引擎：
 * - WPS: JS(cscript) → VBA → Python
 * - Office Excel: VBA → Python
 */
export interface ExcelScriptBridge {
  /** 检测脚本环境，返回可用语言及推荐 */
  detectEnvironment(): Promise<ScriptEnvironment>;
  /** 执行脚本代码，未指定语言时使用推荐语言，失败自动 fallback */
  executeScript(code: string, language?: string): Promise<ScriptResult>;
}

/**
 * UI 控件桥接接口
 */
export interface ExcelUiBridge {
  /** 添加工作表控件 */
  addControl(params: {
    sheetName: string;
    controlType: string;
    name: string;
    left: number;
    top: number;
    width: number;
    height: number;
    caption?: string;
    macroName?: string;
    linkedCell?: string;
  }): Promise<unknown>;
  /** 删除工作表控件 */
  removeControl(sheetName: string, name: string): Promise<void>;
  /** 列出工作表控件 */
  listControls(sheetName: string): Promise<unknown[]>;
  /** 创建 UserForm 窗体 */
  createForm(params: {
    formName: string;
    caption: string;
    width?: number;
    height?: number;
    controls?: Array<Record<string, unknown>>;
    eventCode?: string;
  }): Promise<unknown>;
  /** 添加自定义菜单项 */
  addMenu(params: {
    menuBar: string;
    caption: string;
    macroName: string;
    beforeId?: number;
    faceId?: number;
  }): Promise<unknown>;
}
