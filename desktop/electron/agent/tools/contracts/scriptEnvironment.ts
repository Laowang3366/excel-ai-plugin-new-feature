/**
 * 脚本运行契约
 *
 * 被 Excel 脚本桥和 Office 工具执行器共享，描述脚本环境检测结果和执行结果。
 */

/**
 * 脚本环境检测结果
 */
export interface ScriptEnvironment {
  /** 当前宿主类型 */
  host: "excel" | "wps";
  /** 推荐的脚本语言 */
  recommended: string;  // "vba" | "javascript" | "python"
  /** 可用语言列表，按优先级排列 */
  available: Array<{
    language: string;   // "vba" | "javascript" | "python"
    engine: string;     // "VBA" | "WindowsScriptHost" | "MSScriptControl" | "xlwings"
  }>;
}

/**
 * 脚本执行结果
 */
export interface ScriptResult {
  success: boolean;
  output?: string;
  /** 实际使用的语言 */
  language: string;
  /** 实际使用的引擎 */
  engine: string;
}
