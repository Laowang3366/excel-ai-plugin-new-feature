/**
 * ExcelScriptBridgeCom — 统一脚本执行桥接实现
 *
 * 根据环境自动选择最优引擎（所有宿主统一优先级）：
 * Python(xlwings) → JavaScript(cscript) → VBA
 */

import type { ExcelScriptBridge } from "../../contracts/excel";
import type { ScriptEnvironment, ScriptResult } from "../../contracts/scriptEnvironment";
import { createLogger } from "../../../../shared/logger";
import {
  executePythonScript,
  getPythonPath,
} from "../../../automation/python";
import { executeJScript } from "../../../automation/jscript";
import type { ExcelComBridge } from "./excelComBridge";
import { ExcelVbaComBridge } from "./excelVbaComBridge";

const logger = createLogger("ExcelScriptBridge");

/**
 * 统一脚本执行 Bridge
 *
 * 检测结果会缓存，避免重复检测浪费 token
 */
export class ExcelScriptBridgeCom implements ExcelScriptBridge {
  private comBridge: ExcelComBridge;
  private vbaBridge: ExcelVbaComBridge;
  /** 缓存的环境检测结果 */
  private _cachedEnv: ScriptEnvironment | null = null;

  constructor(comBridge: ExcelComBridge) {
    this.comBridge = comBridge;
    this.vbaBridge = new ExcelVbaComBridge(comBridge);
  }

  async detectEnvironment(): Promise<ScriptEnvironment> {
    // 返回缓存结果
    if (this._cachedEnv) return this._cachedEnv;

    const host = this.comBridge.host === "wps" ? "wps" : "excel";
    const available: ScriptEnvironment["available"] = [];

    // 统一优先级：Python → JS → VBA（所有宿主）
    await this.detectPython(available);
    await this.detectJs(available);
    await this.detectVba(available);

    const recommended = available.length > 0 ? available[0].language : "none";

    const env: ScriptEnvironment = { host, recommended, available };
    this._cachedEnv = env;
    return env;
  }

  async executeScript(code: string, language?: string): Promise<ScriptResult> {
    const env = await this.detectEnvironment();

    if (env.available.length === 0) {
      throw new Error("当前环境无可用的脚本引擎。请确保有 Python(xlwings)、cscript.exe 或 VBA 可用");
    }

    // 如果指定了语言，直接用对应引擎
    if (language) {
      const target = env.available.find((a) => a.language === language);
      if (!target) {
        throw new Error(`指定的语言 ${language} 在当前环境不可用。可用: ${env.available.map((a) => a.language).join(", ")}`);
      }
      return this.doExecute(code, target.language, target.engine);
    }

    // 未指定语言：按优先级尝试，失败自动 fallback
    const errors: string[] = [];
    for (const engine of env.available) {
      try {
        return await this.doExecute(code, engine.language, engine.engine);
      } catch (err: any) {
        errors.push(`${engine.language}(${engine.engine}): ${err.message}`);
        logger.warn(`脚本执行 ${engine.language} 失败，尝试下一个引擎: ${err.message}`);
      }
    }

    throw new Error(`所有可用引擎均执行失败:\n${errors.join("\n")}`);
  }

  /** 实际执行脚本 */
  private async doExecute(code: string, language: string, engine: string): Promise<ScriptResult> {
    switch (language) {
      case "javascript":
        return this.executeJs(code, engine);
      case "vba":
        return this.executeVba(code);
      case "python":
        return this.executePython(code);
      default:
        throw new Error(`不支持的语言: ${language}`);
    }
  }

  // ---- 环境检测 ----

  /** 检测 JS 引擎可用性（所有宿主统一检测，Excel 也支持 GetObject） */
  private async detectJs(available: ScriptEnvironment["available"]): Promise<void> {
    // 优先检测 cscript.exe（Windows 自带，不依赖 VBA）
    try {
      const progId = this.comBridge.host === "wps" ? "Ket.Application" : "Excel.Application";
      // 测试 cscript 是否可用 + GetObject 是否能连接 Excel
      const testScript = `
var excel;
try {
  excel = GetObject("", "${progId}");
  WScript.Echo("OK");
} catch(e) {
  WScript.Echo("NO_COM:" + e.message);
}
`;
      const result = await executeJScript(testScript, 5000);
      if (result.startsWith("OK")) {
        available.push({ language: "javascript", engine: "WindowsScriptHost" });
        return;
      }
    } catch { /* cscript 不可用 */ }

    // Fallback: 检测 MSScriptControl（需要 VBA 中转，仅 WPS）
    if (this.comBridge.host === "wps") {
      try {
        await this.vbaBridge.executeCode(`
Sub Main()
    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    Dim result As String
    result = sc.Eval("1+1")
End Sub
`);
        available.push({ language: "javascript", engine: "MSScriptControl" });
      } catch { /* MSScriptControl 不可用 */ }
    }
  }

  /** 检测 VBA 可用性 */
  private async detectVba(available: ScriptEnvironment["available"]): Promise<void> {
    try {
      const vbaAvailable = await this.vbaBridge.executeVbaCheck();
      if (vbaAvailable) {
        available.push({ language: "vba", engine: "VBA" });
      }
    } catch { /* VBA 不可用 */ }
  }

  /** 检测 Python 可用性（内置优先，系统次之） */
  private async detectPython(available: ScriptEnvironment["available"]): Promise<void> {
    try {
      const pythonPath = getPythonPath();
      const { execFileSync } = require("child_process");
      execFileSync(pythonPath, ["-c", "import xlwings"], {
        timeout: 5000, encoding: "utf8",
      });
      available.push({ language: "python", engine: "xlwings" });
    } catch { /* Python/xlwings 不可用 */ }
  }

  // ---- 脚本执行 ----

  /** 执行 JS 代码 */
  private async executeJs(code: string, engine: string): Promise<ScriptResult> {
    if (engine === "WindowsScriptHost") {
      return this.executeViaCscript(code);
    } else if (engine === "MSScriptControl") {
      return this.executeViaMsscript(code);
    }
    throw new Error(`未知的 JS 引擎: ${engine}`);
  }

  /** 通过 cscript.exe 执行 JS（所有宿主统一，不依赖 VBA） */
  private async executeViaCscript(code: string): Promise<ScriptResult> {
    const host = this.comBridge.host;
    const progId = host === "wps" ? "Ket.Application" : "Excel.Application";

    // 包装用户代码：自动获取 Excel COM 对象
    const wrappedCode =
      `// === 自动注入的 Excel 连接代码 ===\n` +
      `var excel = GetObject("", "${progId}");\n` +
      `var wb = excel.ActiveWorkbook;\n` +
      `var ws = excel.ActiveSheet;\n` +
      `// === 用户代码 ===\n` +
      code;

    const result = await executeJScript(wrappedCode, 90000);
    return { success: true, output: result, language: "javascript", engine: "WindowsScriptHost" };
  }

  /** 通过 MSScriptControl 执行 JS（VBA 中转，fallback） */
  private async executeViaMsscript(code: string): Promise<ScriptResult> {
    const jsB64 = Buffer.from(code, "utf16le").toString("base64");
    const vbaWrapper = `Sub Main()
    Dim sc As Object
    Set sc = CreateObject("MSScriptControl.ScriptControl")
    sc.Language = "JScript"
    sc.AddObject "excel", Application, True
    Dim jsCode As String
    jsCode = DecodeBase64("${jsB64}")
    sc.AddCode jsCode
    sc.Run "main"
End Sub

Function DecodeBase64(b64 As String) As String
    Dim objXML As Object
    Dim objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = b64
    Dim bytes() As Byte
    bytes = objNode.nodeTypedValue
    Dim i As Long
    Dim result As String
    result = ""
    For i = 0 To UBound(bytes) - 1 Step 2
        Dim ch As Long
        ch = CLng(bytes(i + 1)) * 256 + CLng(bytes(i))
        If ch > 0 Then
            result = result & ChrW(ch)
        End If
    Next i
    DecodeBase64 = result
End Function`;
    await this.vbaBridge.executeCode(vbaWrapper);
    return { success: true, language: "javascript", engine: "MSScriptControl" };
  }

  /** 执行 VBA 代码 */
  private async executeVba(code: string): Promise<ScriptResult> {
    await this.vbaBridge.executeCode(code);
    return { success: true, language: "vba", engine: "VBA" };
  }

  /** 执行 Python 代码（通过 executePythonScript，不再经 PowerShell 中转） */
  private async executePython(code: string): Promise<ScriptResult> {
    // executePythonScript 自动注入 xlwings 连接代码
    const result = await executePythonScript(code, 90000);
    return { success: true, output: result, language: "python", engine: "xlwings" };
  }
}
