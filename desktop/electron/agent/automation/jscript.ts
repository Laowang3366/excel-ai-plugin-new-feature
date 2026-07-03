import { DEFAULT_PROCESS_MAX_BUFFER } from "./processLimits";
import { decodeProcessOutput } from "./stdioEncoding";

/**
 * JScript 自动化基础能力
 *
 * 负责通过 Windows cscript.exe 执行 JScript，并提供安全字符串变量注入。
 */

/**
 * 执行 JScript 脚本（通过 cscript.exe），返回 stdout。
 */
export function executeJScript(script: string, timeout = 90000): Promise<string> {
  const { execFile } = require("child_process");
  const fs = require("fs");
  const path = require("path");
  const os = require("os");

  return new Promise((resolve, reject) => {
    const tmpFile = path.join(os.tmpdir(), `excel_${Date.now()}_${Math.random().toString(36).slice(2)}.js`);
    try {
      const bom = Buffer.from([0xff, 0xfe]);
      const contentBuf = Buffer.from(script, "utf16le");
      fs.writeFileSync(tmpFile, Buffer.concat([bom, contentBuf]));
    } catch (writeErr: any) {
      reject(new Error(`写入临时 JS 文件失败: ${writeErr.message}`));
      return;
    }

    execFile(
      "cscript.exe",
      ["//NoLogo", "//E:JScript", tmpFile],
      { timeout, maxBuffer: DEFAULT_PROCESS_MAX_BUFFER, encoding: "buffer", windowsHide: true },
      (err: any, stdout: Buffer, stderr: Buffer) => {
        try { fs.unlinkSync(tmpFile); } catch { /* 忽略清理失败 */ }

        const stdoutText = decodeProcessOutput(stdout).trim();
        const stderrText = decodeProcessOutput(stderr).trim();
        if (err) {
          reject(new Error(stderrText || err.message));
        } else {
          resolve(stdoutText);
        }
      }
    );
  });
}

/**
 * 将字符串值以 JScript 变量赋值的方式注入脚本。
 */
export function jsVar(name: string, value: string): string {
  const b64 = Buffer.from(value, "utf16le").toString("base64");
  return `var ${name} = (function() {
  var objNode = new ActiveXObject("MSXML2.DOMDocument").createElement("b64");
  objNode.DataType = "bin.base64";
  objNode.Text = "${b64}";
  var bytes = objNode.NodeTypedValue;
  var stream = new ActiveXObject("ADODB.Stream");
  stream.Type = 1;
  stream.Open();
  stream.Write(bytes);
  stream.Position = 0;
  stream.Type = 2;
  stream.Charset = "Unicode";
  var result = stream.ReadText();
  stream.Close();
  return result;
})();`;
}
