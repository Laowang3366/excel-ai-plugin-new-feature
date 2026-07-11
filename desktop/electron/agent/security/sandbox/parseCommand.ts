/**
 * 命令解析 — 把命令字符串切成 token 数组
 *
 * 参考 Codex execpolicy 的命令语义分析前置环节：策略层只对 token 数组生效，
 * 必须先把"明文一整条"拆解成可识别的命令 + 参数序列，才能做前缀匹配。
 *
 * 设计取舍：
 * - 不引入完整 shlex 依赖；本插件主要在 Windows 上执行 PowerShell，
 *   而完整 PowerShell tokenize 需要嵌入解析器，成本过高、收益边际。
 * - 采用保守的纯 TS 扫词：
 *   * 单/双引号包围的字符串整体作为一个 token
 *   * 反引号 ` 转义下一个字符（PowerShell）；非 Windows 下反斜杠
 *   * 管道 `|`、分号 `;`、`&&` / `&`、换行作为命令边界
 *   * 注释 `#...` 至行末被剥离
 *   * 字符串外的连续空白为分隔
 *
 * 这与 Codex prefix_rule 的匹配方式契合：取每个子命令的前缀 token 做决策，
 * 破坏性子命令不论它出现在管道前/中/后都会被命中。
 */

/** 一段切分出的子命令（一个管道段或一行） */
export interface ParsedCommand {
  /** 子命令的原始文本（不含边界符；用于审计回放） */
  raw: string;
  /** token 数组：首元素是命令名，后续为参数 */
  tokens: string[];
  /** 该段是否以引号/转义失败等不可解析标志开头 */
  parseFailed?: boolean;
}

/**
 * 把一整段命令文本切为子命令 + token 数组
 *
 * @param command 多行/多段命令
 * @returns 子命令数组；空命令返回 []
 */
export function parseCommand(command: string): ParsedCommand[] {
  if (typeof command !== "string" || command.length === 0) return [];

  const results: ParsedCommand[] = [];
  let token = "";
  let inSingle = false;
  let inDouble = false;
  let escape = false;
  let failed = false;
  let currentTokens: string[] = [];
  let currentRaw = "";

  const closeToken = () => {
    if (token.length > 0) {
      currentTokens.push(token);
    }
    token = "";
  };

  const closeCommand = () => {
    closeToken();
    if (currentTokens.length > 0) {
      results.push({ raw: currentRaw.trim(), tokens: currentTokens, parseFailed: failed });
    }
    currentTokens = [];
    currentRaw = "";
    failed = false;
  };

  for (let i = 0; i < command.length; i++) {
    const c = command[i];

    // 注释（仅当不在引号内）—— 剥到行末
    if (!inSingle && !inDouble && !escape && c === "#") {
      while (i < command.length && command[i] !== "\n" && command[i] !== "\r") {
        i++;
      }
      continue;
    }

    currentRaw += c;

    if (escape) {
      token += c;
      escape = false;
      continue;
    }

    if (c === "`" && !inSingle) {
      // PowerShell 反引号转义；下一字符原样入 token
      escape = true;
      continue;
    }
    if (c === "\\" && !inSingle && !inDouble && process.platform !== "win32") {
      // Unix bash 反斜杠转义（仅非字符串内）
      escape = true;
      continue;
    }

    if (c === "'" && !inDouble) {
      inSingle = !inSingle;
      continue;
    }
    if (c === '"' && !inSingle) {
      inDouble = !inDouble;
      continue;
    }

    // 命令边界
    if (!inSingle && !inDouble) {
      if (c === "|" || c === ";" || c === "\n" || c === "\r") {
        closeCommand();
        continue;
      }
      if (c === "&") {
        if (command[i + 1] === "&") {
          i++; // 吃掉第二个 &
        }
        closeCommand();
        continue;
      }
      if (/\s/.test(c)) {
        closeToken();
        continue;
      }
    }

    token += c;
  }

  // 收尾
  if (inSingle || inDouble || escape) {
    failed = true;
  }
  closeCommand();

  return results;
}
