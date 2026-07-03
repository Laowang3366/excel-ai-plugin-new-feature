/**
 * 默认命令策略规则
 *
 * 参考 Codex execpolicy 的 prefix_rule 风格：每条规则按前缀 token 匹配一目
 * 的子命令，命中后产出 decision。决策语义：
 *   - "forbidden"：拒绝执行，记审计并告诉模型理由（justifycation）
 *   - "prompt"  ：进入用户审批（即便 permissionMode 是 confirm_all）
 *   - "allow"   ：保持原审批/沙箱路径
 *
 * 多规则同时命中取最严：forbidden > prompt > allow。
 *
 * 规则按"首 token"建立索引（first 字段），其余 token 在 rest 中，每个
 * entry 可为字符串或可选集合（表示互斥的多个等价 token）。
 *
 * 注：首 token 必须唯一字符串，与 Codex rule.rs 中 PrefixPattern.first 一致。
 */

import type { PatternToken, PrefixRule } from "./execPolicy";

/** 便捷构造：把可读的 ["cmd", ["altA","altB"], ...] 形式转为规则对象 */
function r(
  pattern: Array<string | string[]>,
  decision: PrefixRule["decision"],
  justification?: string
): PrefixRule {
  if (pattern.length === 0) {
    throw new Error("rule pattern must have at least one token");
  }
  const first = pattern[0];
  if (typeof first !== "string") {
    throw new Error("first pattern token must be a single string");
  }
  const rest: PatternToken[] = [];
  for (let i = 1; i < pattern.length; i++) {
    const tok = pattern[i];
    if (Array.isArray(tok)) {
      if (tok.length === 0) throw new Error("alts token must not be empty");
      rest.push(tok.length === 1 ? { kind: "single", value: tok[0] } : { kind: "alts", values: tok });
    } else {
      rest.push({ kind: "single", value: tok });
    }
  }
  return { first, rest, decision, justification };
}

/**
 * 出厂默认规则 — 严禁破坏性命令；后台长期运行进入 prompt；
 * 其它命令交给 permissionMode 决定是否审批。
 *
 * 任何模式（包括 confirm_all + alwaysAllow）下，forbidden 永远拒绝。
 * 仅保留少量高风险持久化类 prompt，继续强制走审批对话框。
 */
export const DEFAULT_RULES: PrefixRule[] = [
  // === 破坏性系统命令 ===
  r(["rm", "-rf", "/"], "forbidden", "防止清盘：rm -rf /"),
  r(["rm", "-rf", "~"], "forbidden", "防止清空用户主目录"),
  r(["Remove-Item", ["-Recurse", "-r", "-R"], ["-Force", "-f"]], "forbidden", "递归强制删除"),
  r(["Remove-Item", "-Force"], "forbidden", "强制删除"),
  r(["del", "/s", "/q"], "forbidden", "批量静默删除"),
  r(["rd", "/s", "/q"], "forbidden", "递归强制删目录"),
  r(["rmdir", "/s", "/q"], "forbidden", "递归强制删目录"),
  r(["Format"], "forbidden", "格式化磁盘"),
  r(["format"], "forbidden", "格式化磁盘"),
  r(["diskpart"], "forbidden", "磁盘分区工具可清盘，禁止直接执行"),
  r(["Stop-Computer"], "forbidden", "关机"),
  r(["Restart-Computer"], "forbidden", "重启"),
  r(["shutdown"], "forbidden", "关机/重启"),
  r(["reg", "delete"], "forbidden", "删除注册表项"),
  r(["sc", "delete"], "forbidden", "删除服务"),
  r(["taskkill", "/f"], "forbidden", "强杀进程"),
  r(["mkfs"], "forbidden", "格式化文件系统"),

  // === 远程脚本执行 / 下载即执行 ===
  r(["iex"], "forbidden", "Invoke-Expression 可执行任意字符串，禁止直接调用"),
  r(["Invoke-Expression"], "forbidden", "可执行任意字符串，禁止直接调用"),
  r(["Invoke-WebRequest"], "allow", "网络下载：跟随用户权限模式"),
  r(["Invoke-RestMethod"], "allow", "网络请求：跟随用户权限模式"),
  r(["iwr"], "allow", "Invoke-WebRequest 别名，跟随用户权限模式"),
  r(["irm"], "allow", "Invoke-RestMethod 别名，跟随用户权限模式"),
  r(["curl"], "allow", "网络下载，跟随用户权限模式"),
  r(["wget"], "allow", "网络下载，跟随用户权限模式"),
  r(["Start-BitsTransfer"], "allow", "后台下载，跟随用户权限模式"),
  r(["powershell"], "allow", "嵌套 PowerShell，跟随用户权限模式"),
  r(["pwsh"], "allow", "嵌套 PowerShell Core，跟随用户权限模式"),
  r(["cmd"], "allow", "嵌套 cmd，跟随用户权限模式"),
  r(["cmd.exe"], "allow", "嵌套 cmd，跟随用户权限模式"),
  r(["bash"], "allow", "嵌套 bash，跟随用户权限模式"),
  r(["nohup"], "prompt", "后台长期运行，请确认"),
  r(["crontab"], "forbidden", "修改定时任务可造成持久化，禁止直接执行"),

  // === 提权 ===
  r(["sudo"], "allow", "提权执行，跟随用户权限模式"),
  r(["runas"], "allow", "提权执行，跟随用户权限模式"),
  r(["Start-Process", "-Verb", "RunAs"], "allow", "提权启动进程，跟随用户权限模式"),

  // === 持久化 / 凭据 ===
  r(["net", "user"], "forbidden", "账户管理，禁止直接执行"),
  r(["net", "localgroup"], "forbidden", "本地组管理，禁止直接执行"),
  r(["Set-ExecutionPolicy"], "allow", "修改 PS 执行策略，跟随用户权限模式"),
  r(["ConvertTo-SecureString"], "allow", "凭据构造，跟随用户权限模式"),
  r(["Get-Credential"], "allow", "凭据获取，跟随用户权限模式"),
  r(["cmdkey"], "forbidden", "凭据管理器，禁止直接执行"),

  // === Excel 主链路常用解释器（放行，由 permissionMode 决定是否确认） ===
  r(["python"], "allow"),
  r(["python3"], "allow"),
  r(["py"], "allow"),
  r(["node"], "allow"),
  r(["cscript"], "allow"),
  r(["wscript"], "allow"),
];
