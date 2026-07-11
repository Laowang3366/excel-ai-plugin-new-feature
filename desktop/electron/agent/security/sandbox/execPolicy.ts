/**
 * 命令策略引擎
 *
 * 对应 Codex `execpolicy` crate：按前缀 token 匹配命令，产出
 * allow / prompt / forbidden 决策。多规则命中取最严
 * `forbidden > prompt > allow`。
 *
 * Codex rule.rs 的核心结构（精简化后）：
 *   PrefixPattern = { first: Arc<str>, rest: Arc<[PatternToken]> }
 *   PatternToken   = Single(String) | Alts(Vec<String>)
 *   PrefixRule     = { pattern, decision, justification }
 * 这里用等价的 TS 表达。
 *
 * 本引擎不进入 OS 沙箱，只做命令语义级过滤；OS 沙箱层（阶段 3）
 * 与审批层各自独立，互不阻塞：
 *   forbidden → 直接拒绝（不入 spawn、不进审批）
 *   prompt    → 强制走用户审批对话框，无视 permissionMode 与 alwaysAllowedTools
 *   allow     → 维持原审批流程（由 toolExecutor 决定）
 */

import * as os from "os";
import * as path from "path";
import { type ParsedCommand } from "./parseCommand";

// ============================================================
// 类型
// ============================================================

/** 决策三档：与 Codex execpolicy Decision 对齐 */
export type Decision = "allow" | "prompt" | "forbidden";

/** 单 token 模式：要么精确匹配一个字符串，要么匹配一组备选 */
export type PatternToken =
  | { kind: "single"; value: string }
  | { kind: "alts"; values: string[] };

/** 前缀规则 */
export interface PrefixRule {
  /** 首 token（必须是具体字符串，与 Codex 一致用于索引） */
  first: string;
  /** 其余 token 模式 */
  rest: PatternToken[];
  /** 决策 */
  decision: Decision;
  /** 规则说明（在审批对话框/审计中展示） */
  justification?: string;
}

/** 命中一条规则的结果 */
export interface RuleHit {
  /** 命中的前缀 token 序列 */
  matchedPrefix: string[];
  rule: PrefixRule;
  /** 命中的子命令 */
  command: ParsedCommand;
}

/** 一条命令的整体评估结果 */
export interface ExecPolicyEvaluation {
  /** 合并后的最终决策 */
  decision: Decision;
  /** 所有命中的规则（按子命令顺序） */
  hits: RuleHit[];
  /** 命中 forbidden 的子命令（用于审计/告警） */
  violations: RuleHit[];
  /** 解析失败的子命令（无法评估时按 prompt 处理） */
  unparseable: ParsedCommand[];
}

// ============================================================
// 权重合并
// ============================================================

const DECISION_WEIGHT: Record<Decision, number> = { allow: 0, prompt: 1, forbidden: 2 };

/** 取最严决策 */
function stricter(a: Decision, b: Decision): Decision {
  return DECISION_WEIGHT[a] >= DECISION_WEIGHT[b] ? a : b;
}

// ============================================================
// 规则索引
// ============================================================

/**
 * 按首 token 建索引 —— Codex Policy 对 first token 做哈希索引以加速匹配。
 * 命令大小写敏感性：Windows 上 PowerShell cmdlet 大小写不敏感，shell 命令相对敏感。
 * 这里在索引与匹配时统一对 Windows 用大小写不敏感比较，其他平台大小写敏感。
 */
export class ExecPolicy {
  /** first(规范化) -> 规则数组 */
  private index = new Map<string, PrefixRule[]>();
  /** 大小写不敏感？ */
  private caseInsensitive: boolean;
  private rules: PrefixRule[];

  constructor(rules: PrefixRule[] = [], caseInsensitive = process.platform === "win32") {
    this.caseInsensitive = caseInsensitive;
    this.rules = rules;
    this.rebuildIndex();
  }

  /** 替换并重建索引 */
  setRules(rules: PrefixRule[]): void {
    this.rules = rules;
    this.rebuildIndex();
  }

  /** 追加一条规则（不去重） */
  addRule(rule: PrefixRule): void {
    this.rules.push(rule);
    const k = this.key(rule.first);
    const arr = this.index.get(k);
    if (arr) arr.push(rule);
    else this.index.set(k, [rule]);
  }

  getRules(): readonly PrefixRule[] {
    return this.rules;
  }

  private key(token: string): string {
    return this.caseInsensitive ? token.toLowerCase() : token;
  }

  private rebuildIndex(): void {
    this.index.clear();
    for (const rule of this.rules) {
      const k = this.key(rule.first);
      const arr = this.index.get(k);
      if (arr) arr.push(rule);
      else this.index.set(k, [rule]);
    }
  }

  /** 单条规则是否匹配某子命令前缀 */
  private matchRule(rule: PrefixRule, tokens: string[]): string[] | null {
    if (tokens.length === 0) return null;
    const firstMatch = this.tokenEq(tokens[0], rule.first);
    if (!firstMatch) return null;
    if (rule.rest.length === 0) return [tokens[0]];
    if (tokens.length < rule.rest.length + 1) return null;
    const matched = [tokens[0]];
    for (let i = 0; i < rule.rest.length; i++) {
      const token = tokens[i + 1];
      const pat = rule.rest[i];
      if (pat.kind === "single") {
        if (!this.tokenEq(token, pat.value)) return null;
      } else {
        if (!pat.values.some((v) => this.tokenEq(token, v))) return null;
      }
      matched.push(token);
    }
    return matched;
  }

  private tokenEq(a: string, b: string): boolean {
    return this.caseInsensitive ? a.toLowerCase() === b.toLowerCase() : a === b;
  }

  /** 评估单条子命令 */
  evaluateCommand(parsed: ParsedCommand): RuleHit[] {
    if (parsed.parseFailed) return [];
    const head = parsed.tokens[0];
    if (!head) return [];
    const candidates = this.index.get(this.key(head));
    if (!candidates) return [];
    const hits: RuleHit[] = [];
    for (const rule of candidates) {
      const matchedPrefix = this.matchRule(rule, parsed.tokens);
      if (matchedPrefix) {
        hits.push({ matchedPrefix, rule, command: parsed });
      }
    }
    return hits;
  }

  /**
   * 评估整段命令：把所有子命令的命中合并
   *
   * 解析失败的子命令视为 prompt（无法证明安全）
   */
  evaluate(parsedCommands: ParsedCommand[]): ExecPolicyEvaluation {
    const hits: RuleHit[] = [];
    const violations: RuleHit[] = [];
    const unparseable: ParsedCommand[] = [];
    let decision: Decision = "allow";

    for (const cmd of parsedCommands) {
      if (cmd.parseFailed) {
        unparseable.push(cmd);
        decision = stricter(decision, "prompt");
        continue;
      }
      if (cmd.tokens.length === 0) continue;
      const cmdHits = this.evaluateCommand(cmd);
      for (const h of cmdHits) {
        hits.push(h);
        decision = stricter(decision, h.rule.decision);
        if (h.rule.decision === "forbidden") violations.push(h);
      }
    }

    return { decision, hits, violations, unparseable };
  }
}

// ============================================================
// 工作目录白名单（对应 Codex FileSystemSandboxPolicy 可写根）
// ============================================================

/** 白名单判定结果 */
export interface CwdCheckResult {
  /** 是否允许在指定 workdir 执行 */
  allowed: boolean;
  /** 实际使用的 workdir（被重定向时与请求的 workdir 不同） */
  effectiveWorkdir: string;
  /** 是否被重定向到临时目录 */
  redirected: boolean;
}

/** 默认可写根：临时目录 + 用户文档/桌面/下载 */
export function defaultWritableRoots(): string[] {
  const home = os.homedir();
  return [
    os.tmpdir(),
    path.join(home, "Desktop"),
    path.join(home, "Documents"),
    path.join(home, "Downloads"),
  ];
}

/**
 * 规范化路径用于比较：小写、去尾分隔符
 */
function normalize(p: string): string {
  return path.resolve(p).replace(/[\\/]+$/, "").toLowerCase();
}

function isInside(parent: string, child: string): boolean {
  const pa = normalize(parent);
  const ca = normalize(child);
  return ca === pa || ca.startsWith(pa + path.sep.toLowerCase());
}

/**
 * 检查 workdir 是否在白名单内；不在则重定向到 os.tmpdir()
 *
 * @param workdir 请求的工作目录
 * @param writableRoots 白名单根目录，默认为 defaultWritableRoots()
 * @param fallback 重定向目标，默认 os.tmpdir()
 */
export function checkWorkdir(
  workdir: string,
  writableRoots: string[] = defaultWritableRoots(),
  fallback: string = os.tmpdir()
): CwdCheckResult {
  if (!workdir || typeof workdir !== "string") {
    return { allowed: false, effectiveWorkdir: fallback, redirected: true };
  }
  const ok = writableRoots.some((root) => isInside(root, workdir));
  if (ok) return { allowed: true, effectiveWorkdir: workdir, redirected: false };
  return { allowed: false, effectiveWorkdir: fallback, redirected: true };
}
