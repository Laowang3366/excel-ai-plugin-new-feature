/**
 * 沙箱审计 — 对应 Codex `windows-sandbox-rs/src/logging.rs`
 *
 * 每次 shell.execute 命中策略、被重定向、被拒绝、被审批，都追加一行 JSONL 到
 * 与 SessionStore 同根目录的 sandbox-logs/YYYY/MM/DD/audit-YYYY-MM-DD.jsonl。
 *
 * 设计取舍：
 * - 不复用 sessionStore 的 rollout 文件（修行某种不同语义，混在一起难排查）
 * - 异步写入、批量 flush 不必：写入频率极低（每次 shell.execute），直接 appendFile 即可
 * - 字段保持稳定，便于 jq 解析
 */

import * as fs from "fs";
import * as path from "path";
import { promisify } from "util";
import * as os from "os";

const appendFile = promisify(fs.appendFile);
const mkdir = promisify(fs.mkdir);

/** 审计根目录 —— 与 sessionStore 同根，便于捆绑排查 */
function getRoot(): string {
  const appData = process.env.APPDATA || path.join(process.env.USERPROFILE || os.homedir(), "AppData", "Roaming");
  return path.join(appData, "excel-ai-assistant", "sandbox-logs");
}

/** 一次审计事件 */
export interface AuditEvent {
  /** ISO 时间戳 */
  ts: string;
  /** 事件类型 */
  type:
    | "decision" // 策略层决策
    | "violation" // forbidden 命中
    | "workdir_redirect" // workdir 被重定向
    | "approval" // 审批流程
    | "execute" // spawn 执行
    | "execute_failure"; // spawn 失败
  /** 命令原始文本 */
  command?: string;
  /** 命中的规则 hits（精简） */
  hits?: Array<{
    matchedPrefix: string[];
    decision: string;
    justification?: string;
  }>;
  /** 最终决策 */
  decision?: string;
  /** 请求的 workdir */
  requestedWorkdir?: string;
  /** 实际 workdir（可能被重定向） */
  effectiveWorkdir?: string;
  /** 是否被重定向 */
  redirected?: boolean;
  /** 用户审批结果 */
  approved?: boolean;
  /** 执行退出码 */
  exitCode?: number;
  /** 错误简述 */
  error?: string;
  /** 任何附加信息 */
  meta?: Record<string, unknown>;
}

let rootCache: string | null = null;

async function getDayFile(): Promise<string> {
  const root = getRoot();
  if (rootCache !== root) rootCache = root;
  const now = new Date();
  const y = now.getFullYear().toString();
  const m = (now.getMonth() + 1).toString().padStart(2, "0");
  const d = now.getDate().toString().padStart(2, "0");
  const dir = path.join(root, y, m, d);
  await mkdir(dir, { recursive: true });
  return path.join(dir, `audit-${y}-${m}-${d}.jsonl`);
}

/**
 * 记一条审计事件到当天 JSONL 文件
 *
 * 失败不抛错（审计本身不应阻塞业务），仅 console.warn。
 */
export async function audit(event: AuditEvent): Promise<void> {
  try {
    const line = JSON.stringify(event) + "\n";
    const file = await getDayFile();
    await appendFile(file, line, "utf8");
  } catch (err) {
    // eslint-disable-next-line no-console
    console.warn("[sandbox.audit] write failed:", err);
  }
}