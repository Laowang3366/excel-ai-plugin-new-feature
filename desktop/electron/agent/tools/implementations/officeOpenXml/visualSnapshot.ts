/**
 * Office 视觉快照选择器
 *
 * 关联模块：
 * - officeOpenXmlFileBridge.ts: 将快照能力暴露给工具执行器。
 * - prompts/templates/scenarios/office-tools.zh-CN.md: 引导模型先请求快照再优化排版。
 */

import { execFile } from "child_process";
import path from "path";
import { promisify } from "util";
import type { OfficeOpenXmlDocumentType, OfficeVisualSnapshotInput, OfficeVisualSnapshotResult } from "./types";

const execFileAsync = promisify(execFile);
const HEADLESS_RENDERER_CANDIDATES = ["soffice", "libreoffice"];

export function selectSnapshotPlan(input: {
  preferEngine?: "openxml" | "com";
  hasHeadlessRenderer: boolean;
  hasComFallback: boolean;
}): { engine: "openxml" | "com"; reason: string } {
  if (input.preferEngine === "com") {
    if (!input.hasComFallback) {
      throw new Error("没有可用的 Office 视觉快照渲染器：COM 不可用");
    }
    return { engine: "com", reason: "用户明确指定 COM" };
  }

  if (input.hasHeadlessRenderer) {
    return { engine: "openxml", reason: "检测到 headless Office 渲染器" };
  }

  if (input.hasComFallback) {
    return { engine: "com", reason: "headless 渲染器不可用，使用 COM 兜底" };
  }

  throw new Error("没有可用的 Office 视觉快照渲染器");
}

function detectDocumentType(filePath: string): OfficeOpenXmlDocumentType {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".docx") return "word";
  if (ext === ".pptx") return "presentation";
  if (ext === ".xlsx") return "spreadsheet";
  throw new Error(`仅支持 .docx、.pptx 和 .xlsx 文件: ${filePath}`);
}

function defaultOutputPath(filePath: string): string {
  const dir = path.dirname(filePath);
  const base = path.basename(filePath, path.extname(filePath));
  return path.join(dir, `${base}-snapshot.png`);
}

async function findHeadlessRenderer(): Promise<string | undefined> {
  for (const candidate of HEADLESS_RENDERER_CANDIDATES) {
    try {
      await execFileAsync(candidate, ["--version"], { timeout: 2000 });
      return candidate;
    } catch {
      // 继续探测下一个候选命令。
    }
  }
  return undefined;
}

export async function createOfficeVisualSnapshot(input: OfficeVisualSnapshotInput): Promise<OfficeVisualSnapshotResult> {
  const documentType = detectDocumentType(input.filePath);
  const outputPath = input.outputPath || defaultOutputPath(input.filePath);
  const renderer = await findHeadlessRenderer();

  try {
    const plan = selectSnapshotPlan({
      preferEngine: input.preferEngine,
      hasHeadlessRenderer: Boolean(renderer),
      hasComFallback: false,
    });

    // TODO: 接入 headless Office/COM 导出后，把 supported 改为真实渲染结果。
    return {
      engine: plan.engine,
      operation: "snapshot",
      documentType,
      filePath: input.filePath,
      outputPath,
      target: input.target,
      renderer,
      supported: false,
      error: "已选择快照渲染路径，但 Open XML/headless 实际导出将在后续渲染适配中接入",
    };
  } catch (error) {
    return {
      engine: "openxml",
      operation: "snapshot",
      documentType,
      filePath: input.filePath,
      outputPath,
      target: input.target,
      supported: false,
      error: error instanceof Error ? error.message : String(error),
    };
  }
}
