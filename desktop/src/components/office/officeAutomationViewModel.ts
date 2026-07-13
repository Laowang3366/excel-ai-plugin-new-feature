import type { OfficeAutomationApp } from "../../electronApi";

export function officeAppLabel(app: OfficeAutomationApp): string {
  if (app === "excel") return "Excel / WPS 表格";
  if (app === "word") return "Word / WPS 文字";
  return "PowerPoint / WPS 演示";
}

export function officeStatusLabel(status: string): string {
  return ({ running: "运行中", paused: "已暂停", done: "已完成", failed: "失败", cancelled: "已取消", pending: "准备中", applied: "已应用", undone: "已撤销", conflicted: "有冲突", skipped: "已跳过" } as Record<string, string>)[status] || status;
}

export function shortOfficePath(filePath?: string): string {
  if (!filePath) return "未保存文档";
  return filePath.split(/[\\/]/).pop() || filePath;
}

export function parseTemplateVariables(value: string): Record<string, unknown> {
  const parsed = JSON.parse(value || "{}") as unknown;
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) throw new Error("运行参数必须是 JSON 对象");
  return parsed as Record<string, unknown>;
}

export function formatOfficeTime(value: string): string {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
}
