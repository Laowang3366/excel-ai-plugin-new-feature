import { z } from "zod";

export const IPC_MAX_PATH_CHARS = 32_767;
export const IPC_MAX_CHAT_CONTENT_CHARS = 50_000;
export const IPC_MAX_RESUME_CONTEXT_CHARS = 200_000;
export const IPC_MAX_ATTACHMENTS = 20;
export const IPC_MAX_OCR_FILES = 20;
export const IPC_MAX_EXCEL_CELLS = 100_000;
export const IPC_MAX_EXCEL_ROWS = 10_000;
export const IPC_MAX_EXCEL_COLUMNS = 16_384;
export const IPC_MAX_CELL_TEXT_CHARS = 32_767;
export const IPC_MAX_FILE_TRANSFER_BYTES = 50 * 1024 * 1024;
export const IPC_MAX_BASE64_CHARS = Math.ceil(IPC_MAX_FILE_TRANSFER_BYTES / 3) * 4;

export const IpcPath = z.string().min(1, "路径不能为空").max(IPC_MAX_PATH_CHARS, "路径过长");

export function estimateBase64DecodedBytes(value: string): number {
  if (!value) return 0;
  const padding = value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0;
  return Math.max(0, Math.floor((value.length * 3) / 4) - padding);
}

export function isBase64PayloadWithinLimit(
  value: string,
  maxBytes = IPC_MAX_FILE_TRANSFER_BYTES,
): boolean {
  return estimateBase64DecodedBytes(value) <= maxBytes;
}

/** 运行时校验工具：校验失败抛出带有字段描述的 Error */
export function validateInput<T>(schema: z.ZodType<T>, data: unknown): T {
  const result = schema.safeParse(data);
  if (!result.success) {
    const issues = result.error.issues.map((i) => `${i.path.join(".")}: ${i.message}`).join("; ");
    throw new Error(`IPC 参数校验失败: ${issues}`);
  }
  return result.data;
}
