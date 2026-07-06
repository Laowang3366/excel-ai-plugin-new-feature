import { ipcApi } from "../../services/ipcApi";
import { readFileAsBase64 } from "../../utils/fileBase64";

export const ACCEPTED_OCR_TYPES = "image/png, image/jpeg, image/webp, image/bmp, application/pdf";

const ACCEPTED_OCR_MIME_TYPES = new Set(ACCEPTED_OCR_TYPES.split(", "));

export async function resolveOcrFilePaths(files: File[]): Promise<string[]> {
  const paths: string[] = [];
  for (const file of files) {
    const existingPath = (file as File & { path?: string }).path;
    if (existingPath) {
      paths.push(existingPath);
      continue;
    }

    const suffix = file.name.includes(".")
      ? `.${file.name.split(".").pop()}`
      : file.type === "application/pdf"
      ? ".pdf"
      : ".png";
    const result = await ipcApi.file.writeTempFile({
      prefix: "ocr",
      suffix,
      data: await readFileAsBase64(file),
    });
    if (!result.success || !result.filePath) {
      throw new Error(result.error || `无法读取 OCR 文件: ${file.name}`);
    }
    paths.push(result.filePath);
  }
  return paths;
}

export async function resolveWriteTarget(input: string): Promise<{ sheetName: string; range: string }> {
  const trimmed = input.trim();
  if (trimmed) {
    const parsed = parseSheetRange(trimmed);
    if (parsed.sheetName && parsed.range) return parsed;
    const selection = await ipcApi.excel.getSelectionAddress();
    return { sheetName: selection.sheetName, range: parsed.range || trimmed };
  }

  const selection = await ipcApi.excel.getSelectionAddress();
  return { sheetName: selection.sheetName, range: selection.address };
}

export function parseSheetRange(value: string): { sheetName: string; range: string } {
  const bangIndex = value.lastIndexOf("!");
  if (bangIndex < 0) return { sheetName: "", range: value };
  const rawSheetName = value.slice(0, bangIndex).trim();
  return {
    sheetName: rawSheetName.replace(/^'(.*)'$/, "$1"),
    range: value.slice(bangIndex + 1).trim(),
  };
}

export function isAcceptedOcrFile(file: File): boolean {
  const name = file.name.toLowerCase();
  return ACCEPTED_OCR_MIME_TYPES.has(file.type) ||
    name.endsWith(".pdf") ||
    name.endsWith(".png") ||
    name.endsWith(".jpg") ||
    name.endsWith(".jpeg") ||
    name.endsWith(".webp") ||
    name.endsWith(".bmp");
}

export function isLikelyInvoiceFile(file: File): boolean {
  return /发票|invoice|fapiao|票据/i.test(file.name);
}
