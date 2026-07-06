import { ipcApi } from "../services/ipcApi";
import { readFileAsBase64 } from "../utils/fileBase64";
import type { AttachedFile } from "../electronApi";

const MIME_TO_EXT: Record<string, string> = {
  "image/png": ".png",
  "image/jpeg": ".jpg",
  "image/webp": ".webp",
  "image/bmp": ".bmp",
  "image/gif": ".gif",
};

const IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".webp", ".bmp", ".gif"]);

function getFileExtension(file: File): string {
  const nameExt = file.name.includes(".") ? `.${file.name.split(".").pop()?.toLowerCase()}` : "";
  if (nameExt && nameExt !== ".undefined") return nameExt;
  return MIME_TO_EXT[file.type] || "";
}

export function getAttachmentFileType(file: File): AttachedFile["fileType"] {
  const ext = getFileExtension(file);
  return IMAGE_EXTENSIONS.has(ext) || file.type.startsWith("image/") ? "image" : "document";
}

function getLocalPathForFile(file: File): string {
  const electronPath = (file as File & { path?: string }).path;
  if (electronPath) return electronPath;
  try {
    return ipcApi.file.getPathForFile?.(file) || "";
  } catch {
    return "";
  }
}

async function fileToTemporaryAttachment(file: File): Promise<AttachedFile | null> {
  try {
    const suffix = getFileExtension(file) || (file.type === "application/pdf" ? ".pdf" : ".bin");
    const base64 = await readFileAsBase64(file);
    const result = await ipcApi.file.writeTempFile({
      prefix: getAttachmentFileType(file) === "image" ? "image" : "attachment",
      suffix,
      data: base64,
    });
    if (!result.success || !result.filePath) return null;
    return {
      filePath: result.filePath,
      fileName: file.name || result.filePath.split(/[\\/]/).pop() || "attachment",
      fileType: getAttachmentFileType(file),
      size: file.size,
    };
  } catch {
    return null;
  }
}

export async function resolveDroppedFiles(files: File[]): Promise<AttachedFile[]> {
  const resolved: AttachedFile[] = [];
  for (const file of files) {
    const filePath = getLocalPathForFile(file);
    if (filePath) {
      resolved.push({
        filePath,
        fileName: file.name || filePath.split(/[\\/]/).pop() || filePath,
        fileType: getAttachmentFileType(file),
        size: file.size,
      });
      continue;
    }

    const temporaryAttachment = await fileToTemporaryAttachment(file);
    if (temporaryAttachment) {
      resolved.push(temporaryAttachment);
    }
  }
  return resolved;
}
