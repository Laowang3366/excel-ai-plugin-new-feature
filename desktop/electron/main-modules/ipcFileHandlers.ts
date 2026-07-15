import { BrowserWindow, clipboard, dialog, shell } from "electron";
import { trustedIpcMain as ipcMain } from "../shared/trustedIpc";
import * as fs from "fs";
import * as path from "path";
import {
  FilePathInput,
  FileWriteTempFileInput,
  FolderPathInput,
  FolderPathsInput,
  IPC_MAX_FILE_TRANSFER_BYTES,
  validateInput,
} from "../shared/ipcSchemas";
import { assertAuthorizedPath, type PathAuthorizer } from "./ipcPathSecurity";
import { guardDataOperation } from "./dataMaintenance";

export interface FolderFileInfo {
  filePath: string;
  fileName: string;
  size: number;
  lastModified: number;
}

const OFFICE_FILE_EXTENSIONS = new Set([".xlsx", ".xls", ".csv", ".doc", ".docx", ".ppt", ".pptx"]);

interface RegisterFileIpcHandlersOptions {
  mainWindowRef: () => BrowserWindow | null;
  pathAuthorizer: PathAuthorizer;
  getDataPath: () => string;
  isDataMaintenanceInProgress?: () => boolean;
}

export async function listAuthorizedOfficeFiles(folderPath: string, pathAuthorizer: PathAuthorizer): Promise<FolderFileInfo[]> {
  const authorizedFolderPath = assertAuthorizedPath(pathAuthorizer, folderPath);
  const entries = await fs.promises.readdir(authorizedFolderPath, { withFileTypes: true });
  const files = entries
    .filter((entry) => entry.isFile() && OFFICE_FILE_EXTENSIONS.has(path.extname(entry.name).toLowerCase()))
    .map((entry) => {
      const fullPath = path.join(authorizedFolderPath, entry.name);
      pathAuthorizer.authorizePath(fullPath);
      return { filePath: fullPath, fileName: entry.name };
    });
  const results = await Promise.all(
    files.map(async (file) => {
      try {
        const stat = await fs.promises.stat(file.filePath);
        return { ...file, size: stat.size, lastModified: stat.mtimeMs };
      } catch {
        return { ...file, size: 0, lastModified: 0 };
      }
    })
  );
  results.sort((a, b) => a.fileName.localeCompare(b.fileName));
  return results;
}

export async function assertFileWithinIpcTransferLimit(
  filePath: string,
  maxBytes = IPC_MAX_FILE_TRANSFER_BYTES
): Promise<void> {
  const stat = await fs.promises.stat(filePath);
  if (!stat.isFile()) throw new Error("目标路径不是文件");
  if (stat.size > maxBytes) {
    throw new Error(`文件过大，IPC 传输最大支持 ${Math.floor(maxBytes / 1024 / 1024)}MB`);
  }
}

export async function writeManagedTempFile(
  data: unknown,
  options: Pick<RegisterFileIpcHandlersOptions,
    "getDataPath" | "pathAuthorizer" | "isDataMaintenanceInProgress">,
): Promise<{ success: boolean; filePath?: string; error?: string }> {
  try {
    const input = validateInput(FileWriteTempFileInput, data);
    return await guardDataOperation(options.isDataMaintenanceInProgress, async () => {
      const prefix = input.prefix?.replace(/[^a-zA-Z0-9_-]/g, "") || "clipboard";
      const suffix = input.suffix?.replace(/[^a-zA-Z0-9.]/g, "") || ".png";
      const tmpDir = path.join(options.getDataPath(), "temp");
      await fs.promises.mkdir(tmpDir, { recursive: true });
      const fileName = `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}${suffix}`;
      const filePath = path.join(tmpDir, fileName);
      await fs.promises.writeFile(filePath, Buffer.from(input.data, "base64"));
      options.pathAuthorizer.authorizePath(filePath);
      return { success: true, filePath };
    })();
  } catch (err: any) {
    return { success: false, error: err.message };
  }
}

export function registerFileIpcHandlers(options: RegisterFileIpcHandlersOptions): void {
  const { mainWindowRef, pathAuthorizer, getDataPath } = options;

  ipcMain.handle("dialog:openFile", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] as string[] };
    const result = await dialog.showOpenDialog(mw, {
      properties: ["openFile"],
      filters: [
        { name: "Documents", extensions: ["xlsx", "xls", "csv", "doc", "docx", "ppt", "pptx", "json", "txt", "pdf", "md"] },
        { name: "All Files", extensions: ["*"] },
      ],
    });
    result.filePaths.forEach((filePath) => pathAuthorizer.authorizePath(filePath));
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  ipcMain.handle("dialog:openImage", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] as string[] };
    const result = await dialog.showOpenDialog(mw, {
      properties: ["openFile"],
      filters: [
        { name: "Images", extensions: ["png", "jpg", "jpeg", "webp", "bmp", "gif"] },
      ],
    });
    result.filePaths.forEach((filePath) => pathAuthorizer.authorizePath(filePath));
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  ipcMain.handle("dialog:openFolder", async () => {
    const mw = mainWindowRef();
    if (!mw) return { canceled: true, filePaths: [] as string[] };
    const result = await dialog.showOpenDialog(mw, {
      title: "选择文件夹",
      properties: ["openDirectory"],
    });
    result.filePaths.forEach((folderPath) => pathAuthorizer.authorizeRoot(folderPath));
    return { canceled: result.canceled, filePaths: result.filePaths };
  });

  ipcMain.handle("folder:listFiles", async (_event, folderPath: unknown) => {
    const validated = validateInput(FolderPathInput, folderPath);
    try {
      return await listAuthorizedOfficeFiles(validated, pathAuthorizer);
    } catch {
      return [];
    }
  });

  ipcMain.handle("folder:listFilesBatch", async (_event, folderPaths: unknown) => {
    const validated = validateInput(FolderPathsInput, folderPaths);
    const entries = await Promise.all(
      validated.map(async (folderPath) => {
        try {
          return [folderPath, await listAuthorizedOfficeFiles(folderPath, pathAuthorizer)] as const;
        } catch {
          return [folderPath, []] as const;
        }
      })
    );
    return Object.fromEntries(entries);
  });

  ipcMain.on("file:authorizePathSync", (event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      pathAuthorizer.authorizePath(validated);
      event.returnValue = { success: true };
    } catch (err: any) {
      event.returnValue = { success: false, error: err?.message || "授权路径失败" };
    }
  });

  ipcMain.handle("file:readAsBase64", async (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      const authorizedFilePath = assertAuthorizedPath(pathAuthorizer, validated);
      await assertFileWithinIpcTransferLimit(authorizedFilePath);
      const buffer = await fs.promises.readFile(authorizedFilePath);
      const ext = path.extname(authorizedFilePath).toLowerCase().replace(".", "");
      const mimeMap: Record<string, string> = {
        png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
        gif: "image/gif", webp: "image/webp", bmp: "image/bmp",
        pdf: "application/pdf", csv: "text/csv", json: "application/json",
        txt: "text/plain", md: "text/markdown", xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        xls: "application/vnd.ms-excel",
        doc: "application/msword",
        docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ppt: "application/vnd.ms-powerpoint",
        pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      };
      const mimeType = mimeMap[ext] || "application/octet-stream";
      return {
        data: buffer.toString("base64"),
        mimeType,
        fileName: path.basename(authorizedFilePath),
        size: buffer.length,
      };
    } catch (err: any) {
      return { error: err.message };
    }
  });

  ipcMain.handle("file:writeTempFile", async (_event, data: unknown) => {
    return writeManagedTempFile(data, options);
  });

  ipcMain.handle("file:trashFile", async (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      await shell.trashItem(assertAuthorizedPath(pathAuthorizer, validated));
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("file:openFile", async (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      const result = await shell.openPath(assertAuthorizedPath(pathAuthorizer, validated));
      if (result) return { success: false, error: result };
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle("file:copyPath", (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      clipboard.writeText(assertAuthorizedPath(pathAuthorizer, validated));
      return { success: true };
    } catch {
      return { success: false };
    }
  });

  ipcMain.handle("file:revealInExplorer", (_event, filePath: unknown) => {
    try {
      const validated = validateInput(FilePathInput, filePath);
      shell.showItemInFolder(assertAuthorizedPath(pathAuthorizer, validated));
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });
}
