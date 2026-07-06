import { BrowserWindow, clipboard, dialog, ipcMain, shell } from "electron";
import * as fs from "fs";
import * as path from "path";
import {
  FilePathInput,
  FileWriteTempFileInput,
  FolderPathInput,
  FolderPathsInput,
  validateInput,
} from "../shared/ipcSchemas";
import { assertAuthorizedPath, type PathAuthorizer } from "./ipcPathSecurity";

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

export function registerFileIpcHandlers(options: RegisterFileIpcHandlersOptions): void {
  const { mainWindowRef, pathAuthorizer } = options;

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
    try {
      const input = validateInput(FileWriteTempFileInput, data);
      const prefix = input.prefix?.replace(/[^a-zA-Z0-9_-]/g, "") || "clipboard";
      const suffix = input.suffix?.replace(/[^a-zA-Z0-9.]/g, "") || ".png";
      const tmpDir = await import("os").then((os) => os.tmpdir());
      const fileName = `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}${suffix}`;
      const filePath = path.join(tmpDir, fileName);
      const buffer = Buffer.from(input.data, "base64");
      await fs.promises.writeFile(filePath, buffer);
      pathAuthorizer.authorizePath(filePath);
      return { success: true, filePath };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
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
