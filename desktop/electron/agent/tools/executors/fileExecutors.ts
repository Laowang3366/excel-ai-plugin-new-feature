/**
 * 文件上下文工具执行器
 *
 * 只负责暴露用户常用路径和当前会话文件夹路径。
 */

import * as os from "os";
import * as path from "path";
import type { ToolExecutor } from "../../shared/types";

export interface FileExecutorDeps {
  sessionFolderPath?: string;
}

export function addFileExecutors(target: Map<string, ToolExecutor>, deps: FileExecutorDeps): void {
  target.set("file.getPaths", {
    name: "file.getPaths",
    execute: async (args: Record<string, unknown>) => {
      const requested = (args.pathNames as string[]) || ["desktop", "documents", "downloads"];
      const homeDir = os.homedir();
      const userName = os.userInfo().username;

      const pathMap: Record<string, { name: string; path: string }> = {
        desktop: { name: "桌面", path: path.join(homeDir, "Desktop") },
        documents: { name: "文档", path: path.join(homeDir, "Documents") },
        downloads: { name: "下载", path: path.join(homeDir, "Downloads") },
        pictures: { name: "图片", path: path.join(homeDir, "Pictures") },
        music: { name: "音乐", path: path.join(homeDir, "Music") },
        videos: { name: "视频", path: path.join(homeDir, "Videos") },
        appData: { name: "AppData", path: path.join(homeDir, "AppData", "Roaming") },
        home: { name: "用户主目录", path: homeDir },
        temp: { name: "临时目录", path: os.tmpdir() },
      };

      const results: Array<{ name: string; key: string; path: string }> = [];
      for (const key of requested) {
        const entry = pathMap[key.toLowerCase()];
        if (entry) {
          results.push({ name: entry.name, key, path: entry.path });
        }
      }

      if (deps.sessionFolderPath) {
        const folderName = deps.sessionFolderPath.split(/[\\/]/).pop() || "当前工作文件夹";
        results.unshift({
          name: `当前工作文件夹（${folderName}）`,
          key: "current_folder",
          path: deps.sessionFolderPath,
        });
      }

      if (results.length === 0) {
        for (const [key, entry] of Object.entries(pathMap)) {
          results.push({ name: entry.name, key, path: entry.path });
        }
      }

      return { success: true, data: { paths: results, userName } };
    },
  });
}
