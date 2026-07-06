/**
 * Preload 脚本 — 安全的 IPC 桥接
 *
 * 通过 contextBridge 暴露有限的 API 给渲染进程，
 * 避免直接暴露 Node.js / Electron 能力。
 */

import { contextBridge, ipcRenderer, webUtils } from "electron";

contextBridge.exposeInMainWorld("electronAPI", {
  // ---- 应用 ----
  app: {
    getDataPath: () => ipcRenderer.invoke("app:getDataPath"),
    selectDataPath: () => ipcRenderer.invoke("app:selectDataPath"),
    migrateDataPath: (targetPath: string) =>
      ipcRenderer.invoke("app:migrateDataPath", targetPath),
    openPath: (targetPath: string) => ipcRenderer.invoke("app:openPath", targetPath),
    openExternal: (targetUrl: string) => ipcRenderer.invoke("app:openExternal", targetUrl),
  },

  // ---- 窗口 ----
  window: {
    getAlwaysOnTop: () => ipcRenderer.invoke("window:getAlwaysOnTop"),
    setAlwaysOnTop: (enabled: boolean) =>
      ipcRenderer.invoke("window:setAlwaysOnTop", enabled),
    getDisplayMode: () => ipcRenderer.invoke("window:getDisplayMode"),
    setDisplayMode: (mode: "normal" | "compact") =>
      ipcRenderer.invoke("window:setDisplayMode", mode),
    onDisplayModeChanged: (callback: (mode: "normal" | "compact") => void) => {
      const handler = (_event: any, mode: "normal" | "compact") => callback(mode);
      ipcRenderer.on("window:displayModeChanged", handler);
      return () => ipcRenderer.removeListener("window:displayModeChanged", handler);
    },
  },

  // ---- 设置 ----
  settings: {
    get: (key: string) => ipcRenderer.invoke("settings:get", key),
    set: (key: string, value: unknown) => ipcRenderer.invoke("settings:set", key, value),
    getAll: () => ipcRenderer.invoke("settings:getAll"),
  },

  // ---- Excel 连接状态 + 数据操作 ----
  excel: {
    detectStatus: () => ipcRenderer.invoke("excel:detectStatus"),
    connect: () => ipcRenderer.invoke("excel:connect"),
    /** 当 Office + WPS 同时运行时，用户选择目标宿主 */
    selectHost: (host: "excel" | "wps") => ipcRenderer.invoke("excel:selectHost", host),
    getSelection: () => ipcRenderer.invoke("excel:getSelection"),
    getSelectionAddress: () => ipcRenderer.invoke("excel:getSelectionAddress"),
    readRange: (sheetName: string, range: string, expand?: "none" | "spill" | "currentArray" | "currentRegion") =>
      ipcRenderer.invoke("excel:readRange", sheetName, range, expand),
    inspectWorkbook: () => ipcRenderer.invoke("excel:inspectWorkbook"),
    writeRange: (sheetName: string, range: string, values: unknown[][]) =>
      ipcRenderer.invoke("excel:writeRange", sheetName, range, values),
  },

  // ---- Word/PPT 连接状态 ----
  office: {
    detectWordStatus: () => ipcRenderer.invoke("word:detectStatus"),
    detectPresentationStatus: () => ipcRenderer.invoke("ppt:detectStatus"),
  },

  // ---- Agent 对话 ----
  agent: {
    /** 启动一个新的 Turn */
    startTurn: (input: { content: string; attachments?: Array<{ filePath: string; fileName: string; fileType: "image" | "document"; size?: number }>; clientId?: string; threadId?: string | null; isResume?: boolean; resumeContext?: string }) =>
      ipcRenderer.invoke("agent:startTurn", input),

    /** 继续当前 Turn（从中断恢复或追问） */
    continueTurn: (input: { content: string; attachments?: Array<{ filePath: string; fileName: string; fileType: "image" | "document"; size?: number }>; clientId?: string; threadId?: string | null }) =>
      ipcRenderer.invoke("agent:continueTurn", input),

    /** 运行中补充输入：当前 turn 结束后自动处理 */
    enqueueTurn: (input: { content: string; attachments?: Array<{ filePath: string; fileName: string; fileType: "image" | "document"; size?: number }>; clientId?: string; threadId?: string | null }) =>
      ipcRenderer.invoke("agent:enqueueTurn", input),

    /** 中断当前 Turn */
    interrupt: (threadId?: string | null) => ipcRenderer.invoke("agent:interrupt", { threadId }),

    /** 监听 Agent 事件 */
    onEvent: (callback: (event: any) => void) => {
      const handler = (_event: any, data: any) => callback(data);
      ipcRenderer.on("agent:event", handler);
      return () => ipcRenderer.removeListener("agent:event", handler);
    },

    /** 监听流式增量 */
    onStreamDelta: (callback: (data: { delta: string; itemType: string; roundId?: number; threadId?: string; clientId?: string }) => void) => {
      const handler = (_event: any, data: any) => callback(data);
      ipcRenderer.on("agent:streamDelta", handler);
      return () => ipcRenderer.removeListener("agent:streamDelta", handler);
    },
  },

  // ---- 会话（Thread）管理 ----
  thread: {
    /** 列出所有会话 */
    list: () => ipcRenderer.invoke("thread:list"),

    /** 加载某个会话的完整数据 */
    load: (threadId: string) => ipcRenderer.invoke("thread:load", threadId),

    /** 删除某个会话 */
    delete: (threadId: string) => ipcRenderer.invoke("thread:delete", threadId),

    /** 恢复某个会话（使其成为活跃会话） */
    resume: (threadId: string) => ipcRenderer.invoke("thread:resume", threadId),

    /** 新建会话（重置 Agent 线程，下次 startTurn 自动创建新线程） */
    newThread: (folderId?: string) => ipcRenderer.invoke("thread:new", folderId),

    /** 更新会话元数据（如 folderId、name 等） */
    updateMetadata: (threadId: string, patch: Record<string, unknown>) =>
      ipcRenderer.invoke("thread:updateMetadata", threadId, patch),

    /** 查找最近更新的会话 ID */
    findLatest: () => ipcRenderer.invoke("thread:findLatest"),

    /** 查看 Agent 内存中的线程运行态 */
    runtimeStatus: () => ipcRenderer.invoke("thread:runtimeStatus"),
  },

  // ---- 线程拓扑图 ----
  threadGraph: {
    upsertSpawnEdge: (parentThreadId: string, childThreadId: string, label?: string) =>
      ipcRenderer.invoke("threadGraph:upsertSpawnEdge", { parentThreadId, childThreadId, label }),
    closeSpawnEdge: (parentThreadId: string, childThreadId: string) =>
      ipcRenderer.invoke("threadGraph:closeSpawnEdge", { parentThreadId, childThreadId }),
    listDescendants: (parentThreadId: string, status?: "open" | "closed" | "all") =>
      ipcRenderer.invoke("threadGraph:listDescendants", { parentThreadId, status }),
  },

  // ---- 文件对话框 ----
  dialog: {
    /** 打开文件选择对话框（文档类型） */
    openFile: () => ipcRenderer.invoke("dialog:openFile"),

    /** 打开图片选择对话框 */
    openImage: () => ipcRenderer.invoke("dialog:openImage"),

    /** 打开文件夹选择对话框 */
    openFolder: () => ipcRenderer.invoke("dialog:openFolder"),
  },

  // ---- 文件读取 + 文件操作 ----
  file: {
    /** 读取文件为 base64 */
    readAsBase64: (filePath: string) =>
      ipcRenderer.invoke("file:readAsBase64", filePath),
    /** 移动文件到系统回收站 */
    trashFile: (filePath: string) =>
      ipcRenderer.invoke("file:trashFile", filePath),
    /** 用系统默认应用打开文件 */
    openFile: (filePath: string) =>
      ipcRenderer.invoke("file:openFile", filePath),
    /** 将文件绝对路径复制到剪贴板 */
    copyPath: (filePath: string) =>
      ipcRenderer.invoke("file:copyPath", filePath),
    /** 在系统文件管理器中显示文件 */
    revealInExplorer: (filePath: string) =>
      ipcRenderer.invoke("file:revealInExplorer", filePath),
    /** 将 base64 数据写入临时文件（截图粘贴等） */
    writeTempFile: (data: { prefix?: string; suffix?: string; data: string }) =>
      ipcRenderer.invoke("file:writeTempFile", data),
    /** 获取拖拽/粘贴 File 对象对应的本地路径（Electron 新版替代 File.path） */
    getPathForFile: (file: any) => {
      const filePath = webUtils.getPathForFile(file);
      if (filePath) {
        ipcRenderer.sendSync("file:authorizePathSync", filePath);
      }
      return filePath;
    },
  },

  // ---- 文件夹操作 ----
  folder: {
    /** 列出文件夹内的 Office 文件（Excel/Word/PowerPoint） */
    listFiles: (folderPath: string) =>
      ipcRenderer.invoke("folder:listFiles", folderPath),
  },

  // ---- 工具列表 ----
  tools: {
    /** 获取所有可用的工具定义 */
    list: () => ipcRenderer.invoke("tools:list"),
  },

  // ---- 沙箱策略配置 ----
  sandbox: {
    /** 获取当前策略配置（默认规则 + 用户规则 + 可写根） */
    getConfig: () => ipcRenderer.invoke("sandbox:getConfig"),
    /** 更新用户自定义规则（覆盖式） */
    setUserRules: (rules: unknown[]) => ipcRenderer.invoke("sandbox:setUserRules", rules),
    /** 更新额外可写根目录（覆盖式） */
    setWritableRoots: (roots: string[]) => ipcRenderer.invoke("sandbox:setWritableRoots", roots),
  },

  // ---- 工具确认 ----
  tool: {
    /** 确认执行挂起的工具调用 */
    confirm: (toolCallId: string, alwaysAllow?: boolean) =>
      ipcRenderer.invoke("tool:confirm", toolCallId, alwaysAllow),
    /** 取消挂起的工具调用 */
    cancel: (toolCallId: string) =>
      ipcRenderer.invoke("tool:cancel", toolCallId),
  },

  // ---- AI 模型列表 + 连接测试 ----
  ai: {
    /** 获取可用模型列表 */
    listModels: (baseUrl: string, apiKey: string, apiFormat: string) =>
      ipcRenderer.invoke("ai:listModels", baseUrl, apiKey, apiFormat),
    /** 测试 API 连接 */
    testConnection: (baseUrl: string, apiKey: string, apiFormat: string, model: string) =>
      ipcRenderer.invoke("ai:testConnection", baseUrl, apiKey, apiFormat, model),
  },

  // ---- 使用统计 ----
  stats: {
    /** 获取聚合的使用统计（单次 IPC 调用，替代 N+1 的 thread:load） */
    getSummary: () => ipcRenderer.invoke("stats:getSummary"),
  },

  // ---- OCR / 视觉识别 ----
  ocr: {
    recognize: (mode: string, filePaths: string[]) =>
      ipcRenderer.invoke("ocr:recognize", mode, filePaths),
  },

  // ---- 知识库 (RAG) ----
  knowledge: {
    /** 列出所有已索引的知识来源 */
    listSources: () => ipcRenderer.invoke("knowledge:listSources"),
    /** 搜索知识库 */
    search: (query: string, topK?: number) =>
      ipcRenderer.invoke("knowledge:search", query, topK),
    /** 索引单个文件 */
    indexFile: (filePath: string) =>
      ipcRenderer.invoke("knowledge:indexFile", filePath),
    /** 索引文件夹 */
    indexFolder: (folderPath: string) =>
      ipcRenderer.invoke("knowledge:indexFolder", folderPath),
    /** 删除文件索引 */
    deleteFile: (sourcePath: string) =>
      ipcRenderer.invoke("knowledge:deleteFile", sourcePath),
    /** 重建全部索引 */
    reindexAll: () => ipcRenderer.invoke("knowledge:reindexAll"),
  },
});
