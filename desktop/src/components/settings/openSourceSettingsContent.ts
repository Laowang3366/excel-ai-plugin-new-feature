export type OpenSourceProject = {
  name: string;
  category: "runtime" | "build" | "service" | "standard";
  purposeZh: string;
  purposeEn: string;
  license: string;
  githubUrl: string;
};

export const OPEN_SOURCE_PROJECTS: OpenSourceProject[] = [
  {
    name: "Electron",
    category: "runtime",
    purposeZh: "承载桌面端窗口、主进程、预加载脚本和系统能力桥接。",
    purposeEn: "Desktop runtime for windows, main process, preload bridge, and OS integration.",
    license: "MIT",
    githubUrl: "https://github.com/electron/electron",
  },
  {
    name: "React",
    category: "runtime",
    purposeZh: "构建聊天、设置、任务面板等前端界面。",
    purposeEn: "Builds the chat, settings, and task panel UI.",
    license: "MIT",
    githubUrl: "https://github.com/facebook/react",
  },
  {
    name: "React DOM",
    category: "runtime",
    purposeZh: "将 React 界面渲染到桌面端 WebView DOM。",
    purposeEn: "Renders the React UI into the desktop WebView DOM.",
    license: "MIT",
    githubUrl: "https://github.com/facebook/react",
  },
  {
    name: "Zustand",
    category: "runtime",
    purposeZh: "管理设置、会话列表、聊天状态等前端状态。",
    purposeEn: "Manages frontend state such as settings, chats, and sessions.",
    license: "MIT",
    githubUrl: "https://github.com/pmndrs/zustand",
  },
  {
    name: "Zod",
    category: "runtime",
    purposeZh: "校验 IPC 输入和结构化参数，减少运行时类型错误。",
    purposeEn: "Validates IPC inputs and structured parameters to reduce runtime type errors.",
    license: "MIT",
    githubUrl: "https://github.com/colinhacks/zod",
  },
  {
    name: "SQLite",
    category: "runtime",
    purposeZh:
      "通过 Node/Electron 内置 node:sqlite 为知识库、会话状态和本地索引提供 SQLite 存储能力。",
    purposeEn:
      "Provides SQLite storage for knowledge base, session state, and local indexes through Node/Electron built-in node:sqlite.",
    license: "blessing",
    githubUrl: "https://github.com/sqlite/sqlite",
  },
  {
    name: "electron-store",
    category: "runtime",
    purposeZh: "持久化本地设置、模型配置、窗口偏好等数据。",
    purposeEn: "Persists local settings, model configuration, and window preferences.",
    license: "MIT",
    githubUrl: "https://github.com/sindresorhus/electron-store",
  },
  {
    name: "fflate",
    category: "runtime",
    purposeZh: "读取 MinerU 结果包并生成经过白名单约束的热更新 ZIP。",
    purposeEn: "Reads MinerU result packages and creates allowlisted hot-update ZIP archives.",
    license: "MIT",
    githubUrl: "https://github.com/101arrowz/fflate",
  },
  {
    name: "MinerU",
    category: "service",
    purposeZh:
      "作为 OCR、发票识别和文件可见内容解析链路的远程文档解析能力，标准接口失败时还会降级到 MinerU 免费 Agent 轻量解析。",
    purposeEn:
      "Provides remote document parsing for OCR, invoice recognition, and visible file content extraction; the tool can fall back to MinerU Agent lightweight parsing when the standard API is unavailable.",
    license: "MinerU Open Source License",
    githubUrl: "https://github.com/opendatalab/mineru",
  },
  {
    name: "Office Open XML",
    category: "standard",
    purposeZh:
      "作为 Word、Excel、PowerPoint 文件级读取、检查、创建和样式修改能力的开放文档格式基础；由 .NET Worker 和 DocumentFormat.OpenXml 实现。",
    purposeEn:
      "Open document format basis for file-level Word, Excel, and PowerPoint inspection, creation, and styling through the .NET Worker and DocumentFormat.OpenXml.",
    license: "ECMA-376 / ISO/IEC 29500",
    githubUrl: "https://github.com/dotnet/Open-XML-SDK",
  },
  {
    name: "Lucide React",
    category: "runtime",
    purposeZh: "提供设置、侧边栏、工具按钮等界面图标。",
    purposeEn: "Provides UI icons for settings, sidebar, and tool buttons.",
    license: "ISC",
    githubUrl: "https://github.com/lucide-icons/lucide",
  },
  {
    name: "react-markdown",
    category: "runtime",
    purposeZh: "渲染模型回复中的 Markdown 内容。",
    purposeEn: "Renders Markdown content in assistant replies.",
    license: "MIT",
    githubUrl: "https://github.com/remarkjs/react-markdown",
  },
  {
    name: "remark-gfm",
    category: "runtime",
    purposeZh: "支持 GitHub Flavored Markdown 表格、任务列表和自动链接。",
    purposeEn: "Adds GitHub Flavored Markdown support for tables, task lists, and autolinks.",
    license: "MIT",
    githubUrl: "https://github.com/remarkjs/remark-gfm",
  },
  {
    name: "Vite",
    category: "build",
    purposeZh: "负责前端开发服务器和生产构建。",
    purposeEn: "Provides the frontend dev server and production build pipeline.",
    license: "MIT",
    githubUrl: "https://github.com/vitejs/vite",
  },
  {
    name: "electron-builder",
    category: "build",
    purposeZh: "打包 Windows 安装程序并处理 Electron 应用发布产物。",
    purposeEn: "Packages the Windows installer and Electron distribution artifacts.",
    license: "MIT",
    githubUrl: "https://github.com/electron-userland/electron-builder",
  },
  {
    name: "Vitest",
    category: "build",
    purposeZh: "运行单元测试和回归测试。",
    purposeEn: "Runs unit and regression tests.",
    license: "MIT",
    githubUrl: "https://github.com/vitest-dev/vitest",
  },
  {
    name: "TypeScript",
    category: "build",
    purposeZh: "提供静态类型检查和 Electron/前端编译约束。",
    purposeEn: "Provides static type checking and compile-time constraints.",
    license: "Apache-2.0",
    githubUrl: "https://github.com/microsoft/TypeScript",
  },
];
