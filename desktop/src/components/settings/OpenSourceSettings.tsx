/**
 * Open source project usage details.
 *
 * Lists direct runtime and build dependencies used by this desktop app.
 */

import React from "react";
import { ExternalLink, Package } from "../common/IconMap";
import { useSettingsStore } from "../../store/settingsStore";
import { ipcApi } from "../../services/ipcApi";

type OpenSourceProject = {
  name: string;
  category: "runtime" | "build" | "service" | "standard";
  purposeZh: string;
  purposeEn: string;
  license: string;
  githubUrl: string;
};

const PROJECTS: OpenSourceProject[] = [
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
    name: "better-sqlite3",
    category: "runtime",
    purposeZh: "为知识库、会话状态和本地索引提供 SQLite 访问能力。",
    purposeEn: "Provides SQLite access for knowledge base, session state, and local indexes.",
    license: "MIT",
    githubUrl: "https://github.com/WiseLibs/better-sqlite3",
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
    name: "JSZip",
    category: "runtime",
    purposeZh: "读取和生成 Office Open XML、MinerU 结果压缩包等 ZIP 内容。",
    purposeEn: "Reads and creates ZIP content for Office Open XML and MinerU result packages.",
    license: "MIT OR GPL-3.0-or-later",
    githubUrl: "https://github.com/Stuk/jszip",
  },
  {
    name: "MinerU",
    category: "service",
    purposeZh: "作为 OCR、发票识别和文件可见内容解析链路的远程文档解析能力，标准接口失败时还会降级到 MinerU 免费 Agent 轻量解析。",
    purposeEn: "Provides remote document parsing for OCR, invoice recognition, and visible file content extraction; the tool can fall back to MinerU Agent lightweight parsing when the standard API is unavailable.",
    license: "MinerU Open Source License",
    githubUrl: "https://github.com/opendatalab/mineru",
  },
  {
    name: "Office Open XML",
    category: "standard",
    purposeZh: "作为内置文件级 Word、Excel、PowerPoint 读取、检查、创建和样式修改能力的开放文档格式基础；本应用直接读写 ZIP/XML 结构，未捆绑 Open XML SDK。",
    purposeEn: "Open document format basis for built-in file-level Word, Excel, and PowerPoint inspection, creation, and styling; this app directly reads and writes ZIP/XML structures and does not bundle the Open XML SDK.",
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

const OPEN_SOURCE_TEXT = {
  "zh-CN": {
    title: "开源项目",
    desc: "本页用于声明软件中直接使用的主要开源项目、用途和项目地址。",
    noticeTitle: "开源使用声明",
    noticeBody: "以下项目为本应用的直接运行依赖、构建工具、外部解析服务或开放格式参考。各项目版权和许可证归原作者及社区所有，具体许可条款以对应仓库、官方标准和随包 LICENSE 文件为准。",
    runtime: "运行时依赖",
    service: "外部服务与开放标准",
    build: "构建与测试工具",
    project: "项目",
    purpose: "主要用途",
    license: "许可证",
    link: "GitHub",
    open: "打开 GitHub",
  },
  "en-US": {
    title: "Open Source",
    desc: "This page declares the main open source projects directly used by the app, their purpose, and project links.",
    noticeTitle: "Open Source Notice",
    noticeBody: "The projects below are direct runtime dependencies, build tools, external parsing services, or open format references used by this app. Copyrights and licenses belong to their authors and communities; the authoritative terms are the linked repositories, official standards, and bundled LICENSE files.",
    runtime: "Runtime Dependencies",
    service: "External Services And Open Standards",
    build: "Build And Test Tools",
    project: "Project",
    purpose: "Purpose",
    license: "License",
    link: "GitHub",
    open: "Open GitHub",
  },
} as const;

export const OpenSourceSettings: React.FC = () => {
  const { language } = useSettingsStore();
  const text = OPEN_SOURCE_TEXT[language];
  const runtimeProjects = PROJECTS.filter((project) => project.category === "runtime");
  const serviceProjects = PROJECTS.filter((project) => project.category === "service" || project.category === "standard");
  const buildProjects = PROJECTS.filter((project) => project.category === "build");

  const openGithub = async (url: string) => {
    await ipcApi.app.openExternal(url);
  };

  return (
    <div className="settings-section-content open-source-settings">
      <h2>{text.title}</h2>
      <p className="section-desc">{text.desc}</p>

      <div className="settings-card open-source-notice">
        <div className="settings-card-header">
          <div className="settings-card-title-row">
            <Package size={16} />
            <h3>{text.noticeTitle}</h3>
          </div>
          <p>{text.noticeBody}</p>
        </div>
      </div>

      <OpenSourceGroup
        title={text.runtime}
        projects={runtimeProjects}
        text={text}
        onOpenGithub={openGithub}
        language={language}
      />
      <OpenSourceGroup
        title={text.service}
        projects={serviceProjects}
        text={text}
        onOpenGithub={openGithub}
        language={language}
      />
      <OpenSourceGroup
        title={text.build}
        projects={buildProjects}
        text={text}
        onOpenGithub={openGithub}
        language={language}
      />
    </div>
  );
};

function OpenSourceGroup({
  title,
  projects,
  text,
  onOpenGithub,
  language,
}: {
  title: string;
  projects: OpenSourceProject[];
  text: typeof OPEN_SOURCE_TEXT["zh-CN"] | typeof OPEN_SOURCE_TEXT["en-US"];
  onOpenGithub: (url: string) => void;
  language: "zh-CN" | "en-US";
}) {
  return (
    <div className="settings-card open-source-group">
      <div className="settings-card-header">
        <h3>{title}</h3>
      </div>
      <div className="open-source-table-wrapper">
        <table className="open-source-table">
          <thead>
            <tr>
              <th>{text.project}</th>
              <th>{text.purpose}</th>
              <th>{text.license}</th>
              <th>{text.link}</th>
            </tr>
          </thead>
          <tbody>
            {projects.map((project) => (
              <tr key={project.name}>
                <td className="open-source-name">{project.name}</td>
                <td>{language === "zh-CN" ? project.purposeZh : project.purposeEn}</td>
                <td>
                  <span className="open-source-license">{project.license}</span>
                </td>
                <td>
                  <button
                    type="button"
                    className="open-source-link"
                    onClick={() => onOpenGithub(project.githubUrl)}
                    title={text.open}
                  >
                    <ExternalLink size={14} />
                    <span>{text.link}</span>
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
