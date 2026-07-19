# AGENTS.md

ZCode 工作区指令文件。本仓库包含三个独立运行的 Node 项目，不存在根 package workspace 或根构建命令。

## 项目结构

```text
desktop/        Windows Electron + React 桌面应用、Agent Runtime、.NET 8 Office Worker
product-site/   Fastify 产品页、更新清单 API、下载统计后台、发布脚本
excel-addin/    独立 Excel Office.js / WPS JSA 任务窗格加载项（浏览器侧载验证，无 Electron/COM/.NET）
release-notes/  面向用户的版本更新日志（每个版本一个 JSON）
docs/           当前架构、开发规范、发布与历史设计文档
overview.md     简明架构总览（项目说明的入口）
README.md       用户向文档与命令清单
CHANGELOG.md    仅记录用户可感知的功能与体验变化
```

涉及 Office 实现时，以 `desktop/electron/agent/officeWorker/` 和 `desktop/dotnet/Wengge.OfficeWorker/` 为准。事实确认顺序：运行代码与 `desktop/package.json` > 根 `README.md` / `overview.md` > `docs/README.md` 索引中的当前文档。`docs/superpowers/`、`session-*.md`、早期 code review 报告和实施计划是历史归档。

## 环境要求

- Windows x64
- Node.js 22.12+（CI 使用 Node 22；Electron 43 要求至少 22.12）
- .NET 8 SDK（即使打包版带 self-contained Worker，本地构建/测试仍需要 SDK）
- 仅在当前窗口 COM 自动化与真实冒烟时需要 Microsoft Office 或 WPS Office；纯 Open XML 文件处理不要求 Office 进程

## 桌面端命令（在 `desktop/` 下执行）

```powershell
npm ci
npm run office:publish          # 发布 self-contained win-x64 Worker，开发期必跑
npm run dev                     # 不发布 Worker，Office 功能前需先 office:publish
```

质量与构建：

```powershell
npm run typecheck               # 渲染进程 + Electron 主进程类型
npm run lint
npm run format:check / format
npm test                        # Vitest
npm run build                   # Vite 构建渲染/主/preload
npm run office:build            # .NET 解决方案 Release
npm run office:test             # xUnit 测试 Release
npm run electron:build          # Worker + 清理 + Vite + x64 NSIS 安装包
npm run patch:build -- --id <id> --base-version <ver>   # 受限热补丁
npm run release:verify -- --manifest ... --public-key ... --artifact-dir ...
```

`electron:build` 会清空 `desktop/dist`、`desktop/dist-electron` 与整个 `desktop/release/`，不要在这些目录保留无关产物。`test:coverage` 脚本存在但仓库未声明 `@vitest/coverage-v8`，不要把 coverage 当作可用门槛。运行单个测试：

```powershell
npm test -- electron/agent/core/agentLoop/toolExecutor.test.ts
npm test -- <path>.test.ts -t "test name"
npm run office:test -- --filter "FullyQualifiedName~Namespace.TestClass"
```

真实 Office 冒烟（按需运行，不是默认门禁，会启动真实 Office/WPS 进程）：`test:office-smoke`、`test:word-smoke`、`test:word-lifecycle`、`test:excel-lifecycle`、`test:presentation-smoke`、`test:office-reliability`、`test:wps-routing`。可用 `WENGGE_OFFICE_SMOKE_TIMEOUT_MS` 调整默认超时。冒烟脚本必须遵守进程所有权规则，不得关闭或附加到与任务无关的用户 Office 进程。

CI 仓级 workflow（`.github/workflows/ci.yml`）含并行 job：`desktop` / `product-site` / `excel-addin`。桌面端：`npm ci` → `npm audit --audit-level=high` → `lint` → `typecheck` → `test` → `build`（及 office restore/audit/test）。CI 不跑 `format:check`、coverage、真实 Office 冒烟和 NSIS 打包；修改 Worker 或 TypeScript/Worker 协议时需本地跑 `office:test`。Tag 触发的桌面发布工作流额外跑 `electron:build`。

## 产品站命令（在 `product-site/` 下执行）

```bash
npm ci
npm run dev
npm start
npm test
node --test test/server.test.mjs
node --test --test-name-pattern="download tracking" test/server.test.mjs
npm run hash-password -- "a-password-of-at-least-12-characters"
npm run publish-release -- --version <ver> --installer <exe> --blockmap <blockmap> --latest-yml <yml> --notes-file <json> --private-key <pem> --output ./.local/releases --base-url https://plugin.shelelove.top
```

产品站无 lint / typecheck / build 脚本；`npm start` 不自动加载 `.env`，由 systemd `EnvironmentFile` 注入。产品站 CI 跑 `npm ci` + 高危 `npm audit` + `npm test`。

## Excel 加载项命令（在 `excel-addin/` 下执行）

```bash
npm ci
npm run typecheck
npm test
npm run build
npm run manifest:check
npm run package:prod -- --base-url https://example.com/excel-addin
```

Excel 加载项 CI 门禁（`.github/workflows/ci.yml` 的 `excel-addin` job）：`npm ci` → `npm audit --audit-level=high` → `manifest:check` → `typecheck` → `test` → `build`。不跑 `sync:prompts`、证书安装、真实 Office/WPS、生产部署。手动生产静态包：`.github/workflows/excel-addin-package.yml`（`workflow_dispatch` only），产出 GitHub Actions artifact（`excel-addin/dist/**`），与桌面 NSIS 发布完全分离。Ed25519 私钥（`desktop/.secrets/update-private.pem`）必须留在 Git 外。生产只监听 `127.0.0.1:18120`，由独立 Nginx 站点代理；下载使用 `X-Accel-Redirect`，不通过 Node 流式传输安装包。

## 桌面端运行时架构

```text
React 渲染进程（desktop/src）
  -> src/services/ipcApi.ts 域包装
  -> context-isolated preload（electron/preload.ts）
  -> Electron 主进程 IPC
  -> Agent Runtime 与本地服务（electron/agent、electron/main-modules）
       -> 模型 provider API
       -> 知识 / 记忆 / 文件 / Web / OCR
       -> 类型化 Office 桥（electron/agent/officeWorker）
            -> 共享子进程 Wengge.OfficeWorker.exe
                 -> Microsoft Office / WPS COM
                 -> C# DocumentFormat.OpenXml
                 -> WPS JSA localhost bridge（端口 45221，token 保护）
```

Vite 三个入口：`src/main.tsx`、`electron/main.ts`、`electron/preload.ts`。Electron 主进程里使用 `electron/shared/logger.ts`，不要用 `console.log`。Provider 请求在 Electron 主进程发起；`electron/agent/providers/aiClientFactory.ts` 选择 provider 协议。产品站不是 AI 代理。

`AgentLoopManager` 全局只允许一个活跃 turn；当前正在运行的线程追加输入会在该 loop 内排队。`npm run dev` 通过 `dotnet/publish/win-x64/` 解析 Worker，可由 `WENGGE_OFFICE_WORKER_PATH` 覆盖。

## 关键边界（不要破坏）

### 渲染进程与 IPC

- 渲染端用 Zustand store；聊天状态是主进程 Agent 事件的投影，不要在渲染端建第二份会话模型。
- 渲染端调用走 `src/services/ipcApi.ts` 与域模块，禁止散落直接调用 `window.electronAPI`。
- 修改 preload API 时同步更新：`electron/preload.ts`、`src/electronApi.d.ts`、`src/services/ipcApiTypes.ts`、渲染端包装、IPC handler、测试。
- IPC 输入必须用 `electron/shared/ipcSchemas.ts` 中的 Zod schema 在运行时校验；TypeScript 类型本身不是信任边界。
- 文件与 OCR 操作走 `electron/main-modules/ipcPathSecurity.ts` 的路径授权。
- 流式事件同时在 `electron/agent/interaction/eventForwarder.ts` 与 `src/store/chatStreamBuffer.ts` 缓冲，改顺序/时序必须同时考虑两层。

### Office 自动化

- Electron 端只做类型化桥与策略/编排，COM 与 Open XML 实现在 `desktop/dotnet/Wengge.OfficeWorker/`。
- Worker 是共享、懒启动子进程，stdio 上用换行分隔 JSON、请求 ID + 协议版本握手；允许多个 RPC in-flight，但 COM 操作在消息泵 STA 调度器上串行，Open XML 与非 COM 工作可走线程池；不要把 COM 移出 STA 调度器。
- 文件级 `.xlsx` / `.docx` / `.pptx` 优先 C# Open XML，渲染、导出、活动对象模型或不支持特性时才回退 COM。当前窗口操作附加到活动 Office/WPS 应用。
- 文件级 COM 自动化在无法建立独立 owned 进程时，不得复用、修改或关闭无关用户 Office 窗口。
- 不要重新引入 Shell / Python / PowerShell / 外部 JScript Office 自动化。已存在的危险宏工具需要审批：`macro.write` 写入当前工作簿的 VBA 或 WPS JSA，`macro.run` 只执行 VBA。
- 工具编排、备份、持久化工作流、事务、撤销/重做、冲突检测在 `electron/agent/tools/officeCore/`；Worker 只执行原语。
- 新增模型可见工具需跨层同步：`tools/registry`（schema / 风险等级）、`tools/executors`（校验与路由）、`tools/contracts`（TS 边界类型）、`tools/officeCore`（策略与事务）、`agent/officeWorker`（类型化 JSON-RPC 桥）、`dotnet/Wengge.OfficeWorker`（COM/Open XML 实现）、`electron/agent/prompts`（场景/工具提示）。仅注册 executor 不会暴露给模型，确认 `tools/registry/toolDefinitions.ts` 已收录。

### 持久化、记忆、OCR

- `sessions/YYYY/MM/DD/*.jsonl`：仅追加的会话 rollout / 审计副本。
- `sessions/state-runtime/`：SQLite 线程状态、日志、目标、记忆。
- `knowledge/knowledge.db`：独立 RAG 数据库。
- `office-backups/`：操作备份。
- `office-automation/workflows/` 与 `office-automation/transactions/`：持久化 Office 编排状态。
- `logs/`：应用日志。

运行时优先查询 SQLite，回退 JSONL；不要在未评估兼容与恢复行为前移除任一表示。可配置 data-path 迁移当前只复制 sessions / knowledge / logs，不复制 Office 备份与自动化日志，修改迁移代码时要意识到这个缺口。OCR 在 Electron 主进程而非 .NET Worker 执行；fallback 顺序为 MinerU API → MinerU Agent 端点 → 本地文档解析；发票模式可额外用当前模型 provider 解析 OCR Markdown。

### WPS JSA 桥

`desktop/public/wps-jsa-bridge/` 由 .NET Worker 复制到用户 WPS 加载项目录（首次安装需要完整重启 WPS），通过 token 保护的 localhost 服务（端口 45221）轮询。`public/wps-jsa-bridge/**` 是允许的热补丁路径，但加载项只有在 Worker 中硬编码的版本号变更时才会被重新复制——单纯改资源不会更新已安装的桥。改动安装/版本逻辑或 .NET 传输需要完整安装包。

## 打包与更新边界

打包资源分布在：`dist/`（渲染进程）、`dist-electron/`（Electron 主/preload）、`public/`（知识/图标/更新公钥/WPS 桥）、`office-worker/`（self-contained .NET Worker）。热补丁只允许覆盖 `dist/**`、`public/knowledge/**`、`public/wps-jsa-bridge/**`。修改主进程、preload、依赖、原生行为、Worker 协议或 .NET 代码必须发完整安装包。桌面端校验器与产品站签名器共用同一自定义清单格式：规范 JSON 签名、签名校验、size、SHA-256 逻辑必须两侧同步。

更新源 `https://plugin.shelelove.top`，Ed25519 签名 + SHA-256 校验。完整版本走 NSIS + `electron-updater` 覆盖安装；安装包输出 `desktop/release/Wengge-AI-Assistant-Setup-<version>.exe`，与 blockmap、`latest.yml` 一同作为 GitHub Release 资产（不要提交为 Git 大文件）。GitHub 桌面发布工作流不会跑产品站发布或生成 `release.json` 与 Ed25519 签名 `manifest.json`；`product-site/scripts/publish-release.mjs` 是独立步骤。

## 代码约定

- 测试与源文件同目录 `{source}.test.ts`；.NET xUnit 测试在 `desktop/dotnet/Wengge.OfficeWorker.Tests/`。常规单元测试不得启动真实 Office 进程，COM/WPS 验证用专项冒烟脚本。
- 当前文档化的单文件行数上限：通用 TS/TSX 400 行、React 组件 300 行、Zustand store 400 行、CSS 500 行。按运行时职责拆分，不要写宽泛的兼容桶（barrel）。
- 用户可见行为变更同步更新 `CHANGELOG.md`；架构、Office、构建或发布边界变更同步更新对应 `docs/` 下当前文档。
- 忽略生成目录：`node_modules`、`dist`、`dist-electron`、`release`、`dotnet/**/bin`、`dotnet/**/obj`、`dotnet/publish`、coverage、本地 data 路径、`desktop/.secrets`。除非任务明确涉及生成产物。

## 进入敏感区域前必读

- `desktop/electron/agent/`：先读 `docs/architecture-map.md` 与 `docs/development-standards.md`。
- `desktop/dotnet/Wengge.OfficeWorker/`：先读 `docs/office-advanced-automation.md`。
- 桌面打包与发布：`docs/update-and-release.md` + `.github/workflows/`。
- 产品站部署：`docs/product-site-deployment.md`。
- 系统提示词与场景引导：`docs/codex-system-prompt-architecture.md`。
- 代码风格与审查：`docs/development-standards.md` + `docs/code-review-standards.md`。