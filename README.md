# Office AI 桌面助手

面向 Excel / Word / PowerPoint 与 WPS Office 的本地桌面 AI 助手。应用运行在 Electron 桌面端，通过模型工具调用自主读取、编辑、验证办公文件，并优先使用 Open XML 文件级引擎，必要时再使用 COM 作为兜底。

## 功能概览

- **Office 三件套自主编辑**：支持 Excel/WPS 表格、Word/WPS 文字、PowerPoint/WPS 演示的读取、编辑、保存和验证。
- **统一 Office 操作入口**：模型优先调用 `office.action.inspect` / `office.action.apply` / `office.action.validate`，避免为 Word/PPT 临场拼 PowerShell 或 Python 脚本。
- **Open XML 优先**：`.xlsx` / `.docx` / `.pptx` 文件级编辑不依赖 Office 进程；PPT 删除页、创建演示文稿、文本替换、表格样式、主题色、Excel 图表/条件格式/数据验证等走统一能力。
- **COM 兜底**：动态图表、目录刷新、快照导出、当前窗口交互等需要 Office 应用对象模型的场景走 COM 桥接。
- **多模型供应商**：支持 OpenAI 兼容协议、Anthropic、DeepSeek、Kimi、智谱、小米、阿里云百炼、腾讯云、火山方舟、讯飞星辰、百度千帆、京东云等配置。
- **本地运行态存储**：使用 `better-sqlite3` 的四库 StateRuntime（`state.db` / `logs.db` / `goals.db` / `memories.db`），JSONL 保留为兼容审计副本。
- **长期记忆**：重点记忆用户偏好、规则约束、纠正反馈、过往文件印象和工具成功率画像。
- **权限与沙箱**：`shell.execute` 走命令策略、工作目录约束、环境变量清洗和审计日志；高风险操作仍需要用户确认。
- **会话体验**：支持文件夹组织、会话搜索、运行中会话状态感知、输入队列、长消息窗口化渲染和侧边功能面板。

## 快速开始

前置条件：

- Windows
- Node.js 20+
- 已安装 Microsoft Office 或 WPS Office（仅 Open XML 文件级编辑不强依赖运行中的 Office 进程）

```bash
cd desktop
npm install
```

| 命令 | 说明 |
|------|------|
| `npm run dev` | 启动 Vite 开发环境 |
| `npm run typecheck` | 渲染进程 + Electron 主进程类型检查 |
| `npm test` | 运行单元测试 |
| `npm run build` | 构建渲染进程 |
| `npm run electron:build` | 构建 Windows 安装包 |
| `npm run native:rebuild` | 为 Electron 运行时重建原生依赖 |

> 注意：`better-sqlite3` 是原生模块。跑 Node/Vitest 测试前可执行 `npm rebuild better-sqlite3`；打包时 electron-builder 会为 Electron ABI 重建依赖。

## 项目结构

```text
.
├─ README.md                         # 项目总览与架构说明
├─ CHANGELOG.md                      # 版本变更记录
├─ docs/                             # 开发规范、设计方案、审查记录和阶段日志
│  ├─ development-standards.md       # 当前开发规范
│  ├─ dev-log.md                     # 主要开发过程日志
│  └─ superpowers/                   # 阶段计划与方案归档
└─ desktop/                          # Electron 桌面应用
   ├─ package.json                   # 桌面端脚本、依赖和 electron-builder 配置
   ├─ electron/                      # Electron 主进程
   │  ├─ main.ts                     # 主进程入口
   │  ├─ preload.ts                  # preload API 桥
   │  ├─ main-modules/               # 设置、窗口、IPC、事件转发等主进程模块
   │  └─ agent/                      # Agent 运行时
   │     ├─ interaction/             # 交互层：IPC、事件转发、审批回调
   │     ├─ runtime/                 # 装配层：模型、工具、记忆、知识库、压缩运行时
   │     ├─ core/agentLoop/          # 核心层：会话轮次、流式收集、工具循环、压缩、中断
   │     ├─ tools/                   # 工具层：注册表、契约、执行器、Office/Excel 实现
   │     │  ├─ registry/             # 模型可见工具定义
   │     │  ├─ contracts/            # 实现无关契约
   │     │  ├─ executors/            # 工具路由与结果封装
   │     │  └─ implementations/      # Excel COM、Office COM、Open XML 引擎
   │     ├─ memory/                  # 记忆层：SQLite StateRuntime、JSONL 审计、压缩、长期记忆
   │     ├─ knowledge/               # 知识层：本地文件解析、索引、检索
   │     ├─ prompts/                 # 提示词层：系统提示词、Office 工具选择、记忆模板
   │     ├─ providers/               # 模型供应商层
   │     ├─ security/sandbox/        # 命令沙箱与审计
   │     ├─ automation/              # PowerShell、Python、JScript、JSON 基础能力
   │     ├─ attachments/             # 图片等附件解析
   │     └─ shared/                  # Agent 内共享类型与消息构建
   ├─ src/                           # React 渲染进程
   │  ├─ components/                 # 页面、侧边栏、聊天、设置、任务面板组件
   │  ├─ store/                      # Zustand 状态管理
   │  ├─ services/ipcApi.ts          # IPC 抽象层
   │  ├─ hooks/                      # 输入框、连接状态、草稿等 hooks
   │  └─ styles/                     # 按功能域拆分的样式
   └─ python/                        # 嵌入式 Python 运行时说明与安装脚本
```

## Agent 分层职责

```text
interaction -> runtime -> core
core -> providers / prompts / memory / knowledge / tools / security / shared
tools/executors -> tools/contracts + tools/registry + 注入的实现
tools/implementations -> automation + tools/contracts
memory -> shared
knowledge -> 本层存储与 embedding
providers -> shared
prompts -> shared
```

- `core` 只做 Agent 编排，不直接依赖 Excel/Word/PPT 具体实现。
- `tools/registry` 只放模型可见 schema 和风险等级。
- `tools/contracts` 只放实现无关接口。
- `tools/executors` 负责参数校验、工具路由和结果封装。
- `tools/implementations` 才承载 COM、Open XML、PowerShell 等实现细节。
- `memory` 和 `knowledge` 不反向依赖核心循环。

## Office 操作路线

### 统一入口

| 意图 | 推荐工具 |
|------|----------|
| 检查文件结构、表格、布局 | `office.action.inspect` |
| 修改 Excel/Word/PPT 文件 | `office.action.apply` |
| 验证输出文件和对象变化 | `office.action.validate` |
| 当前窗口交互 | `word.*` / `presentation.*` / Excel 专用工具 |
| 专用工具覆盖不到的复杂自动化 | `office.script.execute` |

### Open XML 优先能力

- Excel：插入图表、条件格式、数据验证、表格样式、文本替换。
- Word：标题样式、表格样式、页眉/页脚、文本替换。
- PPT：创建基础演示、主题色、删除指定页、文本替换。

### COM 兜底能力

- Excel：图表、条件格式、数据验证、表格样式。
- Word：目录插入/刷新、图片插入、快照导出、窗口内编辑。
- PPT：快照导出、图表、图片占位、形状对齐、版式规范化、删除页兜底。

## 存储与记忆

当前运行态以 SQLite 为主：

| 数据库 | 职责 |
|--------|------|
| `state.db` | 会话元数据、线程状态、活跃/卸载状态 |
| `logs.db` | rollout 事件、工具执行日志、全文检索索引 |
| `goals.db` | 目标、预算和完成状态 |
| `memories.db` | 长期记忆、命名空间、记忆管道游标 |

JSONL rollout 仍作为兼容审计副本保留，并支持后台归档压缩。长期记忆默认面向办公场景，只把用户偏好、规则、纠错、文件印象和工具成功率画像作为核心记忆。

## 测试与验证

```bash
cd desktop
npm run typecheck
npm test
```

当前基线：

- `npm run typecheck`：渲染进程与 Electron 主进程均通过。
- `npm test`：74 个测试文件，420 个测试。
- `npm run electron:build`：生成 `release/Office AI 助手 Setup <version>.exe`。

## 开发规范

详见 [docs/development-standards.md](docs/development-standards.md)。

核心约束：

- 禁止过度设计、过度兜底、过度约束边界导致功能异常。
- 遵循模块单一职责，但不要为了拆分而拆碎。
- 文件按架构层和职责分类存放。
- 注释说明当前模块职责和关联模块。
- 每个阶段完成后必须 review，确认未破坏原业务、未引出新问题，再进入下一阶段。
- 非必须、非必要的测试/烟测临时文件，验证后必须清理。
- 每个任务 review 通过后提交 git。

## License

Private - 仅供内部使用
