# 代码审查报告 — Excel AI 插件项目

> **审查日期**：2026-07-05
> **审查依据**：`docs/code-review-standards.md`
> **审查范围**：desktop/ 全部核心源码（electron/ + src/）
> **审查维度**：安全性 · 正确性 · 可维护性 · 性能 · 测试 · 项目特定规范
> **测试基线**：105 文件 / 609 测试全绿 · typecheck 双通过

---

## 一、审查总结

### 整体评价

项目工程质量**中上**。架构分层清晰、沙箱设计专业、流式增量缓冲和增量持久化等核心模式成熟。但存在两类系统性问题：

1. **IPC 安防线有漏洞** — 多个处理外部输入的 IPC 通道缺少 Zod 校验和路径限制，是最高优先级风险
2. **文件行数大规模超标** — 29 个 TS/TSX 文件 + 4 个 CSS 文件超过规范上限，最严重的 `agentLoop.ts` 达 1276 行

### 问题统计

| 级别 | 数量 | 说明 |
|------|------|------|
| 🔴 P0 | **22** | 必须修复后才能合并 |
| 🟡 P1 | **30** | 应该修复 |
| 💭 P2 | **5** | 酌情处理 |
| ✅ 亮点 | **14** | 值得学习的好实践 |

---

## 修复进度（持续更新）

> 记录规则：每完成一个审查项，回写修复范围、验证证据和关联提交，避免只在代码里静默修复。

### 2026-07-05 — P0 安全：IPC 校验与路径授权

**状态**：✅ 已修复

**关联提交**：本节所在提交 `fix: harden ipc validation and file paths`

**覆盖项**：
- S1：为 `file:writeTempFile`、`ocr:recognize`、`sandbox:setUserRules`、`sandbox:setWritableRoots`、`app:openPath`、`app:openExternal`、`excel:selectHost`、`file:readAsBase64`、`agent:interrupt` 增加 Zod schema 校验。
- S2：接入已存在但未使用的 schema，覆盖 `window:setAlwaysOnTop`、`settings:get/set`、`excel:readRange/writeRange`、`thread:load/delete/resume/new`、`file:trashFile/openFile/copyPath/revealInExplorer`、`folder:listFiles`、`tool:confirm/cancel`、`stats:getSummary`。
- S5：新增主进程路径授权器，文件读取/打开/回收站/显示/文件夹枚举/数据目录打开均先校验授权范围。

**业务链路保护**：
- 授权来源包括：当前数据目录、系统临时目录、用户通过文件/图片/文件夹对话框选择的路径、设置中的 `pinnedFolders`、沙箱额外可写根。
- 拖拽/粘贴文件仍走 `webUtils.getPathForFile(file)`，preload 在返回路径前同步登记授权，避免附件预览、OCR、文件上传出现异步竞态。
- OCR 付费 MinerU → 免费 MinerU → 本地 fallback 的降级顺序未改。

**验证证据**：
- `npm exec vitest run electron/main-modules/ipcPathSecurity.test.ts electron/shared/ipcSchemas.test.ts electron/main-modules/ipcHandlers.ocr.test.ts src/hooks/useComposer.test.ts`
- `npm exec vitest run electron/agent/interaction/ipcAgentHandlers.test.ts electron/agent/interaction/eventForwarder.test.ts`
- `npm run typecheck`

### 2026-07-05 — P0 可维护性：`agentLoop.ts` 阶段性拆分（M1 进行中）

**状态**：🚧 进行中（阶段性拆分已提交，M1 尚未关闭）

**关联提交**：本节所在提交 `refactor: split agent loop helper modules`

**覆盖范围**：
- 从 `agentLoop.ts` 抽出上下文历史收集与用量事件：`electron/agent/core/agentLoop/contextUsage.ts`。
- 从 `agentLoop.ts` 抽出流结束后的 `TurnItem` 补发与落库顺序：`electron/agent/core/agentLoop/streamResultItems.ts`。
- 从 `agentLoop.ts` 抽出压缩进度事件、压缩参数 rollout 记录和冷 rollout 归档触发：`electron/agent/core/agentLoop/compactionProgress.ts`。
- 从 `agentLoop.ts` 抽出线程回调绑定、线程快照/运行态持久化、rollout 事件 sink 绑定和长期记忆抽取调度：`electron/agent/core/agentLoop/threadRuntime.ts`。
- 从 `agentLoop.ts` 抽出 pre-turn / mid-turn 压缩执行逻辑：`electron/agent/core/agentLoop/compactionRunner.ts`。
- 清理 `agentLoop.ts` 中已变成“转一手”的流式结果、压缩进度薄委托包装，减少重复跳转。
- 新增对应单元测试，保护上下文顺序、流式结果事件顺序、压缩成功/失败事件和归档阈值行为。
- 同步更新 `electron/agent/core/agentLoop/README.md`，记录拆分后的模块职责。

**业务链路保护**：
- `AgentLoop` 对外 API 和主流程签名不变，保留原私有方法作为委托入口，降低调用链变更风险。
- reasoning、assistant_message、tool_call 的补发顺序保持原逻辑；工具执行、压缩触发、长期记忆写入和线程运行态写入链路未改。
- 此阶段只降低主循环文件复杂度，`agentLoop.ts` 仍超过 400 行，后续还需要继续拆主循环编排、线程运行态和压缩执行逻辑。

**验证证据**：
- `npm exec vitest run electron/agent/core/agentLoop/streamResultItems.test.ts electron/agent/core/agentLoop/compactionProgress.test.ts electron/agent/core/agentLoop/contextUsage.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/threadRuntime.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`
- `npm exec vitest run electron/agent/core/agentLoop/compactionRunner.test.ts electron/agent/core/agentLoop/agentLoop.test.ts`

---

## 二、🔴 P0 问题清单（必须修复）

### 安全性（8 项）

---

#### 🔴 S1 — IPC 通道缺少 Zod Schema 校验

**位置**：`electron/main-modules/ipcHandlers.ts` 多处

**问题**：以下处理外部输入的 IPC 通道未使用 Zod schema 校验，仅依赖手写检查或直接消费参数：

| 通道 | 行号 | 风险 |
|------|------|------|
| `file:writeTempFile` | 422 | 接收 `{prefix,suffix,data}`，仅手写正则清洗 |
| `ocr:recognize` | 693 | 接收 `mode/filePaths`，仅 `normalizeOcrFilePaths` 手写检查 |
| `sandbox:setUserRules` | 490 | 仅 `Array.isArray` 检查 |
| `sandbox:setWritableRoots` | 503 | 仅 `Array.isArray` 检查 |
| `app:openPath` | 150 | 无任何校验 |
| `app:openExternal` | 154 | 有 URL 协议检查，但无 Zod |
| `excel:selectHost` | 272 | 无校验 |
| `file:readAsBase64` | 394 | 无校验 |
| `agent:interrupt` | ipcAgentHandlers.ts:198 | 无校验 |

**原因**：项目规范 PR3 明确要求"新 IPC 通道有 Zod schema"为 P0。`ipcSchemas.ts` 已定义了 `validateInput` 工具和多个 schema，但上述通道未接入。

**建议**：为上述通道在 `ipcSchemas.ts` 补充 Zod schema，在 handler 首行调用 `validateInput`。`file:writeTempFile` 和 `ocr:recognize` 处理文件路径/外部内容，优先级最高。

---

#### 🔴 S2 — IPC Schema 定义未接入（形同虚设）

**位置**：`electron/main-modules/ipcHandlers.ts` + `electron/agent/interaction/ipcAgentHandlers.ts`

**问题**：`ipcSchemas.ts` 中已定义的 schema 未被对应 handler 使用：

| Schema | 对应通道 | 状态 |
|--------|----------|------|
| `SetAlwaysOnTopInput` | `window:setAlwaysOnTop` | ❌ 未调用 |
| `SettingsGetInput/SettingsSetInput` | `settings:get/set` | ❌ 未调用 |
| `ExcelReadRangeInput/ExcelWriteRangeInput` | `excel:readRange/writeRange` | ❌ 未调用 |
| `ThreadIdInput` | `thread:load/delete/resume` | ❌ 未调用 |
| `ThreadNewInput` | `thread:new` | ❌ 未调用 |
| `FilePathInput` | `file:trashFile/openFile/copyPath/revealInExplorer` | ❌ 未调用 |
| `FolderPathInput` | `folder:listFiles` | ❌ 未调用 |
| `ToolConfirmInput/ToolCancelInput` | `tool:confirm/cancel` | ❌ 未调用 |
| `StatsGetSummaryInput` | `stats:getSummary` | ❌ 未调用 |

**原因**：schema 定义了却没调用，给人一种"已校验"的错觉，比"没定义 schema"更危险。

**建议**：在所有已定义 schema 的 handler 首行接入 `validateInput`，或删除未使用的 schema。

---

#### 🔴 S5 — 文件操作缺少路径穿越防护

**位置**：`electron/main-modules/ipcHandlers.ts`

**问题**：以下 IPC 通道接受任意文件路径，无路径范围限制：

```
file:readAsBase64 (394行)    — fs.promises.readFile(filePath)
file:trashFile    (442行)    — shell.trashItem(filePath)
file:openFile     (451行)    — shell.openPath(filePath)
file:revealInExplorer (470行)— shell.showItemInFolder(filePath)
folder:listFiles  (366行)    — fs.promises.readdir(folderPath)
app:openPath      (150行)    — shell.openPath(targetPath)
```

**原因**：虽然当前渲染进程受 `contextIsolation` 保护，但如果渲染进程被 XSS 攻击（如 Markdown 渲染漏洞），攻击者可通过这些 IPC 通道读取/删除/打开系统任意文件。

**建议**：
- 对用户可见的文件操作（trashFile/openFile），限制在用户通过 dialog 选择的路径范围内
- 对 `file:readAsBase64`，限制只允许读取临时目录或用户选择目录下的文件
- 对 `folder:listFiles`，校验 folderPath 在已授权的 pinnedFolders 范围内
- 至少对路径做 `path.resolve()` 后检查是否在允许的根目录下

---

#### 🔴 S3 — Electron 安全配置

**位置**：`electron/main.ts:128`、`electron/main-modules/windowManager.ts:73-74`

**状态**：✅ **通过** — 两个 BrowserWindow 均配置 `contextIsolation: true, nodeIntegration: false`

> 列出此项是为了确认检查通过，无需修复。

---

### 可维护性 — 文件行数超标（14 项）

> 规范上限：TS/TSX ≤ 400 行，React 组件 ≤ 300 行，Store ≤ 400 行，CSS ≤ 500 行

#### 🔴 M1 — 超标文件 TOP 15（完整列表见附录）

| 文件 | 行数 | 上限 | 超出 | 拆分建议 |
|------|------|------|------|----------|
| `electron/agent/core/agentLoop/agentLoop.ts` | **1276** | 400 | +876 | 按职责拆为 runner/streamHandler/snapshot |
| `electron/main-modules/ipcHandlers.ts` | **1057** | 400 | +657 | 按业务域拆为 threadIpc/settingsIpc/excelIpc 等 |
| `src/components/Sidebar.tsx` | **775** | 300 | +475 | 拆出 SidebarCollapsed/Expanded/SortMenu |
| `src/store/settingsStore.ts` | **749** | 400 | +349 | PROVIDER_TEMPLATES 常量抽到独立文件 |
| `src/services/ipcApi.ts` | **745** | 400 | +345 | 按域拆分（agentApi/threadApi/settingsApi） |
| `electron/agent/memory/sessionStore.ts` | **674** | 400 | +274 | — |
| `electron/agent/memory/stateRuntimeStore.ts` | **671** | 400 | +271 | — |
| `electron/agent/tools/implementations/officeOpenXml/advancedExcel.ts` | **662** | 400 | +262 | — |
| `src/store/chatStore.ts` | **623** | 400 | +223 | action 抽到 chatActions.ts |
| `src/components/task/OCRTaskComposerPanel.tsx` | **600** | 300 | +300 | 拆分子表单组件 |
| `electron/agent/core/agentLoop/toolExecutor.ts` | **566** | 400 | +166 | — |
| `src/styles/settings.css` | **1553** | 500 | +1053 | 按子组件区块拆分 |
| `src/styles/sidebar.css` | **1101** | 500 | +601 | — |
| `src/styles/chat.css` | **841** | 500 | +341 | — |
| `src/styles/composer.css` | **674** | 500 | +174 | — |

**完整超标清单**：共 29 个 TS/TSX 文件 + 4 个 CSS 文件（详见附录 A）

---

## 三、🟡 P1 问题清单（应该修复）

### 性能（6 项）

---

#### 🟡 P1-perf — Sidebar 搜索触发 N+1 查询

**位置**：`src/components/Sidebar.tsx:210-220`

```typescript
// 当前：每个 pinnedFolder 一次 IPC
pinnedFolders.forEach((folder) => {
  if (folderFiles[folder.path]) return;
  ipcApi.folder.listFiles(folder.path).then(...);  // ← N+1
});
```

**原因**：用户固定的文件夹越多，IPC 调用次数线性增长。

**建议**：主进程新增 `folder:listFilesBatch(paths: string[])` 批量接口。

---

#### 🟡 P4-perf — ChatPage useEffect Promise 未取消

**位置**：`src/components/ChatPage.tsx:87-95`

```typescript
useEffect(() => {
  if (currentFolderId) {
    ipcApi.folder.listFiles(currentFolderId).then((files) => {
      setCurrentFolderFiles(files);  // ← 组件卸载后仍可能被调用
    });
  }
}, [currentFolderId]);
```

**建议**：使用 cancelled flag 模式：
```typescript
useEffect(() => {
  let cancelled = false;
  if (currentFolderId) {
    ipcApi.folder.listFiles(currentFolderId).then((files) => {
      if (!cancelled) setCurrentFolderFiles(files);
    });
  }
  return () => { cancelled = true; };
}, [currentFolderId]);
```

---

#### 🟡 P4-perf — useExcelConnection 6 处 setTimeout 未清理

**位置**：`src/hooks/useExcelConnection.ts:63, 69, 72, 77, 93, 97, 102`

**问题**：6 处 `setTimeout` 用于 pulseDot/connectFailed 状态重置，均未在卸载时 `clearTimeout`。

**建议**：用 `useRef` 保存 timer id，在 `useEffect` cleanup 中统一清理。

---

#### 🟡 P3-perf — Sidebar 渲染内未 memo 的排序数组

**位置**：`src/components/Sidebar.tsx:429-461`

**问题**：每次 render 都重新创建 `compareText`、`sortThreads`、`ungroupedThreads`、`groupedByFolder`，当 threads > 50 时排序开销显著。

**建议**：用 `useMemo` 包裹排序结果，依赖 `[threads, pinnedFolders, conversationSortMode, ...]`。

---

#### 🟡 P5-perf — Sidebar 多处内联箭头函数

**位置**：`src/components/Sidebar.tsx:543, 547, 645-647, 765-770`

**问题**：内联箭头函数导致子组件 memo 失效。

**建议**：用 `useCallback` 提取回调。

---

#### 🟡 P2-perf — settingsStore 已正确使用增量持久化 ✅

> 此项为**亮点确认**，非问题。`savePartial` 实现完整，所有 setX action 均使用增量写入。

---

### 可维护性（15 项）

---

#### 🟡 M3 — FolderSection 组件 21 个 Props

**位置**：`src/components/sidebar/FolderSection.tsx:35-61`

**问题**：Props 数量远超 10 个上限，是典型 prop drilling。

**建议**：将文件右键菜单相关 6 个 props 合并为 `fileContextMenuApi` 对象；线程回调合并为 `threadActions` 对象。

---

#### 🟡 M8 — electron/ 目录 13 处 console.warn/error 未用 logger

| 文件 | 行号 | 当前 | 建议 |
|------|------|------|------|
| `electron/main.ts` | 121 | `console.error("Fatal startup error:", err)` | `logger.error` |
| `electron/main-modules/settingsManager.ts` | 97, 177 | `console.warn` | `logger.warn` |
| `electron/agent/security/sandbox/audit.ts` | 91 | `console.warn` | `logger.warn` |
| `electron/agent/runtime/knowledgeRuntime.ts` | 76 | `console.warn` | `logger.warn` |
| `electron/agent/runtime/bridgeRegistry.ts` | 143 | `console.warn` | `logger.warn` |
| `electron/agent/core/agentLoop/agentLoop.ts` | 539, 555, 565, 579, 1128 | `console.warn` × 5 | `logger.warn` |

> **注意**：`desktop/src/` 目录 ✅ 完全无 console.log 使用。

---

### 测试（5 项）

---

#### 🟡 T3 — 74 个 electron/agent 源文件无测试

**高风险无测试文件**（节选）：

| 文件 | 风险说明 |
|------|----------|
| `security/sandbox/parseCommand.ts` | 沙箱命令解析核心 |
| `security/sandbox/execPolicy.ts` | 策略评估引擎 |
| `security/sandbox/defaultRules.ts` | 默认安全规则 |
| `core/agentLoop/turnRunner.ts` | 轮次调度 |
| `core/agentLoop/threadStateManager.ts` | 线程状态机 |
| `knowledge/textChunker.ts` | 文本分块 |
| `knowledge/retriever.ts` | 知识检索 |
| `tools/executors/pythonExecutor.ts` | Python 执行器 |

**建议**：优先为 `security/sandbox/*`、`core/agentLoop` 状态机、`knowledge` 分块/检索补测试。

---

#### 🟡 T3 — src/ 下 45 个文件无测试

**优先补测试的纯函数**：
- `src/utils/reasoningSupport.ts`
- `src/utils/textCleaner.ts`

---

#### 🟡 T5/PR2 — 全项目 0 个 @MOCK_INTERFACE 标记

**问题**：规范要求 Mock 数据必须有 `@MOCK_INTERFACE` 注释，但全仓库 grep 返回 0 匹配。

**涉及文件**：
- `src/store/chatStore.test.ts:3-31`
- `src/services/ipcApi.test.ts:14-19`
- `electron/main-modules/mineruOcr.test.ts:51`
- `electron/agent/interaction/eventForwarder.test.ts:14`
- `electron/agent/providers/openaiResponsesClient.test.ts:13`
- `electron/agent/tools/executors/shellExecutor.test.ts:7`

**建议**：在所有 `vi.mock` / `vi.fn` / `fetchMock` 处补 `// @MOCK_INTERFACE` 注释。

---

#### 🟡 T4 — shellExecutor 测试覆盖不足

**位置**：`electron/agent/tools/executors/shellExecutor.test.ts`

**问题**：仅 1 个测试用例，只覆盖"复用 sandbox 评估"正常路径。缺少：拒绝路径、prompt 触发、cwd 重定向、超时强杀。

---

#### 🟡 PR4 — electronApi.d.ts 与 ipcApi.ts 类型不同步

**位置**：`src/electronApi.d.ts:335-337` vs `src/services/ipcApi.ts:59-62`

**问题**：Excel 状态类型不一致：
- `electronApi.d.ts`: `detectStatus() => Promise<{ connected; host }>`（缺 version/workbookName/availableHosts）
- `ipcApi.ts`: `detectStatus() => Promise<{ connected; host; version?; workbookName?; availableHosts? }>`

**建议**：以 `ipcApi.ts` 为准补全 `electronApi.d.ts` 的类型声明。

---

### 项目规范（4 项）

---

#### 🟡 PR5 — CHANGELOG 严重滞后

**问题**：
- 最新条目为 "2026-06-30"，但 git log 显示其后有多个未记录提交
- CHANGELOG 称安装包版本 0.1.2，实际 package.json 已是 0.1.61
- CHANGELOG 称"74 个测试文件、420 个测试"，实际已 105 文件 / 609 测试

**建议**：补 Unreleased 段，更新版本号和测试基线。

---

#### 🟡 S7 — 依赖安全

**问题**：项目无 `npm audit` / dependabot 配置，新增依赖无安全检查。

**建议**：CI 中增加 `npm audit --audit-level=high` 步骤。

---

## 四、💭 P2 问题清单（酌情处理）

| # | 位置 | 问题 |
|---|------|------|
| 1 | `Sidebar.tsx:190-206` | 拖拽 resize 时 mousemove 高频触发 setState，建议用 rAF 节流 |
| 2 | `useComposer.ts:326-331` | textarea onChange 每键触发 setState + DOM 写入，可接受 |
| 3 | `preload.ts:89` | onStreamDelta 回调 data 类型注解缺 clientId，与 electronApi.d.ts 不一致 |
| 4 | `Sidebar.tsx` 多处 | 代码重复的排序逻辑可提取为公共函数 |
| 5 | 全项目 | 魔法数字（如 50000、10000 的 slice 截断）应提取为命名常量 |

---

## 五、✅ 亮点清单（值得学习）

### 架构设计

1. **✅ 沙箱命令执行设计专业** — `security/sandbox/index.ts` 使用 `execFile`（非 `exec`）、Windows `-EncodedCommand` 防注入、ENV_ALLOWLIST 清洗环境变量、进程树强杀。参考 Codex execpolicy 设计，审计日志完整。

2. **✅ 分层架构 + 自动守护测试** — `electron/agent/` 严格分层（interaction → runtime → core），`architecture.test.ts` 自动校验文件存在性和分层依赖。

3. **✅ IPC 抽象层完整** — `desktop/src/` 中 0 处直接 `window.electronAPI` 调用，全部通过 `ipcApi` 抽象层。

4. **✅ IPC schema 单一事实来源** — `ipcSchemas.ts` 集中定义 + `validateInput` 工具 + `z.infer` 推导类型（架构方向正确，只是接入不完整）。

### 性能优化

5. **✅ savePartial 增量持久化** — `settingsStore.ts:411-467`，所有 setX action 只写入变更字段，KEY_MAP + COMPACTION_FIELDS 双层映射，Promise.all 并行写入。

6. **✅ 流式增量缓冲合并** — `chatStore.ts:44-61, 187-206`，`mergeBufferedStreamDeltas` 合并相邻同源 delta，50ms 节流，大幅降低 React 渲染压力。

7. **✅ stats:getSummary 替代 N+1** — 注释明确"单次 IPC 调用，替代 N+1 的 thread:load"。

### 代码质量

8. **✅ ComposerArea composer prop 模式** — 将 useComposer hook 的 25+ 返回值作为单一 `composer` prop 传递，避免 28 props 的 prop drilling。

9. **✅ agentEventHandler 纯函数化** — 从 chatStore 提取为纯函数 `handleAgentEvent(event, current, patches) -> patches`，不依赖 Zustand，便于测试。

10. **✅ chatStore 事件监听器单例管理** — 模块级 unsubscribe 句柄，setupListeners 在 store 初始化时调用一次。

11. **✅ App.tsx 完整的 useEffect 清理** — resize 监听、displayMode 订阅、blur 监听全部有 return cleanup。

### 测试质量

12. **✅ streamCollector.test.ts 边界覆盖优秀** — 覆盖错误重试、JSON 解析失败保留 raw、流中断 emit 已收集内容。

13. **✅ chatStore.test.ts 竞态测试设计优秀** — 覆盖 streaming 中入队、thread id 未就绪拒绝、stale in_progress 元数据不复活。

14. **✅ IPC 通道命名 100% 合规** — 所有 ipcMain.handle 通道均符合 `domain:action` 格式。

---

## 六、修复优先级建议

### 第一周（P0 安全 + 最严重行数超标）

| 优先级 | 任务 | 预估影响 |
|--------|------|----------|
| 1 | 补全 IPC Zod schema 校验（S1 + S2） | 安全防线加固 |
| 2 | 文件操作路径穿越防护（S5） | 安全防线加固 |
| 3 | `agentLoop.ts` 1276 行拆分 | 可维护性 |
| 4 | `ipcHandlers.ts` 1057 行拆分 | 可维护性 |

### 第二周（P1 性能 + 测试）

| 优先级 | 任务 |
|--------|------|
| 5 | Sidebar N+1 查询改批量 |
| 6 | useExcelConnection setTimeout 清理 |
| 7 | ChatPage useEffect Promise 取消 |
| 8 | 沙箱核心模块补测试（parseCommand/execPolicy） |
| 9 | electron/ 13 处 console.warn → logger |

### 第三周（P1 规范 + 剩余行数）

| 优先级 | 任务 |
|--------|------|
| 10 | @MOCK_INTERFACE 标记补全 |
| 11 | electronApi.d.ts 类型同步 |
| 12 | CHANGELOG 更新 |
| 13 | 剩余超标文件拆分（Sidebar/settingsStore 等） |
| 14 | 引入 ESLint + Prettier 工具链 |

---

## 附录 A：超过行数限制的文件完整列表

### TS/TSX 文件（29 个）

| 行数 | 文件 |
|------|------|
| 1276 | `electron/agent/core/agentLoop/agentLoop.ts` |
| 1057 | `electron/main-modules/ipcHandlers.ts` |
| 775 | `src/components/Sidebar.tsx` |
| 749 | `src/store/settingsStore.ts` |
| 745 | `src/services/ipcApi.ts` |
| 674 | `electron/agent/memory/sessionStore.ts` |
| 671 | `electron/agent/memory/stateRuntimeStore.ts` |
| 662 | `electron/agent/tools/implementations/officeOpenXml/advancedExcel.ts` |
| 623 | `src/store/chatStore.ts` |
| 600 | `src/components/task/OCRTaskComposerPanel.tsx` |
| 566 | `electron/agent/core/agentLoop/toolExecutor.ts` |
| 554 | `electron/agent/tools/executors/webSearchExecutors.ts` |
| 544 | `electron/agent/knowledge/documentParser.ts` |
| 537 | `electron/agent/knowledge/sqliteStore.ts` |
| 509 | `electron/agent/tools/implementations/office/officeComActionBridge.ts` |
| 494 | `src/electronApi.d.ts` |
| 481 | `electron/agent/tools/implementations/officeOpenXml/advancedPresentation.ts` |
| 481 | `electron/agent/providers/openaiResponsesClient.ts` |
| 474 | `src/components/settings/GeneralSettings.tsx` |
| 473 | `electron/agent/tools/implementations/excel/excelComBridge.ts` |
| 470 | `electron/agent/shared/types.ts` |
| 440 | `src/components/settings/EditProviderDialog.tsx` |
| 431 | `src/hooks/useComposer.ts` |
| 427 | `electron/agent/tools/implementations/office/presentationComBridge.ts` |
| 426 | `electron/main-modules/settingsManager.ts` |
| 423 | `electron/agent/providers/openaiCompatibleClient.ts` |
| 419 | `src/components/settings/AddProviderDialog.tsx` |
| 416 | `src/components/settings/UsageStats.tsx` |
| 404 | `electron/main-modules/mineruOcr.ts` |

### CSS 文件（4 个）

| 行数 | 文件 |
|------|------|
| 1553 | `src/styles/settings.css` |
| 1101 | `src/styles/sidebar.css` |
| 841 | `src/styles/chat.css` |
| 674 | `src/styles/composer.css` |

---

## 附录 B：未覆盖测试的文件列表

### electron/agent 下无 .test.ts 的源文件（74 个）

<details>
<summary>点击展开完整列表</summary>

```
electron/agent/attachments/imageAttachmentResolver.ts
electron/agent/automation/jscript.ts
electron/agent/automation/json.ts
electron/agent/automation/processLimits.ts
electron/agent/automation/python.ts
electron/agent/core/agentLoop/sessionCompactionConfig.ts
electron/agent/core/agentLoop/summaryGenerator.ts
electron/agent/core/agentLoop/threadLifecycle.ts
electron/agent/core/agentLoop/threadStateManager.ts
electron/agent/core/agentLoop/turnRunner.ts
electron/agent/core/agentLoop/turnState.ts
electron/agent/knowledge/documentParser.ts
electron/agent/knowledge/embeddingService.ts
electron/agent/knowledge/knowledgeIndexer.ts
electron/agent/knowledge/retriever.ts
electron/agent/knowledge/sqliteStore.ts
electron/agent/knowledge/textChunker.ts
electron/agent/knowledge/workbookNotesStore.ts
electron/agent/memory/rolloutSearchContent.ts
electron/agent/memory/stateRuntimeMappers.ts
electron/agent/memory/stateRuntimePaths.ts
electron/agent/memory/stateRuntimeRecovery.ts
electron/agent/memory/stateRuntimeSchema.ts
electron/agent/memory/stateRuntimeTypes.ts
electron/agent/prompts/sections/folderContextPrompt.ts
electron/agent/prompts/sections/formulaAssistantPrompt.ts
electron/agent/prompts/sections/modelPrompt.ts
electron/agent/prompts/sections/officeToolsPrompt.ts
electron/agent/prompts/sections/permissionPrompt.ts
electron/agent/prompts/sections/qualityPrompt.ts
electron/agent/prompts/sections/scenarioPrompt.ts
electron/agent/prompts/sections/scriptPrompt.ts
electron/agent/providers/aiClient.ts
electron/agent/providers/aiClientFactory.ts
electron/agent/providers/aiClientTypes.ts
electron/agent/providers/anthropicClient.ts
electron/agent/providers/modelContextWindows.ts
electron/agent/runtime/agentGlobalSettings.ts
electron/agent/runtime/agentRuntime.ts
electron/agent/security/sandbox/audit.ts
electron/agent/security/sandbox/defaultRules.ts
electron/agent/security/sandbox/execPolicy.ts
electron/agent/security/sandbox/parseCommand.ts
electron/agent/storage/nodeSqlite.ts
electron/agent/tools/contracts/excel.ts
electron/agent/tools/contracts/office.ts
electron/agent/tools/contracts/scriptEnvironment.ts
electron/agent/tools/data/excelFunctionCatalog.ts
electron/agent/tools/executors/createToolExecutors.ts
electron/agent/tools/executors/fileExecutors.ts
electron/agent/tools/executors/localDocumentParser.ts
electron/agent/tools/executors/pythonExecutor.ts
electron/agent/tools/executors/validation.ts
electron/agent/tools/implementations/excel/connectionMetadata.ts
electron/agent/tools/implementations/excel/excelScriptBridgeCom.ts
electron/agent/tools/implementations/excel/excelUiComBridge.ts
electron/agent/tools/implementations/excel/excelVbaComBridge.ts
electron/agent/tools/implementations/excel/sheetOperations.ts
electron/agent/tools/implementations/excel/workbookOperations.ts
electron/agent/tools/implementations/officeOpenXml/officeOpenXmlFileBridge.ts
electron/agent/tools/implementations/officeOpenXml/presentationTemplate.ts
electron/agent/tools/registry/file.ts
electron/agent/tools/registry/formula.ts
electron/agent/tools/registry/knowledge.ts
electron/agent/tools/registry/memory.ts
electron/agent/tools/registry/ocr.ts
electron/agent/tools/registry/office.ts
electron/agent/tools/registry/python.ts
electron/agent/tools/registry/range.ts
electron/agent/tools/registry/script.ts
electron/agent/tools/registry/sheet.ts
electron/agent/tools/registry/shell.ts
electron/agent/tools/registry/toolDefinitions.ts
electron/agent/tools/registry/ui.ts
electron/agent/tools/registry/web.ts
electron/agent/tools/registry/workbook.ts
```

</details>

### src/ 下无 .test.ts 的源文件（45 个）

<details>
<summary>点击展开完整列表</summary>

```
src/App.tsx
src/components/chat/AssistantGroupBlock.tsx
src/components/chat/AttachmentImagePreview.tsx
src/components/chat/CompactionNotice.tsx
src/components/chat/ComposerArea.tsx
src/components/chat/MessageBubble.tsx
src/components/chat/ModelQuickSwitch.tsx
src/components/chat/ReasoningBubble.tsx
src/components/chat/ResumeHint.tsx
src/components/chat/ToolCallBubble.tsx
src/components/chat/ToolConfirmDialog.tsx
src/components/ChatPage.tsx
src/components/common/ErrorBoundary.tsx
src/components/common/FloatingTaskPanel.tsx
src/components/common/IconMap.tsx
src/components/common/LoadingSpinner.tsx
src/components/excel/HostSelectionDialog.tsx
src/components/office/OfficePreviewPanel.tsx
src/components/settings/AddProviderDialog.tsx
src/components/settings/DeleteConfirmDialog.tsx
src/components/settings/EditProviderDialog.tsx
src/components/settings/ExecPolicySettings.tsx
src/components/settings/GeneralSettings.tsx
src/components/settings/KnowledgeSettings.tsx
src/components/settings/ModelConfigList.tsx
src/components/settings/ModelSettings.tsx
src/components/settings/modelSettingsI18n.tsx
src/components/settings/OpenSourceSettings.tsx
src/components/settings/ReasoningModeSelect.tsx
src/components/settings/UsageStats.tsx
src/components/settings/useTestConnection.ts
src/components/SettingsPage.tsx
src/components/sidebar/FileContextMenu.tsx
src/components/sidebar/FolderSection.tsx
src/components/sidebar/SidebarSearchPalette.tsx
src/components/sidebar/ThreadContextMenu.tsx
src/components/Sidebar.tsx
src/components/task/CodeTaskComposerPanel.tsx
src/components/task/FormulaTaskComposerPanel.tsx
src/components/task/ReportTaskComposerPanel.tsx
src/hooks/useExcelConnection.ts
src/hooks/useOfficeConnection.ts
src/i18n.ts
src/main.tsx
src/utils/reasoningSupport.ts
src/utils/textCleaner.ts
```

</details>

---

> **审查完毕。** 建议按第六章的优先级顺序修复，第一周聚焦安全防线加固和最严重的文件拆分。
