# Code Review 修复计划 — 执行记录

> 生成日期：2026-06-24
> 分支：`feature/new-feature`
> 状态：**Week 1–4 全部完成 ✅**
> 维护说明：本文是早期 code review 执行记录。当前项目规范以 `docs/development-standards.md` 为准；当前测试基线为 165 个测试文件、862 个测试。

---

## 执行总览

| 周次 | 任务数 | 完成 | 状态 |
|------|--------|------|------|
| Week 1 | 4 | 4 | ✅ |
| Week 2 | 4 | 4 | ✅ |
| Week 3 | 4 | 4 | ✅ |
| Week 4 | 6 | 6 | ✅ |
| **合计** | **18** | **18** | **✅** |

---

## Week 1 — P0 Bug 修复 + 基础加固

| # | 任务 | 优先级 | 状态 | 关键文件 |
|---|------|--------|------|----------|
| 1 | TurnItem 显示顺序修复 | P0 | ✅ | `agentLoop.ts`, `chatStore.ts`, `ChatPage.tsx` |
| 2 | API 400 错误修复（孤立 tool_call） | P0 | ✅ | `aiClient.ts`（turnItemsToChatMessages / buildRequestMessages） |
| 3 | Electron 类型检查修复 | P0 | ✅ | `agentLoop.ts`, `excelBridge/` |
| 4 | AgentLoopConfig reasoningMode 补全 | P0 | ✅ | `agentLoop.ts` |

## Week 2 — IPC Zod Schema + 结构化日志

| # | 任务 | 优先级 | 状态 | 关键文件 |
|---|------|--------|------|----------|
| 5 | IPC 输入验证 zod schema | P0 | ✅ | `ipcHandlers.ts` |
| 6 | 结构化日志系统 | P1 | ✅ | `logger.ts`（新建） |
| 7 | UsageStats N+1 查询优化 | P1 | ✅ | `ipcHandlers.ts` |
| 8 | excelBridge.ts 模块拆分 | P1 | ✅ | `excelBridge/` 目录（5 子模块） |

## Week 3 — toolRegistry 拆分 + 主流程加固

| # | 任务 | 优先级 | 状态 | 关键文件 |
|---|------|--------|------|----------|
| 9 | toolRegistry.ts 拆分 | P1 | ✅ | `toolRegistry/` 目录（4 子模块） |
| 10 | agentLoop 首轮拆分 | P1 | ✅ | `imageAttachmentResolver.ts`, `toolExecution.ts`, `compactionManager.ts` |
| 11 | main.ts 拆分 | P1 | ✅ | `main-modules/` 目录（4 子模块） |
| 12 | chatStore 拆分 | P1 | ✅ | `agentEventHandler.ts`, `threadActions.ts` |

## Week 4 — 深度重构 + 测试基础设施

| # | 任务 | 优先级 | 状态 | 关键文件 |
|---|------|--------|------|----------|
| #18 | agentLoop 深度拆分（1,066→6 子模块） | P0 | ✅ | `agentLoop/` 目录（6 文件） |
| #19 | ComposerArea Props 精简（28→7） | P1 | ✅ | `ComposerArea.tsx`, `ChatPage.tsx` |
| #20 | settingsStore 增量持久化 | P1 | ✅ | `settingsStore.ts` |
| #21 | vitest 单元测试基础设施 | P0 | ✅ | `vitest.config.ts`，早期 2 个测试文件；当前已扩展到 165 个测试文件、862 个测试 |
| #22 | OCR Mock 标记 | P2 | ✅ | `OCRTaskComposerPanel.tsx` |
| #23 | IPC 依赖注入抽象层 | P1 | ✅ | `services/ipcApi.ts`（新建）, 5 个核心模块迁移 |

---

## Week 4 改动详情

### #18 agentLoop 深度拆分

**原始**：`electron/agent/agentLoop.ts`（1,066 行）
**目标**：按运行时职责域拆分，对外 API 不变

| 子模块 | 文件 | 行数 | 职责 |
|---|---|---|---|
| streamCollector | `agentLoop/streamCollector.ts` | ~140 | 流式事件收集 |
| toolExecutor | `agentLoop/toolExecutor.ts` | ~260 | 工具执行 + 审批 |
| buildStreamParams | `agentLoop/buildStreamParams.ts` | ~90 | 推理降级链 + 系统提示词 |
| agentLoop（编排器） | `agentLoop/agentLoop.ts` | ~450 | AgentLoop 类 |

> 注：后续架构清理已移除无职责的兼容出口和影子层。当前拆分要求是按职责直接引用具体模块，不再默认要求 `index.ts` barrel。

### #19 ComposerArea Props 精简

**原始**：28 个独立 props
**目标**：useComposer hook 返回值整体作为 `composer` prop

```
// Before
<ComposerArea
  inputText={...} setInputText={...} handleSend={...}
  showAttachPopover={...} setShowAttachPopover={...}
  ... (28 props)
/>

// After
<ComposerArea
  composer={composer}           // ReturnType<typeof useComposer>
  currentFolder={...}
  currentFolderFiles={...}
  showWelcomeComposer={...}
  onSend={...}
  onInterrupt={...}
  onOpenSettings={...}
/>
```

### #20 settingsStore 增量持久化

**原始**：每次变更调用 `saveSettings()` 写入全部 9 个 key
**目标**：`savePartial([changedKey], get)` 仅写变更 key

```
// Before
setPermissionMode: (mode) => {
  set({ permissionMode: mode });
  saveSettings(); // 写 9 个 key
}

// After
setPermissionMode: (mode) => {
  set({ permissionMode: mode });
  savePartial(["permissionMode"], get); // 写 1 个 key
}
```

### #21 vitest 测试基础设施

- `vitest.config.ts`：node 环境、v8 覆盖率、globals 启用
- 早期基线：`compaction.test.ts` 24 tests、`agentLoop.test.ts` 11 tests。
- 当前基线：165 个测试文件、862 个测试，覆盖 agent loop、工具执行、Office action、StateRuntime、长期记忆、沙箱、知识库与桌面更新等模块。
- 运行命令：`npm test` / `npm run test:watch` / `npm run test:coverage`

### #22 OCR Mock 标记

在 `OCRTaskComposerPanel.tsx` 的 mock 代码处添加 `@MOCK_INTERFACE` 标记，便于后续接入真实 OCR 时快速定位。

### #23 IPC 依赖注入抽象层

**新建** `src/services/ipcApi.ts`（~350 行）：
- `IIpcApi` 接口：类型安全的 IPC 调用契约
- `ipcApi` 实例：运行时通过 `getRaw()` 获取 `window.electronAPI`，不可用时安全降级
- `createMockIpcApi()`：测试辅助，一键注入 mock

**已迁移的核心模块**（`window.electronAPI` 直接调用归零）：
- `store/chatStore.ts` — agent + tool 相关 IPC
- `store/threadActions.ts` — thread 管理 IPC
- `store/settingsStore.ts` — settings 读写 IPC
- `hooks/useComposer.ts` — dialog + folder IPC
- `hooks/useExcelConnection.ts` — excel 状态 IPC

---

## 验证结果

| 检查项 | 结果 |
|--------|------|
| TypeScript `tsc --noEmit` | ✅ |
| Vite `vite build` | ✅ |
| vitest `npm test`（当前 420 tests） | ✅ |

---

## 待办（后续周次）

- 本文待办为历史记录。当前后续任务以 `docs/dev-log.md`、`docs/development-standards.md` 和各阶段计划文档为准。
