# 用户偏好记忆

## 沟通风格
- **用户不是专业计算机人员**，不知道太多专业术语
- 必须以**产品经理的角度**理解用户的自然语言需求
- 回答时避免堆砌技术术语，用通俗语言解释
- 需求理解要站在"用户想实现什么效果"的角度，而不是"用户说了什么技术词"
- 当用户描述模糊时，主动用场景化的方式确认理解是否正确

## 项目背景
- 当前项目是一个**Excel AI 助手**，半成品状态
- 核心功能：解答 Excel 相关问题
- 目标：参考 OpenAI Codex 源码，开发一个**桌面端应用**
- Codex 本身是 CLI 版本，需要将其思路转化为桌面版

## 插件体验问题（桌面版必须解决）

### 1. 上下文衔接问题（最严重）
- 生成公式、代码解疑等核心功能，在**中断**或**达到最大输出 token 被迫中断**时，用户追问或让继续都难以衔接中断前的进度
- 根因：Agent 循环的 compaction/summary 逻辑在中断时没有保存足够的"断点状态"，导致继续对话时丢失关键上下文
- 桌面版解决思路：
  - 每次工具调用和中间推理都持久化"断点快照"（turn_id + 已完成的工具结果 + 未完成的 pending_calls + 当前目标）
  - 中断时自动生成"衔接摘要"（类似 Codex 的 goal checkpoint，但要更细粒度）
  - 用户追问"继续"时，自动注入断点摘要，让 AI 知道做到哪了、下一步是什么
  - 达到 max_tokens 时主动触发断点保存，而不是等到彻底结束

### 2. OCR/发票识别模块未打通
- 测试发票识别时仍报错
- 桌面版解决思路：第一版先把 OCR 功能做好端到端测试，确保整条链路通畅

### 3. 国内模型适配 + 推理过程展示
- 需要适配国内主流厂商：讯飞星辰、百度千帆、火山引擎、腾讯云、京东云、阿里云（百炼）等
- **关键需求：推理过程（thinking/reasoning）必须完整展示给用户**，不能像 Codex 那样隐藏或只给摘要
  - 原因：用户需要根据推理信息判断模型理解需求有没有偏，便于及时纠正
  - 很多国内模型（如 DeepSeek-R1、讯飞、Kimi 等）开启思考模式后，推理文本是流式返回的
  - 当前 add-in 有 reasoning 展示（AgentActivityRow 中 Brain 图标），但只是作为活动行展示，不够显眼
- 桌面版解决思路：
  - 推理过程用独立的"思考气泡"展示，流式逐字显示，跟 AI 回复一样有实时感
  - 可以折叠但默认展开，让用户能实时看到 AI 在想什么
  - 每个 AI 客户端适配器都要正确处理各家 reasoning/thinking 字段的差异
  - 国内模型大多兼容 OpenAI 接口格式，但 reasoning 字段位置不同（有的在 `reasoning_content`，有的在 `thinking`，有的在 content 里的 `<think>` 标签）

## 文档维护规范（强制执行）

**每次 Bug 修复或功能新增后，必须维护项目文档**，规则如下：

### 1. 文档位置
- **统一开发日志**：`docs/dev-log.md`，所有日志追加到同一文件，按 `## YYYY-MM-DD — 简要描述` 分区
- 新日志追加到文档末尾，**不删除、不覆盖**已有日志
- 避免碎片化，不再为每次改动单独创建文件

### 2. 文档必须包含的内容
- **日期和分支**：标注修复/新增发生的日期和所在 Git 分支
- **问题描述**（Bug 修复）或 **需求描述**（功能新增）：用通俗语言说明
- **根因分析**（Bug 修复）或 **设计方案**（功能新增）：技术层面的分析
- **改动详情**：涉及哪些文件、哪些函数、具体改了什么
- **验证结果**：编译、构建、测试是否通过
- **影响范围**：对现有功能的影响、行为变化

### 3. CHANGELOG 同步更新
- 每次改动都必须在 `CHANGELOG.md` 的 `Unreleased` 下添加条目
- Bug 修复归入 `### Fixed`
- 功能新增归入 `### Added`
- 行为变更归入 `### Changed`

## 技术决策记录
- Codex 参考源码已克隆到：D:\codex-reference（也在项目内 codex-reference/ 目录）
- 桌面端技术方案：Electron + React + TypeScript
- 目标平台：仅 Windows
- Excel 交互方式：通过 Windows COM 直接操控正在运行的 Excel/WPS
- AI 模型支持：OpenAI、DeepSeek、Claude、自定义兼容接口 + 国内厂商（讯飞/百度/火山/腾讯/京东/阿里）
- AI 请求方式：本地直接调用 AI API，不经服务器；服务器仅负责更新和通知

## 项目架构变更记录

### Week 4（2026-06-25）代码审查修复
- **agentLoop 深度拆分**：`agentLoop.ts`（1,066 行）→ `agentLoop/` 目录（6 子模块，各 ≤ 450 行），对外 API 不变
- **IPC 依赖注入**：新建 `src/services/ipcApi.ts` 抽象层，核心模块（chatStore、threadActions、settingsStore、useComposer、useExcelConnection）已迁移，`window.electronAPI` 直接调用归零
- **vitest 测试基础设施**：`vitest.config.ts` + 2 个测试文件（35 tests），运行命令 `npm test`
- **ComposerArea Props 精简**：28 → 7 props，useComposer hook 返回值整体传入
- **settingsStore 增量持久化**：每次变更仅写变更 key（savePartial），不再全量写入 9 个 key

### Week 4 补充（2026-06-25）TypeScript 编译错误批量修复
- 修复 7 项编译错误，前端 + Electron 端均实现零错误编译
- 关键修复：TurnItem 类型窄化、Zod v4 z.record() 双参数、.ts→.tsx 重命名、ElectronStore 泛型、vitest beforeEach 导入

### 文件结构（Week 4 后）
```
desktop/electron/agent/
  agentLoop/           ← 新目录（Week 4）
    agentLoop.ts       ← AgentLoop 类编排器
    streamCollector.ts ← 流式事件收集
    toolExecutor.ts    ← 工具执行 + 审批
    compactionManager.ts ← 上下文压缩
    buildStreamParams.ts ← 推理降级 + 系统提示词
    index.ts           ← barrel re-export
  toolRegistry/        ← 工具定义 + 执行器
  excelBridge/         ← Excel COM 桥接

desktop/src/
  services/
    ipcApi.ts          ← IPC 抽象层（Week 4 新建）
  store/
    chatStore.ts       ← 聊天状态
    threadActions.ts   ← 会话管理
    settingsStore.ts   ← 设置状态
    agentEventHandler.ts ← 事件处理
  hooks/
    useComposer.ts     ← 输入框逻辑
    useExcelConnection.ts ← Excel 连接
  components/chat/
    ComposerArea.tsx   ← 输入框组件

desktop/electron/agent/agentLoop.test.ts ← vitest 测试
desktop/electron/agent/compaction.test.ts ← vitest 测试
```
