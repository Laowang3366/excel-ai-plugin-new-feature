# Agent 能力提升实施方案（参考 Claude-Code / OpenAI Codex）

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 借鉴 Claude-Code（`github.com/pengchengneo/Claude-Code`）与 OpenAI Codex（`github.com/openai/codex`）两个成熟 agent 的设计，补齐当前 Excel AI 插件在**上下文管理、提示词缓存、工具结果治理、技能体系、记忆、长任务恢复**六个维度的能力短板，降低 token 成本、提升长任务稳定性，并支持跨会话知识沉淀。

**Architecture:** 以最小侵入方式增强现有 `AgentLoopService` 与公式编排链路，不引入新框架。所有改动遵循现有分层：纯函数模块 + 依赖注入 + Python 回归守卫测试。新增模块放在 `backend/app/services/agent_context/`（上下文治理）、`backend/app/services/agent_tools/`（工具注册与结果治理）、`backend/app/services/memory/`（记忆）和 `backend/app/services/skills/`（技能发现）下；前端只消费后端返回的结构化工具事件与技能清单。

**Document Purpose:** 本文是参考 Agent 项目的能力抽取与落地路线。每个任务都必须能追溯到 Claude-Code 或 Codex 的一个运行时机制，例如 prompt section、tool registry、compaction、memory pipeline、skill discovery 或 deferred tool loading，并把这些机制转译成当前 Excel/WPS Agent 可复用的能力。

**Non-goals:** 本文不排期单点产品缺陷，不复述既有 UI/交互/业务问题，也不替代代码审查报告。已发现缺陷只能在后续审查中作为回归输入；进入本文时必须已经被抽象成跨模块能力，例如统一 trace 状态机、工具风险契约或上下文隔离机制。

**Tech Stack:** Python 3.11 + FastAPI + Redis + openai SDK（后端），TypeScript + React + Office.js（插件），pytest 回归守卫。

---

## 现状基线与差距

| 维度 | 当前实现 | Claude-Code / Codex 做法 | 差距 |
|---|---|---|---|
| **缓存** | 零缓存，transcript 平铺，workbook 摘要夹在中间破坏前缀 | 系统提示词切「静态可缓存段 / 动态段」，显式缓存边界标记；备忘 section（`systemPromptSections.ts`） | **无任何 prompt cache，token 全量重算** |
| **提示词** | 4 段超长公式指令高度重复，JSON schema 样例内嵌 | apply_patch 极简 grammar + 3 行示例；continuation prompt 的「完成审计」写法 | **重复、缓存不友好、易漂移** |
| **工具结果** | `range.read`/`workbook.inspect` 全量返回，零截断 | `tool_output.rs` 的 `telemetry_preview`（2KB/64 行上限）；`contains_external_context` 标志 | **大工作簿直接爆上下文** |
| **上下文管理** | transcript 只增不减，`max_turn_count=8` 硬截断 | `compact.rs` 双时机压缩（pre-turn / mid-turn）+ 极简摘要 prompt + pre/post hooks；已消费结果摘要化 | **无压缩，长任务硬截断报错** |
| **记忆** | 完全没有 | memdir 两段式召回（廉价模型先筛 ≤5 条）；AGENTS.md 分层发现注入；Codex 两阶段 memory pipeline + 外部内容隔离 | **跨会话知识无法沉淀** |
| **技能** | task prompt 前端硬编码下发 | skill = 带 frontmatter 的 markdown 目录，后端扫描可发现 | **不可扩展、不可持久化** |
| **工具注册** | `agent_tool_schema.py` 只按名字分组，风险信息较弱 | 工具定义携带 schema、权限、风险、日志摘要、截断策略、是否外部上下文 | **无法统一审批、日志、截断、越权拦截** |

**关键洞察（来自参考源码）：**

1. Claude-Code `systemPromptSections.ts:20-58`：`systemPromptSection()` 默认缓存，`DANGEROUS_uncachedSystemPromptSection()` 必须给理由才允许破坏缓存 —— 把「是否破坏缓存」变成了一等公民的 API。
2. Claude-Code `prompts.ts:526` 注释：「曾因每轮重算 token budget，每翻一次 bust 掉 ~20K token 缓存」。
3. Codex `compact.rs:54-67`：pre-turn 用 `DoNotInject`，mid-turn 用 `BeforeLastUserMessage`，因为模型被训练成「压缩摘要是 history 最后一项」，且不能破坏前缀缓存。
4. Codex `tool_output.rs:23-25, 116-119`：工具结果带 `contains_external_context` 标志，含外部内容时**禁用本条 memory 提取**，防污染。
5. Codex `prompts/templates/compact/prompt.md`：压缩 prompt 只有 6 行（「为另一个 LLM 写交接摘要」）。
6. Claude-Code `findRelevantMemories.ts:88-95`：用户正在用某工具时，不召回该工具的 API 文档类记忆（噪音），但**仍召回**「坑/警告」类记忆。
7. Claude-Code 和 Codex 都把“任务能力”做成可发现、可组合的结构，而不是把所有提示词写死在单个前端文件里；当前项目适合把不同业务模块迁为后端 skill/task fragment。
8. 工具调用日志要和执行器统一：模型看到的 tool result、前端看到的 activity、后端审计日志不能各写一套字段；同一条工具事件应有统一 `call_id`、`risk_level`、`target`、`summary`、`result_status`。

---

## 能力抽取地图

| 参考机制 | 当前项目补强能力 | 对应任务 | 验收关注点 |
|---|---|---|---|
| Claude-Code prompt sections | 稳定 system/task policy 与动态 workbook context 分离 | Task 0A, Task 1, Task 1A | 缓存前缀稳定、隐藏策略不进入用户气泡、追问不重复注入 |
| Codex tool output preview | 大工具结果只进入 preview，完整结果外置可续读 | Task 2, Task 8 | `truncated/success/resume_hint/contains_external_context` 成为统一 tool result contract |
| Codex compaction hooks | pre-turn/mid-turn 压缩与交接摘要 | Task 5 | 长会话压缩后仍保留最近用户目标、工具执行摘要和下一步 |
| Codex AGENTS.md + memory pipeline | 项目级规则、工作簿笔记、外部内容隔离 | Task 6 | 可召回可信工作簿知识，外部 OCR/PDF/网页内容不污染记忆 |
| Claude-Code skill loader | 任务能力以 markdown/frontmatter 发现 | Task 7 | 新增业务模块不需要改前端硬编码 prompt |
| Codex deferred tool loading | ToolSearch 按需暴露 schema | Task 10 | 工具规模增长时模型只看到当前任务必要工具 |
| Agent runtime approval model | 风险、权限、审批、执行状态解耦 | Task 0B, Task 0C, Task 0D, Task 0E | 模型意图、运行时决策、前端执行、审计日志使用同一事件语义 |

---

## 落地优先级与里程碑

| 阶段 | 任务 | 优先级 | 预期收益 |
|---|---|---|---|
| **M0：Agent 运行时骨架** | Task 0A（任务策略片段）、Task 0B（工具注册表）、Task 0C（能力沙箱模型）、Task 0D（审批事件契约）、Task 0E（结构化 trace） | **P0** | 把提示词、工具、审批、日志从散落实现收敛为可组合运行时能力 |
| **M1：缓存与裁剪** | Task 1（提示词缓存断点）、Task 2（工具结果截断）、Task 8（工具结果 success/truncated 标志） | **P0-P1** | token 成本 ↓ 40-70%，长工作簿不再爆上下文 |
| **M2：提示词与上下文治理** | Task 1A（Context Fragment Envelope）、Task 3（公式提示词去重）、Task 4（已消费结果摘要化） | **P1** | 稳定性 ↑，缓存命中率 ↑，追问不重复注入 |
| **M3：压缩、记忆与继续审计** | Task 5（会话压缩）、Task 6（AGENTS.md/工作簿笔记）、Task 9（Goal Checkpoint） | **P1-P2** | 跨会话知识沉淀，超长任务可恢复，继续执行有明确下一步 |
| **M4：扩展能力** | Task 7（技能化 task）、Task 10（ToolSearch 式延迟工具发现） | **P2** | 新任务和新工具可按需发现，不再依赖前端硬编码 |

> **建议执行顺序：** Task 0A → Task 0B → Task 0C → Task 0D → Task 0E → Task 2 → Task 8 → Task 1 → Task 1A → Task 3 → Task 4 → Task 6 → Task 5 → Task 9 → Task 7 → Task 10。M0 先做是为了先建立统一运行时骨架，后续缓存、压缩、记忆和技能化才有稳定落点。

---

## Task 0A: 后端任务策略片段（P0）

**目标：** 学习 Claude-Code 的 system prompt section 和 Codex 的 prompt template 分层方式，把不同业务任务拆成后端可组合的 task policy fragment。前端只传任务类型、用户输入和表单参数，后端负责拼装稳定提示词片段、动态上下文片段和工具能力片段。

**Files:**
- Modify: `backend/app/schemas/agent.py`
- Create: `backend/app/services/agent_policy/task_policy.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Modify: `backend/app/services/agent_tool_schema.py`
- Modify: `add-in/src/features/agent/agentTaskPolicy.ts`
- Modify: `add-in/src/features/agent/useChatStreamSender.ts`
- Create: `backend/tests/test_agent_task_policy.py`
- Modify: `backend/tests/test_agent_loop_service.py`

**后端 contract：**

```python
class AgentTaskKind(str, Enum):
    general = "general"
    formula_generation = "formula_generation"
    data_cleaning = "data_cleaning"
    ocr_fill = "ocr_fill"
    report_generation = "report_generation"
    chart_build = "chart_build"
```

`AgentStartRequest` 新增：
- `task_kind: AgentTaskKind = AgentTaskKind.general`
- `task_payload: dict[str, Any] | None = None`
- `conversation_id: str | None = None`
- 保留 `task_system_prompt` 一段兼容期，但后端默认忽略前端值，并记录 `deprecated_task_system_prompt_ignored=true`。

`AgentContinueRequest` 新增：
- `conversation_id: str`
- 不接受 `task_kind` 和 `task_system_prompt`；追问沿用 start 时落库的 task policy，不重复注入模块提示词。

**策略片段规则：**
- `stable_system`：Agent 基础行为、工具协议、输出约束，适合缓存。
- `task_policy`：单一业务任务的最小规则，例如公式、清洗、OCR、报告、图表；同一会话 start 注入一次，continue 不重复注入。
- `dynamic_context`：工作簿快照、选区、用户表单、工具摘要，按 turn 更新。
- `tool_policy`：该 task 可见的工具命名空间和风险等级，只由后端 registry 派生。
- `user_visible`：用户输入与表单摘要，禁止混入隐藏策略。

- [x] **Step 1: 写失败测试**

`test_task_policy_fragments_are_backend_generated`：
1. 构造 `task_kind=formula_generation` 和结构化 `task_payload`。
2. 调用 `build_task_policy`。
3. 断言返回结果包含 `stable_system/task_policy/dynamic_context/tool_policy`，且前端传入的 `task_system_prompt` 不参与拼装。

`test_continue_does_not_reinject_task_prompt`：
1. start 公式任务。
2. continue 追问。
3. 断言 transcript 中 `task_policy` fragment 只出现一次。

- [x] **Step 2: 实现 `task_policy.py`**

提供：
- `build_task_policy(task_kind, task_payload, access_mode) -> AgentTaskPolicy`
- `AgentTaskPolicy.system_fragments`
- `AgentTaskPolicy.allowed_tools`
- `AgentTaskPolicy.denied_tools`
- `AgentTaskPolicy.audit_tags`

- [x] **Step 3: 接入 `AgentLoopService`**

`_start_messages` 只读取后端生成的 `policy.system_fragments`；用户消息只展示用户自然语言和模块表单摘要，不拼接隐藏提示词。

- [x] **Step 4: 前端只传 `taskKind/taskPayload/displayText`**

`agentTaskPolicy.ts` 改为构造用户可见 `displayText` 和结构化 `taskPayload`，删除向后端传 `taskSystemPrompt` 的路径。

- [x] **Step 5: 回归**

运行：
- `pytest backend/tests/test_agent_task_policy.py backend/tests/test_agent_loop_service.py -q`
- `npm --prefix add-in run build`

---

## Task 0B: 工具注册表与风险元数据（P0）

**目标：** 把工具 schema、权限、风险、日志摘要、截断策略集中到后端注册表，避免前端和后端各维护一套工具定义。

**Files:**
- Create: `backend/app/services/agent_tools/registry.py`
- Modify: `backend/app/services/agent_tool_schema.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Modify: `add-in/src/features/agent/agentWorkbookToolProtocol.ts`
- Modify: `add-in/src/features/agent/agentActivityModel.ts`
- Create: `backend/tests/test_agent_tool_registry.py`

**Tool metadata 字段：**

```python
ToolRisk = Literal["read", "write", "destructive", "macro", "external_file"]

@dataclass(frozen=True)
class AgentToolDefinition:
    name: str
    namespace: str
    input_schema: dict[str, Any]
    risk: ToolRisk
    mutating: bool
    requires_full_access: bool
    approval_required: bool
    max_result_chars: int
    summary_fields: tuple[str, ...]
    redacted_fields: tuple[str, ...]
    contains_external_context: bool
    concurrency_safe: bool
```

**注册要求：**
- `workbook.inspect`、`range.read` 是 read risk，可默认执行。
- `range.write`、清洗填入、OCR 填入是 write risk，受 task policy 和权限模式控制。
- `rows.delete`、`columns.delete`、`worksheet.delete/move/rename` 是 destructive risk，必须完整权限 + task allow。
- 宏执行类工具是 macro risk，必须经过 task policy 和 access mode 双重授权。

- [x] **Step 1: 写失败测试**

`test_registry_exposes_all_agent_tools_with_risk_metadata`：断言现有工具都注册且 `risk/max_result_chars/summary_fields/redacted_fields` 不为空。

- [x] **Step 2: 把 `agent_tool_schema.py` 改为从 registry 派生**

`build_agent_tool_definitions(policy)` 不再接收前端 `tool_capabilities` 原样透传，而是：
1. 读取 registry。
2. 应用 task policy allow/deny。
3. 应用权限模式。
4. 输出模型可见 tool schema。

- [x] **Step 3: 前端 activity 使用后端 tool metadata**

后端 start response 带 `tool_manifest_version` 和工具展示 metadata；前端展示以 `call_id/tool_name/risk/status` 为准。

- [x] **Step 4: 回归**

运行：
- `pytest backend/tests/test_agent_tool_registry.py backend/tests/test_agent_schema_tooling.py -q`

---

## Task 0C: 能力沙箱模型（P0）

**目标：** 学习 Codex 对工具能力、审批、执行状态的分层处理，把“模型想调用什么”和“运行时允许执行什么”分开。能力沙箱只描述通用决策模型，便于所有任务类型统一执行 `allow/needs_approval/deny`。

**Files:**
- Create: `backend/app/services/agent_policy/capability_sandbox.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Modify: `backend/app/schemas/agent.py`
- Modify: `add-in/src/features/agent/agentWorkbookToolPolicy.ts`
- Modify: `add-in/src/features/agent/agentVbaToolExecutor.ts`
- Create: `backend/tests/test_agent_capability_sandbox.py`
- Create: `add-in/src/features/agent/__tests__/agentToolPolicy.test.ts`

**沙箱决策：**

```python
class ToolDecision(str, Enum):
    allow = "allow"
    needs_approval = "needs_approval"
    deny = "deny"

@dataclass(frozen=True)
class ToolPolicyDecision:
    decision: ToolDecision
    reason_code: str
    user_message: str
    audit_tags: tuple[str, ...]
```

**决策输入：**
- `task_policy.tool_policy`
- `tool_registry.risk`
- `access_mode`
- `approval_state`
- `conversation_state`

**决策输出：**
- `allow`：可立即执行。
- `needs_approval`：需要用户确认，进入审批事件流。
- `deny`：运行时拒绝，并把拒绝原因作为 tool result 返回给模型。

- [x] **Step 1: 写失败测试**

`test_tool_decision_uses_registry_policy_before_execution`：模拟模型返回一个高风险 tool call，后端先基于 registry 和 task policy 产出 `needs_approval` 或 `deny`，不会直接进入前端执行器。

- [x] **Step 2: 实现 backend sandbox**

在 tool call 进入待执行队列前调用 `decide_tool_execution(policy, tool_call)`。

- [x] **Step 3: 前端执行器二次校验**

`agentVbaToolExecutor.ts` 检查 `taskKind` 和 `toolPolicyDecision`，被禁止时返回 `blocked`，不触发 Office/WPS API。

- [x] **Step 4: 回归**

运行：
- `pytest backend/tests/test_agent_capability_sandbox.py -q`
- `npm --prefix add-in run test -- agentToolPolicy`

执行记录：当前 `add-in` 没有 `test` 脚本，已用 `backend/tests/test_agent_frontend_tool_loop.py` 的前端守卫测试和 `npm --prefix add-in run build` 覆盖二次校验路径。

---

## Task 0D: Human Approval 事件契约（P0）

**目标：** 借鉴 Agent runtime 对人工确认事件的结构化处理，把高风险工具调用转换成稳定的 approval event。审批事件由后端工具元数据和策略规则生成，前端只负责呈现与响应。

**Files:**
- Modify: `backend/app/schemas/agent.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Modify: `backend/app/services/agent_tools/registry.py`
- Modify: `add-in/src/features/agent/components/AgentActivityRow.tsx`
- Modify: `add-in/src/features/agent/agentActivityModel.ts`
- Create: `backend/tests/test_agent_approval_packet.py`

**审批包字段：**

```python
class AgentApprovalPacket(BaseModel):
    call_id: str
    tool_name: str
    risk: str
    target: str
    operation_summary: str
    argument_summary: dict[str, str | int | bool]
    expected_change_summary: str
    rollback_hint: str | None = None
    policy_reason: str
```

**事件规则：**
- `target` 示例：`Sheet1!A1:D20`、`worksheet: 结果表`、`VBA module: Module1`。
- `argument_summary` 只能显示摘要，不显示完整单元格内容、完整 VBA、完整 API Key。
- 审批事件由后端 registry 和 policy 生成，不由模型自由编写。

- [x] **Step 1: 写失败测试**

`test_approval_event_is_generated_from_tool_metadata`：构造 destructive 或 macro tool call，断言 approval event 必含 `target/operation_summary/expected_change_summary/policy_reason`，且这些字段来自 registry summary rule。

- [x] **Step 2: 后端生成审批包**

由 registry 的 `summary_fields/redacted_fields` 生成，不让模型自己写审批说明。

- [x] **Step 3: 前端展示审批包**

展开卡片显示目标、操作、风险、预期影响；前端只渲染确认/拒绝动作，不自行生成审批文案。

- [x] **Step 4: 回归**

运行：
- `pytest backend/tests/test_agent_approval_packet.py -q`
- `npm --prefix add-in run build`

---

## Task 0E: 结构化事件与安全日志（P0）

**目标：** 统一模型响应、工具调用、审批、执行结果和错误日志，把 Agent loop 从“黑盒请求”升级为可观测运行时，同时确保不写入敏感 payload。

**Files:**
- Create: `backend/app/services/agent_trace_service.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Modify: `backend/app/api/ai_agent.py`
- Modify: `backend/tests/test_agent_api.py`
- Create: `backend/tests/test_agent_trace_logging.py`

**事件字段：**
- `event_name`
- `conversation_id`
- `turn_id`
- `call_id`
- `task_kind`
- `tool_name`
- `target`
- `risk`
- `status`
- `duration_ms`
- `token_usage`
- `target_count`
- `actual_count`
- `deviation`
- `error_code`

**安全规则：**
- 不记录 API Key、完整单元格内容、完整公式、完整 VBA、OCR 原文、PDF 原文。
- 对范围和文件可记录 `sheet/range/file_ext/size/hash_prefix`。
- 对 tool result 只记录 `status/counts/truncated/preview_hash`。

- [x] **Step 1: 写失败测试**

`test_agent_trace_logs_outcome_without_sensitive_payload`：提交含 API Key 样式、VBA 代码、公式文本和单元格值的模拟结果，断言日志中没有原文，只保留摘要字段。

- [x] **Step 2: 在 loop 每个状态点写 trace**

覆盖 start、model_request、tool_call_created、approval_waiting、tool_result_received、model_final、turn_failed。

- [x] **Step 3: 定义 trace 状态机**

每个 `call_id` 必须遵循 `created -> waiting_approval|executing|blocked -> completed|failed|cancelled` 的有限状态流。前端 activity、后端日志和模型 tool result 共用同一状态语义。

- [x] **Step 4: 回归**

运行：
- `pytest backend/tests/test_agent_trace_logging.py backend/tests/test_agent_api.py -q`

---

## Task 1: 提示词缓存断点（P0）

**目标：** 让 AGENT_SYSTEM_PROMPT、task prompt 命中厂商前缀缓存；workbook 摘要从 system 中段移到末尾，避免破坏缓存。

**Files:**
- Modify: `backend/app/services/agent_loop_service.py`（`_start_messages:256-275`）
- Modify: `backend/app/ai_client/openai_request_options.py`（增加 `cache_control` 支持）
- Create: `backend/app/services/agent_context/prompt_cache_breakpoint.py`
- Create: `backend/tests/test_agent_prompt_cache_breakpoint.py`

**设计要点：**
- 消息顺序固定为：`[system: AGENT_SYSTEM_PROMPT]` → `[system: task_prompt]` → `...history...` → `[system: workbook_summary]`（末尾）→ `[user: message]`。
- `AGENT_SYSTEM_PROMPT` 与 task prompt 是稳定前缀，命中缓存；workbook 摘要含 Preview 行会变，放最后。
- 通过 `extra_body` / `cache_control` 字段为支持的厂商（DeepSeek/阿里/Kimi/Anthropic）标注缓存断点；不支持时为 no-op。

- [x] **Step 1: 写失败测试**

测试断言：`_start_messages` 返回的 workbook 摘要 system 消息位于 `history` 之后、`user` 消息之前；`build_openai_request_kwargs` 在启用缓存时输出 `cache_control` 相关字段。

- [x] **Step 2: 实现 `prompt_cache_breakpoint.py`**

提供 `apply_cache_breakpoints(messages, provider_id) -> messages` 纯函数：按 provider 能力在 system 段尾部插入 `cache_control: {"type": "ephemeral"}`（Anthropic 风格）或保留前缀不动（DeepSeek/阿里自动前缀匹配）。

- [x] **Step 3: 改 `_start_messages` 顺序**

当前实现已满足顺序要求：`workbook_summary` 位于 history 之后、user message 之前；用回归测试固定，避免后续破坏前缀缓存。

- [x] **Step 4: 在 `chat_with_tools` 调用前 apply cache breakpoints**

- [x] **Step 5: 全量回归 + 联调一次支持缓存的 agent 会话，确认 usage 里出现 `cached_tokens` / `prompt_cache_hit_tokens`**

自动化记录：
- `pytest backend/tests/test_agent_openai_tool_calls.py backend/tests/test_agent_prompt_cache_breakpoint.py -q`：13 passed。
- Live smoke（OpenAI-compatible `gpt-5.5`，两步 `chat_with_tools`，强制 `workbook_inspect` + `range_read`）：第 2 次请求 `cached_tokens=2944`，tool call counts `[1, 1]`。

---

## Task 1A: Context Fragment Envelope（P1）

**目标：** 把 system、developer、module policy、workbook snapshot、memory、tool summary、user message 拆成可标记、可去重、可缓存的 context fragment。该 envelope 提供隐藏策略、用户可见内容、动态上下文和工具摘要之间的隔离能力。

**Files:**
- Create: `backend/app/services/agent_policy/prompt_fragments.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Modify: `backend/app/services/agent_context/prompt_cache_breakpoint.py`
- Create: `backend/tests/test_agent_prompt_fragments.py`

**Fragment 类型：**

```python
class FragmentKind(str, Enum):
    stable_system = "stable_system"
    task_policy = "task_policy"
    workbook_snapshot = "workbook_snapshot"
    conversation_summary = "conversation_summary"
    memory = "memory"
    tool_summary = "tool_summary"
    user_visible = "user_visible"
```

**规则：**
- `stable_system` 与 `task_policy` 可缓存，且同一 conversation 只注入一次。
- `workbook_snapshot` 每个新任务开始前必须注入；追问时若未涉及表格新状态，可用摘要或按需刷新。
- `user_visible` 只来自用户输入和模块表单摘要，不包含隐藏提示词。
- `memory/tool_summary` 带 `contains_external_context` 标志；外部内容不能进入长期记忆。

- [x] **Step 1: 写失败测试**

`test_user_message_does_not_contain_hidden_policy`：业务任务 start 后，断言用户气泡文本只包含用户输入和表单摘要，不包含后端 task policy 规则。

`test_task_policy_fragment_deduplicated_on_continue`：追问时 `task_policy` fragment 不重复。

- [x] **Step 2: 实现 fragment builder**

`build_start_fragments(policy, workbook_summary, history, user_message)` 返回有序 fragments，再由 `render_fragments_to_messages` 转成 provider messages。

- [x] **Step 3: 接入缓存断点**

`prompt_cache_breakpoint.py` 不再按消息位置猜测缓存段，而是按 `FragmentKind` 判断。

- [x] **Step 4: 回归**

运行：
- `pytest backend/tests/test_agent_prompt_fragments.py backend/tests/test_agent_prompt_cache_breakpoint.py -q`

自动化记录：
- `pytest backend/tests/test_agent_prompt_fragments.py backend/tests/test_agent_prompt_cache_breakpoint.py backend/tests/test_agent_loop_service.py backend/tests/test_agent_openai_tool_calls.py -q`
- `npm --prefix add-in run build`

---

## Task 2: 工具结果截断（P0）

**目标：** `range.read`/`workbook.inspect` 返回值超过阈值时截断，附带「还有 N 行」提示，仿 Codex `telemetry_preview`。

**Files:**
- Create: `backend/app/services/agent_context/tool_result_truncation.py`
- Modify: `add-in/src/features/agent/agentWorkbookToolExecutor.ts`（`inspectWorkbook:188-231`、`range.read` 分支）
- Create: `backend/tests/test_tool_result_truncation.py`
- Create: `add-in/src/features/agent/agentToolResultTruncation.ts`

**设计要点：**
- 默认上限：`MAX_ROWS=200`、`MAX_COLUMNS=20`、`MAX_BYTES=8192`。
- 截断策略：保留前 N 行 + 末尾摘要行 `[... 还有 12345 行未显示，请用 range.read 读取 Sheet1!A201:D400 继续 ...]`，给出续读地址。
- 截断信息必须**可操作**（给地址），不能只是「已截断」。
- 同时作用于前端 bridge 返回与后端拼装 transcript 两处，避免重复大对象。

- [x] **Step 1: 写失败测试**（Python 纯函数：给定 500×30 矩阵，断言截断后 ≤ 200 行且含续读提示）

- [x] **Step 2: 实现后端 `tool_result_truncation.py`** 纯函数 `truncate_matrix(values, address, limits)`

- [x] **Step 3: 实现前端 `agentToolResultTruncation.ts`**，在 `inspectWorkbook` 与 `range.read` 返回前调用

- [x] **Step 4: 回归测试大工作簿（≥1 万行）确认上下文不爆**

自动化记录：
- `pytest backend/tests/test_tool_result_truncation.py -q`

自动化记录：
- `pytest backend/tests/test_tool_result_truncation.py backend/tests/test_agent_tool_result_flags.py backend/tests/test_agent_frontend_tool_loop.py backend/tests/test_agent_loop_service.py backend/tests/test_agent_trace_logging.py -q`
- `npm --prefix add-in run build`

---

## Task 3: 公式提示词去重与稳定化（P1）

**目标：** 把 4 段重复公式指令合并为「稳定核心」+「任务可变头」，删除内嵌 JSON 样例（改用 `response_format`），提升缓存命中率与可维护性。

**Files:**
- Modify: `backend/app/services/formula_orchestrator/formula_model_instructions.py`
- Modify: `backend/app/services/formula_orchestrator/formula_prompt_builder.py`
- Modify: `backend/app/services/formula_orchestrator/formula_model_message_builder.py`
- Create: `backend/tests/test_formula_prompt_stability.py`

**设计要点：**
- 抽出 `FORMULA_CORE_RULES`（自检、function_policy 语义、IFERROR 包裹、稳定输出）作为所有 task 共享的稳定前缀。
- `FORMULA_GENERATION_JSON_INSTRUCTION` 当前把 `WORKBOOK_FIRST_VERIFICATION` 整段拼接再重复动态数组/硬编码规则 —— 合并去重。
- 删除 prompt 内嵌的 JSON schema 样例（`:70-87`、`:106-125`），依赖已有的 `response_format`（`openai_request_options.py:167-169`）。样例移到 `docs/14-公式生成中间层重构方案.md`。
- 加入 Codex `continuation.md` 的「完成审计」段落精神，强化「未验证不得声称已验证」。

- [x] **Step 1: 写对比测试** —— 断言重构后 prompt 的稳定前缀（核心规则）与原版语义等价（关键约束词都在），且总长度 ↓ ≥ 30%

- [x] **Step 2: 抽出 `FORMULA_CORE_RULES` 常量，4 段指令改为「核心 + 任务头」组合**

- [x] **Step 3: 删除内嵌 JSON 样例，验证 `response_format` 仍生效**

- [x] **Step 4: 公式生成回归（intent/generate/repair 三条链路）确认输出质量不降**

自动化记录：
- `pytest backend/tests/test_formula_generation_prompt_controls.py backend/tests/test_formula_prompt_stability.py backend/tests/test_formula_response_parser.py backend/tests/test_formula_policy_boundaries.py backend/tests/test_formula_repair_process_events.py backend/tests/test_formula_intent_timeout.py -q`
- 重构后 `FORMULA_GENERATION_JSON_INSTRUCTION=2092 chars`，`FORMULA_REPAIR_JSON_INSTRUCTION=1833 chars`；相对旧基线分别下降 30.1% / 37.4%，已删除 generation/repair prompt 内嵌 JSON schema 示例，改由 `response_format` 约束。

---

## Task 4: 已消费工具结果摘要化（P1）

**目标：** agent 读过一轮的大工具结果，下一轮 LLM 调用前替换为摘要 + Redis 全量指针，仿 Claude-Code `SUMMARIZE_TOOL_RESULTS_SECTION`。

**Files:**
- Modify: `backend/app/services/agent_loop_service.py`（`_run_continue:159-185`）
- Create: `backend/app/services/agent_context/consumed_result_summarizer.py`
- Modify: `backend/app/services/agent_session_store.py`（存储全量结果指针）
- Create: `backend/tests/test_consumed_result_summarizer.py`

**设计要点：**
- 在 `_run_continue` 拼新 transcript 前，遍历上一轮的 `tool` 消息：超过阈值（如 2KB）的，替换为 `[工具结果已摘要：原始 X 行 Y 列存于 redis:agent:toolresult:{id}，如需细节请说明]`。
- 在 `AGENT_SYSTEM_PROMPT` 加一条（仿 `prompts.ts:841`）：「工具结果稍后可能被清除，重要信息请写进你的回复」。
- 全量结果按 `agent_session_id + tool_call_id` 存 Redis，TTL 与 session 一致，支持「展开」。

- [x] **Step 1: 写失败测试**（给定大 tool 消息，断言摘要后 transcript 缩短且含指针 key）

- [x] **Step 2: 实现 `consumed_result_summarizer.py`**

- [x] **Step 3: 在 `_run_continue` 接入，session_store 增加全量结果存取方法**

- [x] **Step 4: AGENT_SYSTEM_PROMPT 补摘要提示条**

- [x] **Step 5: 多轮 agent 回归（≥4 轮工具调用）确认上下文增长被控制**

自动化记录：
- `pytest backend/tests/test_consumed_result_summarizer.py backend/tests/test_agent_loop_service.py -q`

自动化记录：
- `pytest backend/tests/test_consumed_result_summarizer.py backend/tests/test_agent_loop_service.py backend/tests/test_agent_prompt_fragments.py -q`

---

## Task 5: 会话压缩（P2）

**目标：** transcript token 超阈值时，用简短摘要 prompt 替换前 N 轮，仿 Codex `compact.rs`。

**Files:**
- Create: `backend/app/services/agent_context/compaction.py`
- Create: `backend/app/services/agent_context/compaction_prompt.py`（参考 Codex 极简压缩思路，使用本项目自有中文模板）
- Modify: `backend/app/services/agent_loop_service.py`（在 `_run_continue` 开头检查阈值）
- Modify: `backend/app/services/agent_session_store.py`（压缩前后版本快照）
- Create: `backend/tests/test_agent_compaction.py`

**设计要点：**
- 阈值：transcript 估算 token > `model_context_window * 0.7` 时触发。
- 压缩 prompt 使用本项目自有中文模板，只借鉴 Codex 的极简交接摘要思路：
  > 你正在执行上下文检查点压缩。为接手的另一个 LLM 创建交接摘要，包含：当前进度与关键决策、重要约束/偏好、剩余待办（清晰下一步）、继续工作所需的关键数据/示例/引用。保持简洁、结构化。
- 采用 Codex `BeforeLastUserMessage` 策略：压缩摘要注入到最后一条真实 user 消息之前，避免破坏前缀缓存。
- 压缩前后在 session 存快照，支持「撤销压缩」。
- 保留 `max_turn_count` 作为最终兜底，但优先用压缩续命。

- [x] **Step 1: 写失败测试**（给定超长 transcript，断言压缩后前缀被摘要替换且保留最后 user 消息）

- [x] **Step 2: 实现 `compaction_prompt.py` 与 `compaction.py`**

- [x] **Step 3: `agent_loop_service` 接入压缩触发点**

- [x] **Step 4: session_store 增加快照存取**

- [x] **Step 5: 长任务回归（模拟 12+ 轮）确认不再硬截断**

自动化记录：
- `pytest backend/tests/test_agent_compaction.py backend/tests/test_agent_loop_service.py backend/tests/test_consumed_result_summarizer.py -q`
- `npm --prefix add-in run build`

---

## Task 6: AGENTS.md 与工作簿笔记（P1-P2）

**目标：** 支持项目级/工作簿级知识注入，跨会话沉淀字段含义、公式风格、踩坑记录。

**Files:**
- Create: `backend/app/services/memory/agents_md_loader.py`
- Create: `backend/app/services/memory/workbook_notes_store.py`
- Modify: `backend/app/services/agent_loop_service.py`（`_start_messages` 注入笔记）
- Create: `backend/tests/test_agents_md_loader.py`

**设计要点：**
- **AGENTS.md 分层发现**（仿 Codex `agents_md.rs:1-17`）：从工作簿所在目录向上找到项目根（`.git` 标记），收集路径上所有 `AGENTS.md`，按目录顺序拼接，作为稳定 system 段（缓存友好）。
- **工作簿笔记**：每张工作簿可在 DB/Redis 存「已知事实」（字段映射、命名规则、VBA 坑），开新会话时注入。
- 注入位置在 task prompt 之后、workbook 摘要之前（稳定优先）。
- 笔记格式：Markdown，单文件 ≤ 4KB，超出走 Task 7 的两段式召回。

- [x] **Step 1: 写失败测试**（给定目录树含两个 AGENTS.md，断言按序拼接且截断到 4KB）

- [x] **Step 2: 实现 `agents_md_loader.py`**（纯函数，输入 cwd 输出拼接文本）

- [x] **Step 3: 实现 `workbook_notes_store.py`**（Redis 存取，key=`agent:notes:{workbook_fingerprint}`）

- [x] **Step 4: `_start_messages` 注入笔记段**

当前接入说明：当前前端尚未传真实工作簿文件路径；后端仅在 `AgentStartRequest.workbook_path` 可用时注入 AGENTS.md，避免猜测路径。

- [x] **Step 5: 自动化验证：放一个 AGENTS.md，确认 agent 启动上下文受其影响**

自动化记录：
- `pytest backend/tests/test_agents_md_loader.py -q`

自动化记录：
- `pytest backend/tests/test_agents_md_loader.py backend/tests/test_agent_prompt_fragments.py backend/tests/test_agent_loop_service.py -q`
- `npm --prefix add-in run build`

---

## Task 7: 技能化 task 定义（P2）

**目标：** 把前端硬编码的 task prompt 抽成后端可发现的 skill 文件，支持新增任务不改代码。

**Files:**
- Create: `backend/app/services/skills/task_skill_loader.py`
- Create: `backend/app/skills/`（技能目录，每个 task 一个 `.md`）
- Modify: `backend/app/services/agent_loop_service.py`（按 skill 元数据选 prompt + 工具白名单）
- Modify: `add-in/src/features/agent/agentTaskPolicy.ts`（改为从后端拉 skill 列表）
- Create: `backend/tests/test_task_skill_loader.py`

**设计要点：**
- skill 文件格式（仿 Claude-Code `bundledSkills.ts` frontmatter）：
  ```markdown
  ---
  name: formula-generation
  trigger: 用户要求生成或修改公式
  tools: [workbook.inspect, range.read]
  ---
  <system prompt 正文>
  ```
- 后端启动时扫描 `backend/app/skills/`，构建 `{name -> skill}` 注册表。
- agent 启动时按 `task_type` 解析 skill，注入其 system prompt 与工具白名单（与 `AgentToolCapabilities` 叠加）。
- 默认内置 formula-generation / formula-explain / formula-repair / workbook-agent 四个 skill，内容来自现有 `formula_model_instructions.py`。

- [x] **Step 1: 写失败测试**（给定 skills 目录，断言 loader 返回注册表且 frontmatter 解析正确）

- [x] **Step 2: 实现 `task_skill_loader.py`**

- [x] **Step 3: 把现有 4 段公式指令迁移成 skill 文件**

- [x] **Step 4: agent_loop_service 接入 skill 解析**

- [x] **Step 5: 前端 agentTaskPolicy 改为从后端拉列表**

- [x] **Step 6: 全链路回归确认 4 个 task 行为不变**

自动化记录：
- `pytest backend/tests/test_agent_task_policy.py backend/tests/test_task_skill_catalog.py backend/tests/test_task_skill_loader.py backend/tests/test_agent_frontend_tool_loop.py -q`
- `npm --prefix add-in run build`

完成说明：后端已暴露 `/ai/agent/skills` 元数据接口，前端 `agentTaskPolicy.ts` 通过后端 skill catalog 动态判断任务模块是否开放；隐藏 prompt 正文不下发给前端。

---

## Task 8: 工具结果 success / truncated 标志（P2）

**目标：** `AgentToolResult` 增加 `success` 与 `truncated` 字段，让模型能区分「工具失败」vs「读到空」vs「结果被截断需续读」。

**Files:**
- Modify: `backend/app/schemas/agent.py`（`AgentToolResult`）
- Modify: `add-in/src/features/agent/agentWorkbookToolProtocol.ts`（`AgentWorkbookToolResult`）
- Modify: `add-in/src/features/agent/agentWorkbookToolExecutor.ts`（填充新字段）
- Modify: `backend/app/services/agent_loop_service.py`（`_tool_result_message`）
- Create: `backend/tests/test_agent_tool_result_flags.py`

**设计要点：**
- `success: bool`：工具执行成功为 true（即使业务结果为空）。
- `truncated: bool` + `truncation_hint: str | None`：结果被 Task 2 截断时为 true，hint 给续读地址。
- 可选 `contains_external_context: bool`（仿 Codex `tool_output.rs:23`）：未来做 memory 提取时，含外部内容（如 url fetch）的结果不参与记忆生成。

- [x] **Step 1: 写失败测试**（断言 schema 含新字段且 executor 填充正确）

- [x] **Step 2: 后端 schema 加字段（向后兼容，默认值）**

- [x] **Step 3: 前端 protocol + executor 填充字段**

- [x] **Step 4: `_tool_result_message` 透传新字段**

自动化记录：
- `pytest backend/tests/test_agent_loop_service.py backend/tests/test_agent_tool_result_flags.py -q`

- [x] **Step 5: 回归失败工具调用，确认 agent 能识别失败并重试**

---

## Task 9: Goal Checkpoint 与继续审计（P2）

**目标：** 借鉴 Codex `goals/continuation.md` 的完成审计思路，为长任务引入 goal checkpoint。Agent 在每轮结束时记录“已完成、仍缺失、下一步动作、阻塞原因”，下一轮继续时先审计目标状态，而不是从完整 transcript 里重新推断。

**Files:**
- Create: `backend/app/services/agent_context/goal_checkpoint.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Modify: `backend/app/services/agent_session_store.py`
- Create: `backend/tests/test_agent_goal_checkpoint.py`

**Checkpoint 字段：**

```python
class AgentGoalCheckpoint(BaseModel):
    conversation_id: str
    turn_id: str
    objective: str
    completed_items: list[str]
    pending_items: list[str]
    next_action: str | None
    blockers: list[str]
    evidence_refs: list[str]
```

**设计要点：**
- checkpoint 由模型最终响应和工具事件摘要生成，不保存完整敏感 payload。
- continue 时把最近一次 checkpoint 作为 `conversation_summary` fragment 注入。
- 如果 `pending_items` 为空且无 blocker，下一轮提示模型先确认是否已经完成，避免重复执行同一任务。

- [x] **Step 1: 写失败测试**

`test_goal_checkpoint_is_saved_after_agent_turn`：模拟一轮含工具调用和最终响应的 agent turn，断言 session store 写入 `objective/completed_items/pending_items/next_action`。

`test_continue_injects_latest_goal_checkpoint`：已有 checkpoint 时继续对话，断言 `_start_messages` 或 fragment builder 注入 `conversation_summary`。

- [x] **Step 2: 实现 `goal_checkpoint.py`**

提供 `build_goal_checkpoint(turn_events, final_message) -> AgentGoalCheckpoint` 和 `render_checkpoint_fragment(checkpoint) -> str`。

- [x] **Step 3: 接入 session store**

每轮结束后保存 checkpoint；continue 时读取最近一次 checkpoint。

- [x] **Step 4: 回归**

运行：
- `pytest backend/tests/test_agent_goal_checkpoint.py backend/tests/test_agent_loop_service.py -q`

自动化记录：
- `pytest backend/tests/test_agent_goal_checkpoint.py backend/tests/test_agent_loop_service.py backend/tests/test_agent_compaction.py -q`
- `npm --prefix add-in run build`

---

## Task 10: ToolSearch 式延迟工具发现（P2）

**目标：** 当工具数量继续增长时，不把所有工具 schema 一次性塞进模型上下文，而是先暴露 `tool.search`，让模型按任务发现相关工具，再临时启用。

**Files:**
- Create: `backend/app/services/agent_tools/tool_search.py`
- Modify: `backend/app/services/agent_tool_schema.py`
- Modify: `backend/app/services/agent_loop_service.py`
- Create: `backend/tests/test_agent_tool_search.py`

**规则：**
- 默认公开核心工具：`workbook.inspect`、`range.read`、`range.write`、`tool.search`。
- 非核心工具如宏执行、工作表结构操作、报告导出、OCR 填入、图表搭建可 deferred。
- `tool.search` 只返回当前 task policy 允许的工具；被 deny 的工具不出现在搜索结果里。
- 搜索结果返回工具名、用途、风险、是否需要审批，不返回完整 schema；模型选择后后端再注入完整 schema。

- [x] **Step 1: 写失败测试**

`test_tool_search_only_returns_tools_allowed_by_current_task_policy`：在一个受限 task 中搜索高风险工具，返回结果只包含当前 task policy 允许的工具。

- [x] **Step 2: 实现 `tool_search.py`**

BM25 或简单关键词匹配即可，先使用工具 metadata 的 `name/namespace/description/risk`。

- [x] **Step 3: 接入 loop**

当模型调用 `tool.search` 后，把选中的工具 schema 加入下一轮模型请求。

- [x] **Step 4: 回归**

运行：
- `pytest backend/tests/test_agent_tool_search.py backend/tests/test_agent_tool_registry.py -q`

---

## 风险与回滚

| 风险 | 缓解 |
|---|---|
| 提示词缓存改动导致部分厂商报错 | `apply_cache_breakpoints` 按 `provider_id` 白名单启用，不支持时 no-op |
| 工具结果截断丢失关键数据 | 截断阈值保守（200 行），且必须给出续读地址；全量结果仍存 Redis 可展开 |
| 公式提示词去重导致输出质量回归 | Step 1 先写语义等价对比测试，三条公式链路单独回归 |
| 压缩摘要丢失关键上下文 | 保留压缩前快照支持撤销；摘要注入位置遵循 Codex `BeforeLastUserMessage` 不破坏缓存 |
| AGENTS.md 注入引入注入风险 | 只读受信任目录，文件大小硬上限 4KB，内容经现有 guard_service 审计 |
| 后端策略片段迁移导致旧前端调用失败 | `task_system_prompt` 保留兼容期但默认忽略；旧字段只记录 deprecated warning，不参与策略片段 |
| 运行时骨架抽象过重 | 每个模块先用纯函数和小 schema 实现，只有被两个以上任务复用后再沉淀公共接口 |

**回滚策略：** 每个 Task 独立提交，feature flag 控制（`agent_context_*` 配置项），出问题可单 Task 回滚而不影响其他。

---

## 验收标准

- [x] **M0：** 后端能产出 task policy fragment、tool registry metadata、approval event、trace event；前端不再自己拼装完整 task prompt 或工具说明。
- [x] **M1：** 一次典型 agent 会话（含 2 次工具调用）的 `cached_tokens` > 0（支持的厂商）；1 万行工作簿的 `workbook.inspect` 不触发上下文超限。
- [x] **M2：** 公式 prompt 稳定前缀长度 ↓ ≥ 30%；4 轮工具调用后 transcript token 增长 ≤ 20%（摘要化生效）；用户气泡不展示隐藏提示词，追问不重复注入模块提示词。
- [x] **M3：** 模拟 12 轮长任务不再触发 `turn limit exceeded`；放入 AGENTS.md 后 agent 行为可观测受其影响；continue 时注入最近一次 goal checkpoint。
- [x] **M4：** 新增一个 skill 文件无需改代码即可在 agent 中使用；ToolSearch 只返回当前 task 可用工具；失败工具调用 agent 能识别并重试。
- [x] 全部现有 pytest 回归通过；新增模块测试覆盖率 ≥ 80%。

验证记录：
- `pytest backend/tests -q`：205 passed
- `pytest backend/tests -q --cov=app.services.agent_context --cov=app.services.agent_policy --cov=app.services.agent_tools --cov=app.services.skills --cov=app.services.memory --cov=app.services.agent_trace_service --cov-report=term-missing --cov-fail-under=80`：205 passed，Total coverage 88.58%
- `npm --prefix add-in run build`：passed
- `pytest backend/tests/test_agent_openai_tool_calls.py backend/tests/test_agent_prompt_cache_breakpoint.py -q`：13 passed
- Live smoke（OpenAI-compatible `gpt-5.5`，两步工具调用路径）：第 2 次请求 `cached_tokens=2944`，工具调用数 `[1, 1]`

**最小验证命令：**

```powershell
pytest backend/tests/test_agent_task_policy.py `
  backend/tests/test_agent_tool_registry.py `
  backend/tests/test_agent_capability_sandbox.py `
  backend/tests/test_agent_approval_packet.py `
  backend/tests/test_agent_trace_logging.py `
  backend/tests/test_agent_prompt_fragments.py `
  backend/tests/test_agent_tool_result_flags.py `
  backend/tests/test_agent_goal_checkpoint.py `
  backend/tests/test_agent_tool_search.py -q

npm --prefix add-in run build
```

**能力补强证据矩阵：**

该矩阵只核验参考项目机制在当前项目中的落地效果。

| 能力 | 自动化证据 | 运行时证据 | 通过标准 |
|---|---|---|---|
| Prompt cache section | `test_prompt_cache_breakpoints_keep_stable_prefix` | request summary 记录 `stable_system/task_policy/dynamic_context` fragment 顺序 | 稳定前缀不被工作簿快照和会话消息破坏 |
| Tool output preview | `test_large_tool_result_is_truncated_with_resume_hint` | tool result 含 `truncated=true/resume_hint/contains_external_context` | 大范围读取不会把完整表格塞进模型上下文 |
| Consumed tool summary | `test_consumed_tool_results_are_summarized` | transcript 中已消费工具结果变成摘要引用 | 多轮任务 token 增长受控 |
| Compaction | `test_compaction_preserves_last_user_message` | session store 保存 compaction snapshot | 长任务压缩后还能继续执行 |
| Memory retrieval | `test_memory_filters_external_context` | memory event 记录候选数、入选数和外部内容过滤数 | 工作簿笔记可召回，外部内容不会污染记忆 |
| Skill discovery | `test_task_skill_loader_discovers_frontmatter` | skill registry 暴露 name、description、tool hints | 新增 skill 文件无需改前端硬编码 |
| Goal checkpoint | `test_goal_checkpoint_is_saved_after_agent_turn` | continue 注入最近一次 checkpoint fragment | 继续执行不用从完整 transcript 重新猜当前进度 |
| ToolSearch | `test_agent_loop_executes_tool_search_internally_and_enables_deferred_schema` | tool.search event 记录 query、matched tools、deferred schema ids | 工具数量增长时仍能按需加载 schema |

---

## 参考来源

- **Claude-Code**（`github.com/pengchengneo/Claude-Code`，TypeScript 逆向实现）：
  - `src/constants/systemPromptSections.ts` —— 系统提示词分段缓存
  - `src/constants/prompts.ts:114-576` —— 缓存边界标记与静态/动态段分离
  - `src/memdir/findRelevantMemories.ts` —— 两段式记忆召回
  - `src/memdir/memoryScan.ts` —— 记忆文件 frontmatter 扫描
  - `src/skills/loadSkillsDir.ts` —— 技能发现机制
- **OpenAI Codex**（`github.com/openai/codex`，Rust 官方实现）：
  - `codex-rs/core/src/compact.rs` —— 双时机上下文压缩
  - `codex-rs/core/src/agents_md.rs` —— AGENTS.md 分层发现
  - `codex-rs/tools/src/tool_output.rs` —— 工具结果截断 + `contains_external_context`
  - `codex-rs/tools/src/tool_search.rs` —— 工具 defer_loading
  - `codex-rs/prompts/templates/compact/prompt.md` —— 极简压缩 prompt
  - `codex-rs/prompts/templates/goals/continuation.md` —— 完成审计提示词
  - `codex-rs/prompts/templates/apply_patch_tool_instructions.md` —— 确定性编辑工具提示词
  - `codex-rs/memories/README.md` —— 两阶段 memory pipeline
