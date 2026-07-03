# Bug 修复记录 — API 400 错误：assistant tool_calls 消息后缺少对应 tool result

> 日期：2026-06-24
> 分支：feature/new-feature
> 修复文件：`desktop/electron/agent/aiClient.ts`

---

## 一、问题描述

### 现象

使用 AI 对话时，发送消息后返回 HTTP 400 错误：

```
API 请求失败 (400): {
  "error": {
    "message": "An assistant message with 'tool_calls' must be followed by tool messages responding to each 'tool_call_id'. (insufficient tool messages following tool_calls message)",
    "type": "invalid_request_error",
    "param": null,
    "code": "invalid_request_error"
  }
}
```

### 触发条件

当对话历史中存在**孤立的 `tool_call`**（即 `tool_call` 没有对应的 `tool_result`）时，API 请求会被拒绝。OpenAI 协议要求每一条包含 `tool_calls` 的 assistant 消息后面，必须紧跟对应的 `tool` 消息，且 `tool_call_id` 一一匹配。

---

## 二、根因分析

### 2.1 数据流概览

```
用户输入
  → agentLoop 产生 TurnItem[]（持久化到 JSONL）
    → turnItemsToChatMessages() 将 TurnItem[] 转换为 ChatMessage[]
      → buildRequestMessages() 将 ChatMessage[] 转换为 API 请求格式
        → 发送给 AI API
```

### 2.2 `turnItemsToChatMessages()` 的三个缺陷

**文件**：`electron/agent/aiClient.ts`，原第 859-962 行

#### 缺陷 1：不校验 tool_call / tool_result 配对

原代码逐项遍历 `TurnItem[]`，遇到 `tool_call` 就无条件追加到 assistant 消息，遇到 `tool_result` 就无条件插入 tool 消息。没有检查两者是否一一配对。

**孤立 tool_call 的产生场景**：

| 场景 | 原因 |
|------|------|
| 历史压缩（compaction） | compaction 可能保留 `tool_call` 但移除对应的 `tool_result` |
| 执行中断 | agentLoop 执行工具时异常退出，`tool_result` 未写入 |
| 外部编辑 | 用户或程序手动修改了 JSONL 历史文件 |

#### 缺陷 2：交错模式下 tool_call 分组错误

agentLoop 的 TurnItem 顺序为：

```
assistant_message → tool_call(tc1) → tool_call(tc2) → tool_result(tc1) → tool_result(tc2)
```

但如果因为中断或重试导致交错排列：

```
assistant_message → tool_call(tc1) → tool_result(tc1) → tool_call(tc2) → tool_result(tc2)
```

原代码在遇到第二个 `tool_call(tc2)` 时，会追加到上一条消息（此时已经是 `tool` 消息，不是 `assistant`），于是创建**新的** assistant 消息。如果 `tc2` 的 `tool_result` 丢失，这条新 assistant 消息就是孤立的。

#### 缺陷 3：无清理空壳消息

如果所有 `tool_call` 都被某种方式移除，assistant 消息可能变成 `content=""` 且 `toolCalls=[]` 的空壳，依然会被发送到 API。

### 2.3 `buildRequestMessages()` 的缺陷

**文件**：`electron/agent/aiClient.ts`，原第 230-271 行

原代码直接将 `ChatMessage[]` 转换为 API 请求格式，没有任何校验。即使 `turnItemsToChatMessages` 产生了不合规的消息序列，这里也会原样发出。

---

## 三、修复方案

### 3.1 `turnItemsToChatMessages()` — 三遍处理

```
第一遍：收集所有 tool_call ID 和 tool_result ID
        → 识别孤立的 tool_call（有 call 无 result）

第二遍：构建 ChatMessage[]
        → 跳过孤立的 tool_call
        → 跳过无对应 tool_call 的 tool_result（防御性）
        → tool_call 只追加到尚未插入 tool 消息的 assistant 消息
          （如果上一条是 tool 消息，则创建新 assistant 消息）

第三遍：清理空的 assistant 消息
        → 移除 content="" 且 toolCalls=[] 的空壳消息
```

**设计决策**：对于孤立的 `tool_call`，选择**跳过**而非**合成错误 result**。理由：

- `tool_call` 没有对应的 `tool_result`，意味着该工具从未成功执行，模型也从未看到过结果
- 与其虚构一个结果，不如让模型从历史中完全不感知这次调用，保持上下文一致性
- 如果合成假 result，模型可能基于错误信息做出错误决策

### 3.2 `buildRequestMessages()` — 双重安全校验

作为最后一道防线，在构建 API 请求前进行校验：

```
首次校验：全局 ID 匹配
  → 收集所有 tool 消息的 tool_call_id
  → 过滤掉没有对应 tool 消息的 toolCalls 项
  → 如果全部孤立，移除整条 assistant 消息（保留文本内容则仅移除 toolCalls）

二次校验：连续性验证
  → 确保每条 assistant(toolCalls) 后紧跟对应的 tool 消息
  → 如果 tool 消息被其他消息（如 user 消息）打断，
    导致部分 toolCall 缺失 result，从 assistant 中移除缺失项
  → 如果移除后 toolCalls 为空，删除整条 assistant 消息
```

---

## 四、改动详情

### 文件：`desktop/electron/agent/aiClient.ts`

#### 函数 `turnItemsToChatMessages()`（完全重写）

| 原代码行 | 改动 |
|----------|------|
| 859-962 | 三遍处理替换原来的单遍遍历 |

关键改动点：

1. **新增第一遍扫描**：收集 `toolCallIds`、`toolResultIds`、`orphanedToolCallIds` 三个 Set
2. **`tool_call` 分支**：增加 `orphanedToolCallIds.has(item.id)` 判断，跳过孤立项；追加上一条消息前检查 `lastMsg.role === "assistant"`
3. **`tool_result` 分支**：增加 `!toolCallIds.has(item.toolCallId)` 判断，跳过无对应 call 的 result
4. **新增第三遍清理**：倒序遍历移除空壳 assistant 消息

#### 函数 `buildRequestMessages()`（增加校验逻辑）

| 原代码行 | 改动 |
|----------|------|
| 230-271 | 在构建 API 格式前增加两轮校验 |

关键改动点：

1. **首次校验**：`toolResultIds` Set 收集所有 tool 消息 ID → 过滤 `validToolCalls` → 处理全孤立/部分孤立/全部配对三种情况
2. **二次校验**：逐条遍历 `sanitizedMessages`，用 `expectedIds` / `foundIds` 追踪配对状态 → 未配对的从 assistant 中移除
3. **最终构建**：对校验后的 `finalMessages` 执行原有的 API 格式转换

---

## 五、验证结果

| 验证项 | 结果 |
|--------|------|
| TypeScript 编译（`tsconfig.json`） | ✅ 0 错误 |
| TypeScript 编译（`tsconfig.electron.json`） | ✅ 0 错误 |
| Vite 构建（renderer） | ✅ `index.html` + CSS + JS |
| Vite 构建（main process） | ✅ `main.js` |
| Vite 构建（preload） | ✅ `preload.js` |

---

## 六、影响范围

- **直接影响**：修复所有因孤立 `tool_call` 导致的 API 400 错误
- **间接影响**：压缩历史后的对话请求将更加健壮，不会因 compaction 剔除部分 `tool_result` 而崩溃
- **行为变化**：孤立的 `tool_call` 不会出现在发送给 AI 的历史中（模型不再感知未完成的工具调用），这是预期行为
- **无破坏性变更**：正常的 `tool_call` / `tool_result` 配对不受影响

---

## 七、相关代码路径

```
desktop/electron/agent/aiClient.ts
  ├─ turnItemsToChatMessages()   ← 主要修复点：三遍处理
  └─ buildRequestMessages()      ← 安全校验：双重配对验证

desktop/electron/agent/agentLoop.ts
  └─ tool_call/tool_result 产生逻辑  ← 只读参考，未修改
     （agentLoop 总是成对产生 tool_call + tool_result，
       包括用户拒绝和审批出错场景）

desktop/electron/agent/compaction.ts
  └─ 历史压缩逻辑  ← 只读参考，未修改
     （compaction 可能导致 tool_result 丢失，
       本次修复在消费端防御）
```
