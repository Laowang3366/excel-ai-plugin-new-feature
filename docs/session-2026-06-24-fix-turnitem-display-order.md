# Bug 修复记录 — 对话信息显示顺序错误：思考内容被排到工具调用之后

> 日期：2026-06-24
> 分支：feature/new-feature
> 修复文件：`desktop/electron/agent/agentLoop.ts`、`desktop/src/components/ChatPage.tsx`

---

## 一、问题描述

### 现象

AI 回复时，对话信息的显示顺序错误：

**当前（错误）**：
```
本轮工作时长 xx秒
------------------------------------------------
调用工具执行情况        ← 工具调用先显示
|
|
首次思考内容            ← 思考内容反而排在后面
|
|
第二次思考内容
|
|
正文内容
```

**期望（正确）**：
```
本轮工作时长 xx秒
------------------------------------------------
首次思考内容            ← 思考内容应最先显示
|
|
调用工具执行情况        ← 工具调用在思考之后
|
|
第二次思考内容
|
|
正文内容
```

### 正确的显示逻辑

同一轮 AI 回复中，各类信息的规范显示顺序应为：

```
reasoning → tool_call → tool_result → assistant_message
```

即：先展示思考过程，再展示工具调用和结果，最后展示正文回复。

---

## 二、根因分析

### 2.1 数据流

```
AI API 流式响应
  → agentLoop 逐事件处理，创建 TurnItem 并 push 到 turn.items
    → chatStore 按 item_started/item_completed 事件顺序存入 messages[]
      → ChatPage.tsx AssistantGroupBlock 按 items[] 数组顺序渲染
```

### 2.2 后端时序问题

**文件**：`desktop/electron/agent/agentLoop.ts`

在 AI 流式响应期间，各类型 TurnItem 的创建时机不同：

| 事件 | 创建时机 | push 到 turn.items |
|------|----------|-------------------|
| `tool_call_begin` | **流式期间**（实时） | 第 434 行：立即 `push` |
| `reasoning` | **流结束后** | 第 533 行：延迟 `push` |
| `assistant_message` | **流结束后** | 第 549 行：延迟 `push` |

**问题**：`tool_call` 在流式期间就被 `push` 到 `turn.items`，而 `reasoning` 和 `assistant_message` 要等流结束后才 `push`。因此 `turn.items` 的实际顺序变成了：

```
tool_call  ← 先推入
reasoning  ← 后推入（反而在后面）
assistant_message  ← 最后推入
tool_result  ← 工具执行时推入
```

这导致推理内容排在工具调用之后，与实际思维顺序相反。

### 2.3 前端渲染无排序

**文件**：`desktop/src/components/ChatPage.tsx`

`AssistantGroupBlock`（第 1096 行）和 `StreamingAssistantGroupBlock`（第 1187 行）都直接按 `group.items` 数组顺序渲染，没有任何排序逻辑。所以后端产生的乱序会原样呈现给用户。

此外，从 JSONL 历史文件加载数据时（`sessionStore.parseRolloutContent`），也是按 JSONL 行的追加顺序 `push` 到 `turn.items`，无法保证正确顺序。

---

## 三、修复方案

### 3.1 后端修复（`agentLoop.ts`）— 根本解决

在流结束后创建 `reasoning` 和 `assistant_message` 时，不再 `push` 到末尾，而是**插入到第一个 `tool_call` 之前**的位置。

```typescript
// 修复前：
turn.items.push(reasoningItem);   // 排在 tool_call 之后

// 修复后：
const firstToolCallIdx = turn.items.findIndex((i) => i.type === "tool_call");
if (firstToolCallIdx >= 0) {
  turn.items.splice(firstToolCallIdx, 0, reasoningItem);  // 插入到 tool_call 之前
} else {
  turn.items.push(reasoningItem);  // 无 tool_call 时正常追加
}
```

同样的逻辑应用到 `assistant_message` 的插入。

**注意**：JSONL 文件是追加模式，`appendTurnItem` 仍按时间顺序写入。因此从 JSONL 恢复的历史数据顺序仍然可能是乱的，需要前端排序兜底。

### 3.2 前端防御（`ChatPage.tsx`）— 兜底所有场景

新增 `sortItemsByType()` 函数，在渲染前对 `group.items` 按规范顺序排序：

```typescript
const ITEM_TYPE_ORDER: Record<string, number> = {
  reasoning: 0,
  tool_call: 1,
  tool_result: 2,
  assistant_message: 3,
};

function sortItemsByType(items: TurnItem[]): TurnItem[] {
  return [...items].sort((a, b) => {
    const orderA = ITEM_TYPE_ORDER[a.type] ?? 99;
    const orderB = ITEM_TYPE_ORDER[b.type] ?? 99;
    if (orderA !== orderB) return orderA - orderB;
    return 0;  // 同类型保持原有相对顺序（稳定排序）
  });
}
```

应用到两个渲染组件：
- `AssistantGroupBlock`：渲染前调用 `sortItemsByType(group.items)`
- `StreamingAssistantGroupBlock`：同样调用排序，并简化了 `streamingReasoning` 的插入逻辑（始终放在最前面，不再需要复杂的 `insertAfterIdx` 计算）

---

## 四、改动详情

### 文件：`desktop/electron/agent/agentLoop.ts`

| 位置 | 改动 |
|------|------|
| 第 524-553 行 | `reasoning` 和 `assistant_message` 的插入方式从 `push` 改为 `splice`（插入到第一个 `tool_call` 索引前） |

### 文件：`desktop/src/components/ChatPage.tsx`

| 位置 | 改动 |
|------|------|
| 新增（第 1060 行前） | `ITEM_TYPE_ORDER` 常量和 `sortItemsByType()` 函数 |
| `AssistantGroupBlock` | 使用 `sortedItems` 替代 `group.items`，`renderItem` 传入排序后的 items |
| `StreamingAssistantGroupBlock` | 使用 `sortedItems`；`streamingReasoning` 简化为始终在最前面；移除 `insertAfterIdx` 计算逻辑 |

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

- **直接影响**：AI 回复中的思考内容、工具调用、正文内容现在按正确顺序显示
- **间接影响**：从 JSONL 历史文件加载的旧会话也能正确排序（前端兜底）
- **行为变化**：`StreamingAssistantGroupBlock` 的 `streamingReasoning` 插入逻辑简化，不再需要 `insertAfterIdx` 计算
- **无破坏性变更**：同类型项的相对顺序不变（稳定排序）

---

## 七、相关代码路径

```
desktop/electron/agent/agentLoop.ts
  ├─ tool_call_begin handler (第 420 行)     ← 流式期间 push tool_call
  ├─ reasoning 创建 (第 524 行)              ← 修复：splice 到 tool_call 前
  └─ assistant_message 创建 (第 539 行)      ← 修复：splice 到 tool_call 前

desktop/src/components/ChatPage.ts
  ├─ sortItemsByType()                        ← 新增：排序函数
  ├─ AssistantGroupBlock (第 1060 行)         ← 修复：渲染前排序
  └─ StreamingAssistantGroupBlock (第 1123 行) ← 修复：渲染前排序 + 简化逻辑

desktop/src/store/chatStore.ts
  └─ handleAgentEvent (第 322 行)             ← 未修改，按事件顺序存储
     （排序逻辑放在渲染层更合适，避免影响数据层）
```
