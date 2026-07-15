# 代码审查标准与流程

> 生成日期：2026-07-05 | 最近修订：2026-07-15
> 适用项目：Excel AI 插件（Electron + React + TypeScript）
> 关联文档：`docs/development-standards.md`（开发规范）、`docs/code-review-plan.md`（历史审查记录）
> 运行事实以 `desktop/package.json`、`product-site/package.json`、根 `global.json` 和 `.github/workflows/` 为准；本文不复制易变化的测试数量。

---

## 目录

1. [总则](#一总则)
2. [审查标准（Checklist）](#二审查标准checklist)
3. [审查流程](#三审查流程)
4. [自动化工具链](#四自动化工具链)
5. [CI 门禁规则](#五ci-门禁规则)
6. [审查模板](#六审查模板)
7. [审查文化](#七审查文化)
8. [附录](#八附录)

---

## 一、总则

### 1.1 文档定位

本文档是项目的**代码审查操作手册**，回答四个问题：

| 问题 | 本文档对应章节 |
|------|--------------|
| 审什么？ | [审查标准](#二审查标准checklist) |
| 怎么审？ | [审查流程](#三审查流程) |
| 机器能自动查的？ | [自动化工具链](#四自动化工具链) + [CI 门禁](#五ci-门禁规则) |
| 人怎么沟通？ | [审查模板](#六审查模板) + [审查文化](#七审查文化) |

### 1.2 与现有规范的关系

```
docs/development-standards.md   ← 代码该怎么写（编码规范）
docs/code-review-standards.md   ← 代码该怎么查（本文档）
docs/code-review-plan.md        ← 历史审查执行记录（归档）
```

`development-standards.md` 定义了六大编码规范（模块拆分、接口精简、持久化优化、测试基础设施、IPC 解耦、Bug 防御），本文档将其转化为**可执行的审查 checklist**，并补充流程和工具链。

### 1.3 适用范围

- **适用**：所有合入主分支的代码变更（Pull Request / Merge Request）
- **不适用**：紧急 Hotfix 可走快速通道（见 [3.6 紧急通道](#36-紧急通道hotfix)），但事后需补审

### 1.4 审查目标

```
         ┌──────────────────────────────────────────┐
         │              审查的真正目标               │
         ├──────────────────────────────────────────┤
         │  1. 拦截缺陷：安全漏洞、数据丢失、崩溃    │
         │  2. 传递知识：让团队理解"为什么这样写"     │
         │  3. 统一风格：减少认知负担，降低维护成本   │
         │  4. 培养习惯：让好代码成为肌肉记忆         │
         └──────────────────────────────────────────┘
```

---

## 二、审查标准（Checklist）

### 2.0 优先级体系

每条审查意见必须标注优先级，便于作者判断处理顺序：

| 标记 | 级别 | 含义 | 处理要求 |
|------|------|------|----------|
| 🔴 P0 | **阻断** | 安全漏洞、数据丢失、崩溃、API 契约破坏 | **必须修复后才能合并** |
| 🟡 P1 | **建议** | 输入校验缺失、命名不清、缺少测试、性能问题 | **应该修复**，如有充分理由可跳过 |
| 💭 P2 | **微调** | 风格细节、文档补全、可选的替代方案 | **酌情处理**，不阻塞合并 |
| ✅ | **亮点** | 值得学习的好实践 | 审查者应主动表扬 |

---

### 2.1 安全性审查

> Electron 桌面应用，主进程有 Node.js 完整权限，安全审查不可省略。

#### 🔴 P0 — 必须修复

| # | 检查项 | 说明 |
|---|--------|------|
| S1 | **IPC 输入校验** | 新增 IPC 通道必须使用 Zod schema 校验输入（见 `ipcSchemas.ts`）。未校验的 IPC 通道可被恶意网页利用 |
| S2 | **命令注入防护** | `shell.execute` 等命令执行工具必须经过沙箱策略引擎审批，禁止拼接用户输入到命令字符串 |
| S3 | **Electron 安全配置** | `webPreferences.nodeIntegration: false`、`contextIsolation: true` 必须保持。新增 BrowserWindow 需审查安全配置 |
| S4 | **敏感信息泄露** | 代码中不得硬编码 API Key、密钥、Token。日志中不得打印完整密钥或用户隐私数据 |
| S5 | **路径穿越防护** | 文件操作必须校验路径不超出允许范围（防止 `../../` 逃逸） |

#### 🟡 P1 — 应该修复

| # | 检查项 | 说明 |
|---|--------|------|
| S6 | **XSS 防护** | React 组件中禁止使用 `dangerouslySetInnerHTML`，除非内容已消毒。Markdown 渲染需确认 `react-markdown` 配置安全 |
| S7 | **依赖安全** | 新增 npm 依赖需检查是否有已知漏洞（`npm audit`）。避免引入已废弃或无维护的包 |

---

### 2.2 正确性审查

#### 🔴 P0 — 必须修复

| # | 检查项 | 说明 |
|---|--------|------|
| C1 | **逻辑正确** | 代码行为与意图一致。条件判断、循环边界、状态转换需逐一验证 |
| C2 | **错误处理** | 关键路径（IPC 调用、文件操作、AI API 请求）必须有 try-catch，不得静默吞异常 |
| C3 | **异步安全** | Promise 链必须有 catch；async/await 不得遗漏 await；并发操作需考虑竞态条件 |
| C4 | **类型安全** | 禁止 `any`（除非有注释说明原因）。可选属性必须用 `?.` 链式访问。`as` 断言需有类型守卫保障 |

```typescript
// 🔴 P0 示例：缺少错误处理
const data = await ipcApi.agent.startTurn(input);
return data.result; // 如果 startTurn 抛异常，整个流程崩溃

// ✅ 正确：关键路径有错误处理
try {
  const data = await ipcApi.agent.startTurn(input);
  return data.result;
} catch (err) {
  log.error("startTurn failed", { error: err instanceof Error ? err.message : String(err) });
  return { success: false, error: "Turn 启动失败" };
}
```

#### 🟡 P1 — 应该修复

| # | 检查项 | 说明 |
|---|--------|------|
| C5 | **边界条件** | 空数组、空字符串、undefined/null、极大值等边界情况是否处理 |
| C6 | **API 消息格式** | AI API 调用中，assistant 消息含 tool_calls 时必须有对应 tool result（防 400 错误，见开发规范 6.2） |
| C7 | **事件顺序** | 事件驱动架构中，流式阶段不立即发出中间状态事件（见开发规范 6.1） |

---

### 2.3 可维护性审查

> 本项目的文件行数和接口设计规范（见 `development-standards.md`）是硬性约束。

#### 🔴 P0 — 必须修复

| # | 检查项 | 上限 | 说明 |
|---|--------|------|------|
| M1 | **文件行数** | TS/TSX ≤ 400 行，React 组件 ≤ 300 行，Store ≤ 400 行，CSS ≤ 500 行 | 超限必须拆分后才能合入 |
| M2 | **分层依赖** | — | 新代码必须遵守分层架构：`interaction → runtime → core`，core 不直接依赖具体实现。`architecture.test.ts` 会自动检查 |

#### 🟡 P1 — 应该修复

| # | 检查项 | 说明 |
|---|--------|------|
| M3 | **Props 数量** | 组件 props ≤ 10 个。超过时使用 hook 返回值整体传入（见开发规范 2.1） |
| M4 | **命名清晰** | 函数名是动词、变量名是名词、布尔值以 is/has/should 开头。避免 `data`、`temp`、`info` 等无意义命名 |
| M5 | **单一职责** | 一个函数/类只做一件事。如果函数超过 50 行，考虑提取子函数 |
| M6 | **注释必要** | 复杂逻辑必须有注释说明"为什么"。简单代码不需要注释说"做什么"——代码本身就是说明 |
| M7 | **IPC 抽象** | 禁止直接访问 `window.electronAPI`，必须通过 `src/services/ipcApi.ts`（见开发规范 5.1） |
| M8 | **结构化日志** | 禁止 `console.log`，必须使用 `electron/shared/logger.ts`（见开发规范 5.3） |

#### 💭 P2 — 酌情处理

| # | 检查项 | 说明 |
|---|--------|------|
| M9 | **代码重复** | 超过 3 处的重复逻辑应提取为公共函数 |
| M10 | **魔法数字** | 硬编码的数字应提取为有意义的常量 |

---

### 2.4 性能审查

#### 🟡 P1 — 应该修复

| # | 检查项 | 说明 |
|---|--------|------|
| P1-perf | **N+1 查询** | 列表数据的逐项 IPC 查询必须改为批量查询（见开发规范 6.4） |
| P2-perf | **持久化优化** | 设置变更必须使用 `savePartial` 增量写入，禁止全量写入（见开发规范 3.1） |
| P3-perf | **渲染性能** | React 组件避免在 render 中创建新对象/数组（导致不必要的重渲染）。大列表使用虚拟化 |
| P4-perf | **内存泄漏** | useEffect 清理函数必须清理定时器、事件监听器、IPC 监听器 |

```typescript
// 🟡 P1 示例：N+1 查询
// ❌ 每个 turn 一次 IPC
for (const turn of turns) {
  const usage = await ipcApi.settings.get(`usage_${turn.turnId}`);
}

// ✅ 批量查询
const allUsage = await ipcApi.settings.getAll();
```

#### 💭 P2 — 酌情处理

| # | 检查项 | 说明 |
|---|--------|------|
| P5-perf | **不必要的分配** | 循环内避免重复创建对象、字符串拼接用数组 join |
| P6-perf | **防抖/节流** | 高频事件（搜索输入、窗口 resize）应使用防抖或节流 |

---

### 2.5 测试审查

#### 🔴 P0 — 必须修复

| # | 检查项 | 说明 |
|---|--------|------|
| T1 | **现有测试通过** | `npm test` 必须全部通过。不得合并破坏现有测试的 PR |
| T2 | **类型检查通过** | `npm run typecheck` 必须通过 |

#### 🟡 P1 — 应该修复

| # | 检查项 | 说明 |
|---|--------|------|
| T3 | **核心逻辑有测试** | 纯函数、工具执行逻辑、状态转换逻辑必须有 `.test.ts` 文件（见开发规范 4.2） |
| T4 | **测试覆盖边界** | 测试不仅覆盖正常路径，还需覆盖空值、异常、边界条件 |
| T5 | **Mock 标记** | Mock 数据必须有 `@MOCK_INTERFACE` 注释（见开发规范 4.4） |
| T6 | **测试命名** | 测试描述应表达意图："should return empty array when input is null" 而非 "test1" |

---

### 2.6 项目特定规范审查

> 以下检查项来自 `development-standards.md`，是本项目的硬性约定。

| # | 检查项 | 级别 | 说明 |
|---|--------|------|------|
| PR1 | IPC 通道命名格式 `domain:action` | 🟡 P1 | 如 `agent:startTurn`、`settings:get` |
| PR2 | Mock 数据有 `@MOCK_INTERFACE` 标记 | 🟡 P1 | 便于后续接入真实实现时定位 |
| PR3 | 新 IPC 通道有 Zod schema | 🔴 P0 | 见 `ipcSchemas.ts` |
| PR4 | electronApi.d.ts 与 types.ts 同步 | 🟡 P1 | 前端视角与主进程视角的类型声明保持一致 |
| PR5 | CHANGELOG.md 已更新 | 🟡 P1 | 使用 Keep a Changelog 风格 |
| PR6 | 新文件遵循目录结构约定 | 🟡 P1 | electron/agent 分层、src/components 按功能域组织 |

---

## 三、审查流程

### 3.1 流程总览

```
┌─────────┐    ┌──────────┐    ┌──────────┐    ┌──────────┐    ┌──────────┐
│  开发    │───▶│  自查     │───▶│  提交     │───▶│  审查     │───▶│  合并     │
│  分支    │    │  清单     │    │  PR      │    │  反馈     │    │  入库     │
└─────────┘    └──────────┘    └──────────┘    └──────────┘    └──────────┘
                                    │               │
                                    │               ▼
                                    │          ┌──────────┐
                                    │          │  修改     │
                                    │          │  迭代     │
                                    │          └────┬─────┘
                                    │               │
                                    └───────────────┘
```

### 3.2 分支策略

```
main (受保护)
  │
  ├── feature/xxx    ← 功能开发分支
  ├── fix/xxx        ← Bug 修复分支
  ├── refactor/xxx   ← 重构分支
  └── hotfix/xxx     ← 紧急修复分支（可直接合 main，事后补审）
```

**规则**：
- `main` 分支受保护，禁止直接 push
- 所有变更通过 PR 合入
- 分支命名：`{类型}/{简短描述}`，如 `feature/excel-chart-support`、`fix/orphan-tool-call-400`

### 3.3 PR 提交前自查

开发者提交 PR 前，**必须**完成以下自查（对应 `development-standards.md` 第 4.5 节）：

```markdown
## 提交前自查清单

- [ ] 新文件 ≤ 400 行，组件 ≤ 300 行
- [ ] 组件 props ≤ 10 个，或使用 hook 归组
- [ ] 设置变更使用 savePartial 增量写入
- [ ] 核心逻辑有对应 .test.ts 文件
- [ ] IPC 调用使用 ipcApi，非 window.electronAPI
- [ ] 新增 IPC 通道有 zod schema 校验
- [ ] mock 数据有 @MOCK_INTERFACE 标记
- [ ] 无 console.log，使用 logger
- [ ] npm run typecheck 通过
- [ ] npm test 全部通过
- [ ] CHANGELOG.md 已更新
```

### 3.4 审查者分配

| PR 规模 | 审查者数量 | 说明 |
|---------|-----------|------|
| 小型（< 100 行变更） | 1 人 | 常规审查 |
| 中型（100–500 行） | 1 人 + 1 人确认 | 一人主审，一人确认 |
| 大型（> 500 行） | 2 人 | 两人独立审查 |
| 架构变更 | 2 人 + 技术负责人 | 需技术负责人签字 |

**审查者指派原则**：
- 优先指派对变更模块熟悉的成员
- 避免审查自己写的代码（作者不能是唯一审查者）
- 新成员的 PR 由资深成员审查

### 3.5 审查时效

| 阶段 | 时效要求 | 超时处理 |
|------|----------|----------|
| 首次响应 | **4 小时内**（工作时间内） | 自动提醒审查者 |
| 审查完成 | **1 个工作日内** | 作者可申请更换审查者 |
| 作者响应反馈 | **1 个工作日内** | 审查者可标记 stale |
| 争议升级 | **2 小时内**有回应 | 升级到技术负责人 |

### 3.6 紧急通道（Hotfix）

```
生产环境紧急 Bug
       │
       ▼
┌──────────────┐     ┌────────────────┐     ┌──────────────┐
│ 技术负责人    │────▶│ 直接合入 main  │────▶│ 24h 内补审   │
│ 授权快速通道  │     │ （跳过审查）    │     │ 创建追溯 PR  │
└──────────────┘     └────────────────┘     └──────────────┘
```

**Hotfix 条件**（满足任一）：
- 生产环境崩溃或数据丢失
- 用户无法使用核心功能
- 安全漏洞紧急修复

**Hotfix 后要求**：
- 24 小时内创建追溯审查 PR，补充审查记录
- 在 CHANGELOG.md 中标注 `[HOTFIX]`

### 3.7 合并规则

| 条件 | 要求 |
|------|------|
| CI 通过 | typecheck + test + build 全部绿色 |
| P0 意见 | 全部解决或标记"已知风险，后续处理" |
| P1 意见 | 至少回应（修复 / 解释跳过原因） |
| 审查者 Approve | 至少 1 人 Approve（大型 PR 需 2 人） |
| 冲突解决 | 无合并冲突，与 main 同步 |

**合并方式**：使用 **Squash and Merge**（保持 main 历史整洁，每个 PR 一个 commit）。

### 3.8 争议处理

```
审查者与作者意见不一致
          │
          ▼
   ┌──────────────┐
   │  双方讨论     │  ← 在 PR 评论中公开讨论，不要私聊
   └──────┬───────┘
          │ 未达成一致
          ▼
   ┌──────────────┐
   │  技术负责人   │  ← @技术负责人 介入裁决
   │  介入裁决     │
   └──────┬───────┘
          │ 仍未解决
          ▼
   ┌──────────────┐
   │  团队会议     │  ← 罕见情况，团队讨论决定
   │  讨论决定     │
   └──────────────┘
```

**争议原则**：
- 以数据和安全为准，不以个人偏好为准
- 争议应在 PR 内公开讨论，避免私聊决策
- 技术负责人裁决后，双方需执行

---

## 四、自动化工具链

### 4.1 现状分析

| 工具 | 当前状态 | 审查口径 |
|------|----------|----------|
| ESLint | 已配置并进入桌面 CI | 任何 error 阻断 |
| Prettier | 已配置，`format:check` 未进入 CI | 不得把未执行的格式检查描述为发布门禁 |
| TypeScript | Renderer 与 Electron 主进程均进入 CI | 任何类型错误阻断 |
| Vitest | 已配置并进入桌面 CI | 数量以本次运行输出为准 |
| .NET build/test/audit | 固定 SDK，进入桌面 CI | 漏洞、构建或测试失败阻断 |
| 产品站 Node test/audit | 独立 CI job | 漏洞或测试失败阻断 |
| Coverage | 脚本存在，但 provider 依赖未声明 | 当前不可作为可用门禁 |
| Husky / lint-staged / CommitLint | 未安装 | 只能列为候选改进，不得写成现行流程 |

### 4.2 当前工具链架构

```
开发者本地                              CI / GitHub Actions
┌─────────────────────────────┐        ┌──────────────────────────────┐
│ 按变更范围执行               │        │ desktop                     │
│ lint / typecheck / test      │        │ npm + NuGet audit           │
│ office:test / build          │        │ lint / typecheck / Vitest   │
│ product-site test / audit    │        │ .NET test / Vite build      │
│ 真实 Office 冒烟按需执行      │        ├──────────────────────────────┤
└─────────────────────────────┘        │ product-site audit / test    │
                                       └──────────────────────────────┘
```

本仓库当前没有自动 Git hooks。评审只能要求实际存在的脚本和 CI job；拟引入的工具必须先落地依赖、配置、锁文件和防回归测试，再升级为门禁。

### 4.3 ESLint 收紧候选（非现行配置）

实际配置见 `desktop/eslint.config.js`。以下片段只表示未来可评估的规则方向，不能用于判断当前依赖、插件或门禁已经存在。

```jsonc
// desktop/eslint.config.js (Flat Config 格式, ESLint 9+)
import js from "@eslint/js";
import tseslint from "typescript-eslint";
import react from "eslint-plugin-react";
import reactHooks from "eslint-plugin-react-hooks";
import importPlugin from "eslint-plugin-import";

export default [
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    files: ["src/**/*.{ts,tsx}"],
    ...react.configs.recommended,
    plugins: { "react-hooks": reactHooks },
    rules: {
      ...reactHooks.configs.recommended.rules,
      "react/react-in-jsx-scope": "off", // React 17+ 不需要
    },
  },
  {
    // 全局规则
    rules: {
      // P0 级别：安全相关
      "no-eval": "error",                    // 禁止 eval
      "no-implied-eval": "error",            // 禁止隐式 eval
      "no-new-func": "error",               // 禁止 new Function()
      "@typescript-eslint/no-explicit-any": "warn", // 限制 any

      // P1 级别：代码质量
      "no-console": ["warn", { allow: ["warn", "error"] }], // 限制 console.log
      "no-unused-vars": "off",               // 由 TS 版本接管
      "@typescript-eslint/no-unused-vars": ["error", {
        argsIgnorePattern: "^_",             // 以 _ 开头的参数不检查
      }],
      "prefer-const": "error",               // 能用 const 不用 let
      "no-var": "error",                     // 禁止 var

      // 导入规范
      "import/order": ["warn", {
        "newlines-between": "always",
        groups: ["builtin", "external", "Internal", "parent", "sibling", "index"],
        alphabetize: { order: "asc" },
      }],
    },
  },
  {
    // 测试文件放宽
    files: ["**/*.test.ts", "**/*.test.tsx"],
    rules: {
      "@typescript-eslint/no-explicit-any": "off",
    },
  },
  {
    // 忽略文件
    ignores: ["dist/", "dist-electron/", "node_modules/", "public/"],
  },
];
```

**安装依赖**：
```bash
npm install -D eslint @eslint/js typescript-eslint eslint-plugin-react eslint-plugin-react-hooks eslint-plugin-import
```

### 4.4 Prettier 当前配置与限制

实际配置见 `desktop/.prettierrc` 与 `.prettierignore`。`npm run format:check` 当前未进入 CI，且存量文件仍有格式债务；格式调整应独立分批提交。

```jsonc
// desktop/.prettierrc
{
  "printWidth": 100,
  "semi": true,
  "singleQuote": false,
  "trailingComma": "all"
}
```

```ini
# desktop/.prettierignore
dist/
dist-electron/
node_modules/
release/
package-lock.json
```

Prettier 已作为桌面端开发依赖安装；不要在审查文档中列出仓库未声明的 `eslint-config-prettier` 作为既有依赖。

### 4.5 EditorConfig 候选（尚未启用）

仓库当前没有 `.editorconfig`。以下内容只能作为未来提案；正式启用时需先评估 Windows/Unix 换行基线，避免一次提交重写全仓库。

```ini
# .editorconfig (项目根目录)
root = true

[*]
charset = utf-8
end_of_line = lf
indent_style = space
indent_size = 2
insert_final_newline = true
trim_trailing_whitespace = true

[*.md]
trim_trailing_whitespace = false
```

### 4.6 Husky + lint-staged 候选（尚未启用）

当前 `package.json` 没有对应依赖、脚本或 hooks。以下示例不得作为提交者已经受本地 hook 约束的证据。

```jsonc
// desktop/package.json 中添加
{
  "scripts": {
    "prepare": "husky"
  },
  "lint-staged": {
    "*.{ts,tsx}": ["eslint --fix", "prettier --write"],
    "*.{json,css,md}": ["prettier --write"]
  }
}
```

```bash
# pre-commit 钩子
npx husky add desktop/.husky/pre-commit "npx lint-staged"

# commit-msg 钩子
npx husky add desktop/.husky/commit-msg "npx --no -- commitlint --edit ${1}"
```

**安装依赖**：
```bash
npm install -D husky lint-staged
```

### 4.7 CommitLint 候选（尚未启用）

当前仓库没有 CommitLint 配置或依赖；提交消息规范依靠人工审查。

```jsonc
// desktop/commitlint.config.js
export default {
  extends: ["@commitlint/config-conventional"],
  rules: {
    "type-enum": [
      2,
      "always",
      [
        "feat",     // 新功能
        "fix",      // Bug 修复
        "refactor", // 重构
        "perf",     // 性能优化
        "test",     // 测试
        "docs",     // 文档
        "style",    // 格式调整（不影响逻辑）
        "chore",    // 构建/工具变更
        "ci",       // CI 配置
        "revert",   // 回滚
      ],
    ],
    "subject-max-length": [2, "always", 72],
    "subject-case": [0], // 关闭大小写检查（中文不受限）
  },
};
```

**Commit Message 格式**：
```
<type>(<scope>): <subject>

<body>

<footer>
```

**示例**：
```
feat(agent): 新增 Excel 图表操作工具

- 支持 bar/line/pie 三种图表类型
- 通过 COM Bridge 调用 Excel 原生图表 API
- 添加 12 个单元测试覆盖各图表类型

Closes #42
```

**安装依赖**：
```bash
npm install -D @commitlint/cli @commitlint/config-conventional
```

### 4.8 Coverage 候选（尚未启用）

`desktop/vitest.config.ts` 保留 coverage 配置和 `test:coverage` 脚本，但 `@vitest/coverage-v8` 尚未声明，因此当前 CI 不运行 coverage。若未来启用，必须先补依赖、锁文件、可重复基线和误差处理，再设置阈值；以下仅为候选示例：

```typescript
// vitest.config.ts 增加 coverage 阈值
export default defineConfig({
  test: {
    coverage: {
      provider: "v8",
      reporter: ["text", "json", "html"],
      include: [
        "src/utils/**",
        "src/store/**",
        "electron/agent/core/agentLoop/**",
        "electron/agent/compaction.ts",
      ],
      // 阶梯式提升，先设低门槛，逐步收紧
      thresholds: {
        statements: 60,   // 初始 60%，每季度提升 5%
        branches: 55,
        functions: 60,
        lines: 60,
      },
    },
  },
});
```

---

## 五、CI 门禁规则

### 5.1 当前 CI 流水线

```yaml
# 摘要；完整事实以 .github/workflows/ci.yml 为准
name: CI
on:
  push:
  pull_request:
  workflow_call:

jobs:
  desktop:
    runs-on: windows-latest
    defaults:
      run:
        working-directory: desktop
    steps:
      - uses: actions/checkout@<pinned-commit>
      - uses: actions/setup-node@<pinned-commit>
        with:
          node-version: "22"
      - uses: actions/setup-dotnet@<pinned-commit>
        with:
          dotnet-version: "8.0.422"
      - run: npm ci
      - run: npm audit --audit-level=high
      - run: npm run office:audit
      - run: npm run lint
      - run: npm run typecheck
      - run: npm test
      - run: npm run office:test
      - run: npm run build

  product-site:
    runs-on: ubuntu-latest
    defaults:
      run:
        working-directory: product-site
    steps:
      - uses: actions/checkout@<pinned-commit>
      - uses: actions/setup-node@<pinned-commit>
        with:
          node-version: "22"
      - run: npm ci
      - run: npm audit --audit-level=high
      - run: npm test
```

### 5.2 门禁规则汇总

| 门禁 | 范围 | 阻断条件 |
|------|------|----------|
| npm audit | desktop + product-site | High 及以上漏洞 |
| NuGet audit | .NET Worker 与测试项目 | 已知漏洞包 |
| ESLint | desktop | 任何 error |
| TypeScript | Renderer + Electron | 任何类型错误 |
| Vitest | desktop | 任何测试失败 |
| .NET test | Worker | 构建或测试失败 |
| Vite build | desktop | 生产构建失败 |
| Node test | product-site | 任何测试失败 |

`format:check`、coverage、真实 Office/WPS 冒烟和 NSIS 打包不是当前 `ci.yml` 门禁；Release 工作流和本地发布验收有各自的附加步骤，不得混写。

### 5.3 门禁与人工审查的分工

```
┌─────────────────────────────────────────────────────┐
│                    CI 自动门禁                        │
│  (机器能查的，不浪费人脑)                              │
│  ┌────────┬─────────┬──────────┬────────┬─────────┐ │
│  │ Audit  │ ESLint  │ TypeCheck │ Tests  │ Build   │ │
│  └────────┴─────────┴──────────┴────────┴─────────┘ │
│  覆盖：依赖漏洞、语法、类型、已有测试、生产构建           │
└────────────────────┬────────────────────────────────┘
                     │ CI 通过后才进入人工审查
                     ▼
┌─────────────────────────────────────────────────────┐
│                   人工审查                            │
│  (机器查不了的，才需要人脑)                            │
│  ┌──────────┬──────────┬──────────┬───────────────┐ │
│  │ 逻辑正确性 │ 安全设计  │ 架构合理性 │ 业务意图匹配  │ │
│  └──────────┴──────────┴──────────┴───────────────┘ │
│  覆盖：C1-C3, S1-S5, M1-M2, 业务逻辑                 │
└─────────────────────────────────────────────────────┘
```

**原则**：CI 负责机械性检查（格式、语法、类型、测试通过），人工审查聚焦于机器无法判断的维度（逻辑正确性、架构合理性、业务意图）。

---

## 六、审查模板

### 6.1 PR 描述模板

在 GitHub 仓库设置中添加 PR 模板（`.github/pull_request_template.md`）：

```markdown
## 变更说明

<!-- 一句话描述这个 PR 做了什么 -->

## 变更类型

- [ ] feat — 新功能
- [ ] fix — Bug 修复
- [ ] refactor — 重构（不改行为）
- [ ] perf — 性能优化
- [ ] test — 测试
- [ ] docs — 文档
- [ ] chore — 构建/工具

## 关联 Issue

Closes #

## 自查清单

- [ ] 新文件 ≤ 400 行，组件 ≤ 300 行
- [ ] 组件 props ≤ 10 个，或使用 hook 归组
- [ ] IPC 调用使用 ipcApi，非 window.electronAPI
- [ ] 新增 IPC 通道有 zod schema 校验
- [ ] 核心逻辑有对应 .test.ts 文件
- [ ] 无 console.log，使用 logger
- [ ] npm run typecheck 通过
- [ ] npm test 全部通过
- [ ] CHANGELOG.md 已更新

## 测试说明

<!-- 描述如何测试本次变更 -->

## 截图 / 录屏

<!-- 如果涉及 UI 变更，附上截图 -->

## 审查重点

<!-- 告诉审查者需要特别关注的地方 -->
```

### 6.2 审查评论模板

审查者在 PR 中发表评论时，使用以下格式：

```markdown
<!-- P0 阻断问题 -->
🔴 **[类型]: [简短描述]**
**位置**：`文件路径:行号`
**问题**：[具体描述问题是什么]
**原因**：[为什么这是问题，会导致什么后果]
**建议**：
```typescript
// 建议的修改方式
```

<!-- P1 建议改进 -->
🟡 **[类型]: [简短描述]**
**位置**：`文件路径:行号`
**问题**：[具体描述]
**建议**：[改进方向]

<!-- P2 微调 -->
💭 **[位置]**：[一句话建议]

<!-- 亮点表扬 -->
✅ **亮点**：[具体的好的实践，说明为什么好]
```

### 6.3 审查总结模板

审查完成后，审查者在 PR 底部发表总结评论：

```markdown
## 审查总结

**整体评价**：[1-2 句话整体印象]

### 统计

| 级别 | 数量 |
|------|------|
| 🔴 P0 | X |
| 🟡 P1 | X |
| 💭 P2 | X |
| ✅ 亮点 | X |

### 主要关注点

1. [最需要关注的问题 1]
2. [最需要关注的问题 2]

### 做得好的地方

- [值得学习的实践 1]
- [值得学习的实践 2]

### 结论

- [ ] Approve — 可以合并
- [ ] Request Changes — 需修改后重新审查
- [ ] Comment — 有疑问需讨论
```

---

## 七、审查文化

### 7.1 审查者守则

| 原则 | 说明 | 反面示例 |
|------|------|----------|
| **对事不对人** | 评论代码，不评论人 | ❌ "你这里写错了" → ✅ "这段逻辑可能有问题" |
| **解释为什么** | 不只说改什么，说为什么改 | ❌ "改用 map" → ✅ "用 map 更清晰，因为..." |
| **建议而非命令** | 用"考虑""建议"而非"必须""改为" | ❌ "改为参数化查询" → ✅ "考虑使用参数化查询，因为..." |
| **给出具体方案** | 不只指出问题，给出可操作建议 | ❌ "这里不安全" → ✅ "这里存在注入风险，建议用 zod 校验输入" |
| **表扬好代码** | 看到好实践主动点赞 | ✅ "这个类型守卫用得很巧妙，学习了" |
| **控制审查范围** | 不在 P2 上纠缠，不挑风格之争 | 不在 PR 中争论 tab vs space（交给 Prettier） |
| **及时响应** | 4 小时内首次响应，1 天内完成审查 | — |

### 7.2 被审查者守则

| 原则 | 说明 |
|------|------|
| **CI 先行** | 提交 PR 前确保本地 typecheck + test 通过 |
| **自查清单** | 完成 [提交前自查清单](#33-pr-提交前自查) 再提交 |
| **小步提交** | 一个 PR 只做一件事，控制在 500 行以内。大功能拆成多个 PR |
| **描述清晰** | PR 描述说清楚"做了什么""为什么做""怎么测试" |
| **虚心接受** | 审查意见是帮助改进代码，不是否定个人能力 |
| **有理有据** | 不同意审查意见时，给出技术理由而非情绪回应 |
| **及时修改** | 收到反馈后 1 个工作日内响应 |

### 7.3 审查不做什么

```
┌─────────────────────────────────────────────┐
│              代码审查的红线                   │
├─────────────────────────────────────────────┤
│  ❌ 不审查个人风格偏好（交给 Prettier）       │
│  ❌ 不在 PR 中争论命名（除非真的有歧义）       │
│  ❌ 不要求"我会怎么写"（尊重不同实现方式）     │
│  ❌ 不堆积几十条 nit 让作者淹没在微调中        │
│  ❌ 不在不理解上下文的情况下否定设计决策       │
│  ❌ 不用审查权力强加个人架构偏好              │
│  ❌ 不在公开场合批评作者                      │
│  ❌ 不拖延审查（PR 放着不审是团队阻塞）        │
└─────────────────────────────────────────────┘
```

### 7.4 持续改进机制

| 频率 | 活动 | 内容 |
|------|------|------|
| 每周 | 审查回顾 | 团队站会中花 5 分钟回顾本周审查中的典型问题和好实践 |
| 每月 | 标准校准 | 审查标准是否需要调整？有无新的常见问题需要加入 checklist？ |
| 每季度 | 工具链升级 | ESLint 规则收紧、覆盖率阈值提升、新增检查工具 |
| 按需 | 审查培训 | 新成员入职时进行审查标准培训 |

---

## 八、附录

### 8.1 审查标准速查卡

```
┌─────────────────────────────────────────────────────────┐
│                  代码审查速查卡                          │
├──────────┬──────────────────────────────────────────────┤
│ 🔴 P0    │ 安全漏洞 · 数据丢失 · 崩溃 · API 契约破坏    │
│ 必须修复  │ 类型不安全 · IPC 未校验 · 命令注入           │
│          │ 文件超限 · 分层违规 · 测试失败                │
├──────────┼──────────────────────────────────────────────┤
│ 🟡 P1    │ 输入校验缺失 · 命名不清 · 缺少测试            │
│ 应该修复  │ N+1 查询 · 全量持久化 · 内存泄漏              │
│          │ Props 过多 · IPC 直调 · console.log           │
├──────────┼──────────────────────────────────────────────┤
│ 💭 P2    │ 代码重复 · 魔法数字 · 文档缺失                │
│ 酌情处理  │ 不必要分配 · 防抖优化                         │
├──────────┼──────────────────────────────────────────────┤
│ ✅ 亮点   │ 类型守卫巧妙 · 防御性编程到位 · 函数职责清晰  │
│          │ 错误处理完善 · 测试覆盖全面                    │
└──────────┴──────────────────────────────────────────────┘
```

### 8.2 候选工具评估清单

以下命令不是当前项目初始化步骤。只有在独立变更中完成兼容性评估、锁文件更新、全量门禁和存量格式影响检查后，才可选择性引入；禁止为“对齐文档”直接批量安装。

```bash
# 进入 desktop 目录
cd desktop

# 一次性安装所有开发依赖
npm install -D eslint @eslint/js typescript-eslint \
  eslint-plugin-react eslint-plugin-react-hooks eslint-plugin-import \
  prettier eslint-config-prettier \
  husky lint-staged \
  @commitlint/cli @commitlint/config-conventional

# 初始化 husky
npx husky init

# 创建 pre-commit 钩子
echo "npx lint-staged" > .husky/pre-commit

# 创建 commit-msg 钩子
echo "npx --no -- commitlint --edit \$1" > .husky/commit-msg
```

### 8.3 常见问题

**Q: PR 太大怎么办？**
A: 拆分。一个 PR 只做一件事。如果功能确实大，按子功能拆成多个 PR，逐个合入。

**Q: 审查者和作者对架构方案有分歧怎么办？**
A: 在 PR 中公开讨论 → 技术负责人裁决 → 团队会议（罕见）。以数据和安全性为准，不以个人偏好为准。

**Q: 紧急修复来不及走审查流程怎么办？**
A: 使用 [Hotfix 快速通道](#36-紧急通道hotfix)。技术负责人授权后直接合入，24 小时内补审。

**Q: ESLint 报错太多怎么办？**
A: 分阶段引入。第一周只开 `error` 级规则，`warn` 级暂不阻断。存量问题创建技术债务 Issue 逐步清理。

**Q: 为什么当前没有覆盖率阈值？**
A: coverage provider 依赖尚未落地，仓库也没有经验证的稳定基线。应先让 coverage 在本地与 CI 可重复运行，再基于真实结果制定渐进阈值，不能假定当前已经达到某个百分比。

---

> **文档维护**：本标准随项目演进持续更新。每季度审查一次适用性，如有新框架/新工具引入需同步更新 checklist。
>
> **反馈渠道**：对本标准有任何建议，请创建 Issue 标注 `type: docs` 并 @技术负责人。
