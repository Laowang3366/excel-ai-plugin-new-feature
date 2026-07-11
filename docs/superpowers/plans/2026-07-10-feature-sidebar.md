# 右侧功能模块侧栏 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 删除紫色悬浮球、可拖拽任务浮窗和 Office 文件编辑监控，将六个任务模块迁移到可平滑收起的右侧功能侧栏。

**Architecture:** 用独立纯 reducer 管理侧栏的打开、选择和关闭状态，`ChatPage` 本地持有该状态并继续复用现有任务表单和消息发送链路。`FeatureSidebarPanel` 只渲染标题、模块网格、表单容器和空状态；任务草稿继续按会话键保存，关闭侧栏不会修改草稿。

**Tech Stack:** React 18、TypeScript、Zustand、Vitest、Vite、Electron、CSS

---

## 文件结构

**新增：**

- `desktop/src/utils/featureSidebarState.ts`：功能侧栏状态与纯 reducer。
- `desktop/src/utils/featureSidebarState.test.ts`：打开、选择、关闭状态测试。
- `desktop/src/components/common/FeatureSidebarPanel.tsx`：右侧功能侧栏展示组件。
- `desktop/src/components/common/FeatureSidebarPanel.test.ts`：模块列表、激活态和收起态静态渲染测试。
- `desktop/src/styles/feature-sidebar-panel.css`：侧栏尺寸、滚动、响应式和动画。

**修改：**

- `desktop/src/App.tsx`：删除跨层 `activeIntent` 状态和属性传递。
- `desktop/src/components/ChatPage.tsx`：持有侧栏状态、渲染任务表单、删除文件监控。
- `desktop/src/hooks/useTaskDrafts.ts`：只管理草稿，不再负责关闭面板或清空草稿。
- `desktop/src/i18n.ts`：增加功能侧栏标题、空状态和开关文案。
- `desktop/src/utils/chatHelpers.tsx`：删除仅供旧浮窗使用的元数据与坐标限制函数。
- `desktop/src/styles/global.css`：换成新侧栏样式入口。
- `desktop/src/styles/sidebar.css`：删除已无调用方的旧功能快捷区样式。

**删除：**

- `desktop/src/components/common/FeatureFloatingDock.tsx`
- `desktop/src/components/common/FeatureFloatingDock.test.ts`
- `desktop/src/components/common/featureFloatingDockGeometry.ts`
- `desktop/src/components/common/FloatingTaskPanel.tsx`
- `desktop/src/components/office/OfficePreviewPanel.tsx`
- `desktop/src/utils/officeEditEvents.ts`
- `desktop/src/utils/officeEditEvents.test.ts`
- `desktop/src/styles/floating-task-panel.css`
- `desktop/src/styles/office-preview-panel.css`

### Task 1: 建立可测试的侧栏状态

**Files:**

- Create: `desktop/src/utils/featureSidebarState.ts`
- Test: `desktop/src/utils/featureSidebarState.test.ts`

- [ ] **Step 1: 写状态 reducer 的失败测试**

```ts
import { describe, expect, it } from "vitest";
import {
  INITIAL_FEATURE_SIDEBAR_STATE,
  reduceFeatureSidebarState,
} from "./featureSidebarState";

describe("reduceFeatureSidebarState", () => {
  it("opens the sidebar without selecting a feature", () => {
    expect(reduceFeatureSidebarState(INITIAL_FEATURE_SIDEBAR_STATE, { type: "toggle" })).toEqual({
      isOpen: true,
      activeIntent: null,
    });
  });

  it("selects a feature and keeps the sidebar open", () => {
    expect(reduceFeatureSidebarState(INITIAL_FEATURE_SIDEBAR_STATE, {
      type: "select",
      intent: "formula",
    })).toEqual({
      isOpen: true,
      activeIntent: "formula",
    });
  });

  it("closes the sidebar and clears only the active selection", () => {
    expect(reduceFeatureSidebarState({
      isOpen: true,
      activeIntent: "chart",
    }, { type: "close" })).toEqual(INITIAL_FEATURE_SIDEBAR_STATE);
  });

  it("uses the header toggle to close an open sidebar", () => {
    expect(reduceFeatureSidebarState({
      isOpen: true,
      activeIntent: "clean",
    }, { type: "toggle" })).toEqual(INITIAL_FEATURE_SIDEBAR_STATE);
  });
});
```

- [ ] **Step 2: 运行测试并确认失败**

Run: `npm test -- src/utils/featureSidebarState.test.ts`

Expected: FAIL，提示找不到 `./featureSidebarState`。

- [ ] **Step 3: 实现最小状态 reducer**

```ts
import type { IntentKind } from "./sidebarHelpers";

export interface FeatureSidebarState {
  isOpen: boolean;
  activeIntent: IntentKind;
}

export type FeatureSidebarAction =
  | { type: "toggle" }
  | { type: "select"; intent: NonNullable<IntentKind> }
  | { type: "close" };

export const INITIAL_FEATURE_SIDEBAR_STATE: FeatureSidebarState = {
  isOpen: false,
  activeIntent: null,
};

export function reduceFeatureSidebarState(
  state: FeatureSidebarState,
  action: FeatureSidebarAction
): FeatureSidebarState {
  if (action.type === "select") {
    return { isOpen: true, activeIntent: action.intent };
  }
  if (action.type === "close" || state.isOpen) {
    return INITIAL_FEATURE_SIDEBAR_STATE;
  }
  return { isOpen: true, activeIntent: null };
}
```

- [ ] **Step 4: 运行测试并确认通过**

Run: `npm test -- src/utils/featureSidebarState.test.ts`

Expected: PASS，4 个状态用例全部通过。

- [ ] **Step 5: 提交状态基础**

```powershell
git add desktop/src/utils/featureSidebarState.ts desktop/src/utils/featureSidebarState.test.ts
git commit -m "feat: 添加功能侧栏状态管理"
```

### Task 2: 新增右侧功能侧栏组件

**Files:**

- Create: `desktop/src/components/common/FeatureSidebarPanel.tsx`
- Test: `desktop/src/components/common/FeatureSidebarPanel.test.ts`
- Modify: `desktop/src/i18n.ts`

- [ ] **Step 1: 写组件静态渲染失败测试**

```ts
import React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it, vi } from "vitest";
import { FeatureSidebarPanel } from "./FeatureSidebarPanel";

describe("FeatureSidebarPanel", () => {
  it("renders all six feature shortcuts and the selected state", () => {
    const html = renderToStaticMarkup(React.createElement(
      FeatureSidebarPanel,
      {
        isOpen: true,
        activeIntent: "ocr",
        language: "zh-CN",
        onIntentClick: vi.fn(),
        onClose: vi.fn(),
      },
      React.createElement("div", null, "OCR form"),
    ));

    expect(html).toContain("公式助手");
    expect(html).toContain("代码生成");
    expect(html).toContain("OCR 识别");
    expect(html).toContain("数据清洗");
    expect(html).toContain("报告生成");
    expect(html).toContain("图表制作");
    expect(html).toContain('aria-selected="true"');
    expect(html).toContain("OCR form");
  });

  it("renders a closed, inaccessible panel without an active form", () => {
    const html = renderToStaticMarkup(React.createElement(FeatureSidebarPanel, {
      isOpen: false,
      activeIntent: null,
      language: "zh-CN",
      onIntentClick: vi.fn(),
      onClose: vi.fn(),
    }));

    expect(html).toContain('aria-hidden="true"');
    expect(html).toContain("选择上方功能开始配置");
    expect(html).not.toContain('aria-selected="true"');
  });
});
```

- [ ] **Step 2: 运行测试并确认失败**

Run: `npm test -- src/components/common/FeatureSidebarPanel.test.ts`

Expected: FAIL，提示找不到 `./FeatureSidebarPanel`。

- [ ] **Step 3: 增加中英文侧栏文案**

在 `APP_TEXT["zh-CN"]` 和 `APP_TEXT["en-US"]` 的 `chat` 节点中分别加入：

```ts
featureSidebar: {
  title: "功能模块",
  open: "打开功能模块",
  close: "关闭功能模块",
  empty: "选择上方功能开始配置",
},
```

```ts
featureSidebar: {
  title: "Features",
  open: "Open features",
  close: "Close features",
  empty: "Choose a feature above to configure",
},
```

- [ ] **Step 4: 实现展示组件**

```tsx
import React from "react";
import type { AppLanguage } from "../../store/settingsStore";
import type { IntentKind } from "../../utils/sidebarHelpers";
import { INTENT_SHORTCUTS } from "../../utils/sidebarHelpers";
import { getAppText } from "../../i18n";
import { Sparkles, X } from "./IconMap";

interface FeatureSidebarPanelProps {
  isOpen: boolean;
  activeIntent: IntentKind;
  language: AppLanguage;
  onIntentClick: (intent: NonNullable<IntentKind>) => void;
  onClose: () => void;
  children?: React.ReactNode;
}

export function FeatureSidebarPanel({
  isOpen,
  activeIntent,
  language,
  onIntentClick,
  onClose,
  children,
}: FeatureSidebarPanelProps) {
  const text = getAppText(language);

  return (
    <aside
      className={`feature-sidebar-panel ${isOpen ? "open" : "collapsed"}`}
      aria-hidden={!isOpen}
      inert={isOpen ? undefined : true}
    >
      <div className="feature-sidebar-content">
        <div className="feature-sidebar-header">
          <div className="feature-sidebar-title">
            <Sparkles size={16} />
            <span>{text.chat.featureSidebar.title}</span>
          </div>
          <button
            className="feature-sidebar-close"
            type="button"
            onClick={onClose}
            title={text.chat.featureSidebar.close}
            aria-label={text.chat.featureSidebar.close}
          >
            <X size={16} />
          </button>
        </div>

        <div className="feature-sidebar-shortcuts" role="listbox">
          {INTENT_SHORTCUTS.map((shortcut) => {
            const selected = activeIntent === shortcut.key;
            return (
              <button
                key={shortcut.key}
                className={`feature-sidebar-shortcut ${selected ? "active" : ""}`}
                type="button"
                role="option"
                aria-selected={selected}
                onClick={() => onIntentClick(shortcut.key)}
              >
                <shortcut.icon size={18} />
                <span>{text.sidebar.intents[shortcut.key]}</span>
              </button>
            );
          })}
        </div>

        <div className="feature-sidebar-form">
          {activeIntent ? children : (
            <div className="feature-sidebar-empty">
              <Sparkles size={20} />
              <span>{text.chat.featureSidebar.empty}</span>
            </div>
          )}
        </div>
      </div>
    </aside>
  );
}
```

- [ ] **Step 5: 运行组件测试**

Run: `npm test -- src/components/common/FeatureSidebarPanel.test.ts`

Expected: PASS，六个模块、激活态、关闭态和空状态全部通过。

- [ ] **Step 6: 提交侧栏组件**

```powershell
git add desktop/src/components/common/FeatureSidebarPanel.tsx desktop/src/components/common/FeatureSidebarPanel.test.ts desktop/src/i18n.ts
git commit -m "feat: 添加右侧功能模块侧栏"
```

### Task 3: 将任务表单迁移到 ChatPage

**Files:**

- Modify: `desktop/src/App.tsx`
- Modify: `desktop/src/components/ChatPage.tsx`
- Modify: `desktop/src/hooks/useTaskDrafts.ts`
- Test: `desktop/src/hooks/useTaskDrafts.test.ts`

- [ ] **Step 1: 扩充草稿隔离测试**

在 `useTaskDrafts.test.ts` 增加：

```ts
it("keeps an existing feature draft when unrelated UI state closes", () => {
  const store: TaskDraftStore = {
    "thread-1": {
      formula: {
        dataSourceRanges: ["Sheet1!A1:B5"],
        dataSourceInput: "",
        referenceSampleRange: "",
        referenceSampleMode: "partial",
        outputRange: "Sheet1!C1",
        hostEnvironment: "microsoft_excel",
        task: "计算同比",
      },
    },
  };

  expect(getTaskDraftsForKey(store, "thread-1").formula?.task).toBe("计算同比");
});
```

- [ ] **Step 2: 运行草稿测试作为基线**

Run: `npm test -- src/hooks/useTaskDrafts.test.ts`

Expected: PASS，草稿仍按会话键隔离并可保留。

- [ ] **Step 3: 让草稿 hook 只管理草稿**

将签名改为：

```ts
export function useTaskDrafts(draftKey = "default") {
```

删除 `activeIntent`、`onIntentClick` 参数，删除 `closeActiveTaskPanel` 回调及返回字段。更新文件头说明为“关闭侧栏不清空草稿，侧栏状态由 ChatPage 管理”。

- [ ] **Step 4: 把侧栏状态移入 ChatPage**

`ChatPage` 使用：

```tsx
const [featureSidebar, dispatchFeatureSidebar] = useReducer(
  reduceFeatureSidebarState,
  INITIAL_FEATURE_SIDEBAR_STATE,
);
const { isOpen: featureSidebarOpen, activeIntent } = featureSidebar;

const closeFeatureSidebar = useCallback(() => {
  dispatchFeatureSidebar({ type: "close" });
}, []);

const selectFeature = useCallback((intent: NonNullable<IntentKind>) => {
  dispatchFeatureSidebar({ type: "select", intent });
}, []);
```

标题栏右侧使用：

```tsx
<button
  className={`feature-sidebar-toggle ${featureSidebarOpen ? "active" : ""}`}
  type="button"
  onClick={() => dispatchFeatureSidebar({ type: "toggle" })}
  title={featureSidebarOpen
    ? text.chat.featureSidebar.close
    : text.chat.featureSidebar.open}
  aria-label={featureSidebarOpen
    ? text.chat.featureSidebar.close
    : text.chat.featureSidebar.open}
  aria-pressed={featureSidebarOpen}
>
  <Sparkles size={15} />
</button>
```

`FeatureSidebarPanel` 放在 `.chat-workspace` 之后，并把原六个表单原样放进其 children。所有原 `closeActiveTaskPanel` 改为 `closeFeatureSidebar`，`useTaskDrafts` 改为：

```ts
const {
  taskDrafts,
  setTaskDrafts,
  updateFormulaDraft,
  updateCodeDraft,
  updateOCRDraft,
  updateReportDraft,
  handleSimplePickRange,
} = useTaskDrafts(composerDraftKey);
```

普通发送和任务提交都在进入现有发送链路后执行：

```ts
closeFeatureSidebar();
```

- [ ] **Step 5: 删除 App 的跨层意图状态**

删除：

```ts
const [activeIntent, setActiveIntent] = useState<IntentKind>(null);
```

并把 `ChatPage` 调用改为：

```tsx
<ChatPage
  onOpenSettings={(section = "general") => {
    setSettingsSection(section);
    setCurrentPage("settings");
  }}
/>
```

- [ ] **Step 6: 运行状态、草稿和类型检查**

Run: `npm test -- src/utils/featureSidebarState.test.ts src/hooks/useTaskDrafts.test.ts`

Expected: PASS。

Run: `npm run typecheck`

Expected: PASS，不再存在 `activeIntent` 属性或 hook 参数错误。

- [ ] **Step 7: 提交页面集成**

```powershell
git add desktop/src/App.tsx desktop/src/components/ChatPage.tsx desktop/src/hooks/useTaskDrafts.ts desktop/src/hooks/useTaskDrafts.test.ts
git commit -m "feat: 迁移任务模块到右侧侧栏"
```

### Task 4: 实现平滑收缩样式

**Files:**

- Create: `desktop/src/styles/feature-sidebar-panel.css`
- Modify: `desktop/src/styles/global.css`

- [ ] **Step 1: 新增侧栏样式**

```css
.feature-sidebar-toggle,
.feature-sidebar-close {
  width: 30px;
  height: 30px;
  flex: 0 0 auto;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  border: 1px solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--bg-primary);
  color: var(--text-faint);
  cursor: pointer;
  transition: background 150ms ease, color 150ms ease, border-color 150ms ease;
}

.feature-sidebar-toggle:hover,
.feature-sidebar-close:hover {
  background: var(--bg-tint);
  color: var(--text-primary);
}

.feature-sidebar-toggle.active {
  border-color: var(--primary-border);
  background: var(--primary-light);
  color: var(--primary);
}

.feature-sidebar-panel {
  --feature-sidebar-width: clamp(320px, 30vw, 360px);
  width: 0;
  min-width: 0;
  flex: 0 0 0;
  height: 100%;
  overflow: hidden;
  border-left: 1px solid transparent;
  background: var(--bg-primary);
  opacity: 0;
  pointer-events: none;
  transition:
    width 180ms cubic-bezier(0.22, 1, 0.36, 1),
    flex-basis 180ms cubic-bezier(0.22, 1, 0.36, 1),
    opacity 120ms ease,
    border-color 140ms ease;
}

.feature-sidebar-panel.open {
  width: var(--feature-sidebar-width);
  min-width: 320px;
  flex-basis: var(--feature-sidebar-width);
  border-left-color: var(--border-default);
  opacity: 1;
  pointer-events: auto;
}

.feature-sidebar-content {
  width: var(--feature-sidebar-width);
  min-width: var(--feature-sidebar-width);
  height: 100%;
  display: flex;
  flex-direction: column;
  opacity: 0;
  transform: translateX(8px);
  transition:
    opacity 120ms ease,
    transform 180ms cubic-bezier(0.22, 1, 0.36, 1);
}

.feature-sidebar-panel.open .feature-sidebar-content {
  opacity: 1;
  transform: translateX(0);
  transition-delay: 20ms;
}

.feature-sidebar-header {
  height: var(--chat-header-height);
  box-sizing: border-box;
  flex: 0 0 var(--chat-header-height);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 0 12px 0 14px;
  border-bottom: 1px solid var(--border-default);
}

.feature-sidebar-title {
  min-width: 0;
  display: inline-flex;
  align-items: center;
  gap: 7px;
  color: var(--text-primary);
  font-size: 13px;
  font-weight: 700;
}

.feature-sidebar-shortcuts {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 8px;
  padding: 12px;
  border-bottom: 1px solid var(--border-default);
}

.feature-sidebar-shortcut {
  min-width: 0;
  height: 58px;
  display: flex;
  align-items: center;
  gap: 9px;
  padding: 0 10px;
  border: 1px solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--bg-primary);
  color: var(--text-secondary);
  cursor: pointer;
  font: inherit;
  font-size: 12px;
  font-weight: 600;
  text-align: left;
  transition: background 150ms ease, color 150ms ease, border-color 150ms ease;
}

.feature-sidebar-shortcut span {
  min-width: 0;
  overflow-wrap: anywhere;
}

.feature-sidebar-shortcut:hover {
  border-color: var(--primary-border);
  background: var(--bg-tint);
  color: var(--primary);
}

.feature-sidebar-shortcut.active {
  border-color: var(--primary-border);
  background: var(--primary-light);
  color: var(--primary);
}

.feature-sidebar-form {
  min-height: 0;
  flex: 1;
  overflow-y: auto;
  overflow-x: hidden;
  padding: 12px;
}

.feature-sidebar-form .task-composer-panel {
  margin: 0;
  padding: 0;
  border: none;
  border-radius: 0;
  background: transparent;
  box-shadow: none;
  animation: none;
}

.feature-sidebar-empty {
  min-height: 150px;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 8px;
  color: var(--text-faint);
  font-size: 12px;
  text-align: center;
}

.compact-mode .feature-sidebar-toggle,
.compact-mode .feature-sidebar-panel {
  display: none;
}

@media (max-width: 980px) {
  .feature-sidebar-toggle,
  .feature-sidebar-panel {
    display: none;
  }
}

@media (prefers-reduced-motion: reduce) {
  .feature-sidebar-panel,
  .feature-sidebar-content {
    transition: none;
  }
}
```

- [ ] **Step 2: 更新全局样式入口**

删除：

```css
@import "./floating-task-panel.css";
@import "./office-preview-panel.css";
```

增加：

```css
@import "./feature-sidebar-panel.css";
```

- [ ] **Step 3: 构建验证 CSS**

Run: `npm run build`

Expected: PASS，Vite 成功产出 `dist`，没有缺失样式引用。

- [ ] **Step 4: 提交动画样式**

```powershell
git add desktop/src/styles/feature-sidebar-panel.css desktop/src/styles/global.css
git commit -m "style: 优化功能侧栏展开收缩动画"
```

### Task 5: 删除旧悬浮与文件监控代码

**Files:**

- Delete: `desktop/src/components/common/FeatureFloatingDock.tsx`
- Delete: `desktop/src/components/common/FeatureFloatingDock.test.ts`
- Delete: `desktop/src/components/common/featureFloatingDockGeometry.ts`
- Delete: `desktop/src/components/common/FloatingTaskPanel.tsx`
- Delete: `desktop/src/components/office/OfficePreviewPanel.tsx`
- Delete: `desktop/src/utils/officeEditEvents.ts`
- Delete: `desktop/src/utils/officeEditEvents.test.ts`
- Delete: `desktop/src/styles/floating-task-panel.css`
- Delete: `desktop/src/styles/office-preview-panel.css`
- Modify: `desktop/src/utils/chatHelpers.tsx`
- Modify: `desktop/src/styles/sidebar.css`

- [ ] **Step 1: 删除旧文件**

使用 `apply_patch` 删除上述 9 个文件。

- [ ] **Step 2: 清理旧浮窗工具函数**

从 `chatHelpers.tsx` 删除：

```ts
export type ActiveIntentKind = "formula" | "code" | "ocr" | "clean" | "report" | "chart";

export function getTaskPanelMeta(intent: ActiveIntentKind, language: AppLanguage) {
  const labels = getAppText(language).sidebar.intents;
  switch (intent) {
    case "formula":
      return { title: labels.formula, icon: <Hash size={16} /> };
    case "code":
      return { title: labels.code, icon: <Code size={16} /> };
    case "ocr":
      return { title: labels.ocr, icon: <FileScan size={16} /> };
    case "clean":
      return { title: labels.clean, icon: <Eraser size={16} /> };
    case "report":
      return { title: labels.report, icon: <FileBarChart size={16} /> };
    case "chart":
      return { title: labels.chart, icon: <LineChart size={16} /> };
  }
}

export function clamp(value: number, min: number, max: number) {
  return Math.min(Math.max(value, min), Math.max(min, max));
}
```

同时删除只被这些函数使用的 `Hash`、`Code`、`FileScan`、`Eraser`、`FileBarChart`、`LineChart` 图标导入，并更新文件头说明。

- [ ] **Step 3: 删除无调用方的旧侧边栏快捷样式**

从 `sidebar.css` 删除 `.sidebar-intents`、`.sidebar-intents-title` 和 `.sidebar-intent-item` 整段。

- [ ] **Step 4: 静态搜索确认彻底删除**

Run:

```powershell
git grep -n -E "FeatureFloatingDock|FloatingTaskPanel|OfficePreviewPanel|collectOfficeEditEvents|office-preview|feature-floating|task-floating|getTaskPanelMeta|ActiveIntentKind" -- desktop/src
```

Expected: 命令退出码为 1 且无输出，表示生产代码和测试中不再存在旧引用。

- [ ] **Step 5: 运行全量验证**

Run: `npm test`

Expected: PASS。

Run: `npm run typecheck`

Expected: PASS。

Run: `npm run lint`

Expected: PASS。

Run: `npm run build`

Expected: PASS。

- [ ] **Step 6: 提交旧功能删除**

```powershell
git add desktop/src/components desktop/src/utils desktop/src/styles
git commit -m "refactor: 删除悬浮任务入口和文件监控"
```

### Task 6: Electron 界面验证

**Files:**

- Inspect: `desktop/src/components/ChatPage.tsx`
- Inspect: `desktop/src/styles/feature-sidebar-panel.css`

- [ ] **Step 1: 启动桌面端开发环境**

Run: `npm run dev`

Expected: Vite 启动且 Electron 窗口正常打开。

- [ ] **Step 2: 验证正常宽度**

检查：

- 聊天区不再出现紫色悬浮球。
- 标题栏右上角显示功能按钮。
- 点击按钮后侧栏平滑展开，模块区为两列三行。
- 点击不同模块后显示对应原有完整表单。
- 长表单只在右侧栏内部纵向滚动，不产生横向溢出。
- 收起时表单先淡出并轻微右移，聊天区平滑回弹。

- [ ] **Step 3: 验证草稿和发送**

检查：

- 在一个模块输入草稿，关闭并重新打开后草稿仍在。
- 切换模块后原模块草稿仍在。
- 点击任务提交后消息进入原发送链路，侧栏自动关闭。

- [ ] **Step 4: 验证响应式和主题**

检查：

- 窗口宽度小于等于 `980px` 时功能按钮和侧栏隐藏。
- 紧凑模式下功能按钮和侧栏隐藏。
- 浅色与深色主题下文字、边框、激活态清晰。
- 页面不存在元素重叠或按钮文字溢出。

- [ ] **Step 5: 保存验证截图并检查控制台**

截图保存到禁止提交的 `.workbuddy/screenshots/`，确认控制台无 React、CSS 或运行时错误。

- [ ] **Step 6: 最终工作区检查**

Run: `git status --short`

Expected: 只显示本任务已提交内容之外的既有未跟踪本地数据：

```text
?? .codex/
?? .workbuddy/screenshots/
?? desktop/data/
?? desktop/vite-ui-refactor.log
```
