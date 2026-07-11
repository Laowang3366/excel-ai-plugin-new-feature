# 提示词层

职责：按稳定基础指令、按需场景指令、运行时事实三层组装系统提示词。

## 组装顺序

1. `templates/system/`：稳定基础前缀，所有任务共享，优先保持内容和顺序稳定。
2. `templates/scenarios/`：根据本轮内容和附件按需追加，普通问答不加载。
3. `templates/runtime/`：文件夹、Office 连接、功能开关和时间等运行时事实，始终放在静态规则之后。
4. 长期记忆由 `core/agentLoop/roundStreamParams.ts` 在最后追加。

## 模块

- `systemPrompt.ts`：定义片段注册顺序、场景触发条件和运行时模板变量。
- `promptComposer.ts`：负责确定性排序、按 key 去重、空片段清理和严格变量替换。
- `sections/folderContextPrompt.ts`：把文件夹文件列表渲染进运行时模板。
- `templates/system/*.md`：基础角色、工具边界、安全和质量规则。
- `templates/scenarios/*.md`：公式、OCR、Office 文件操作和通用业务场景。
- `templates/runtime/*.md`：本轮可变事实，不得混入稳定基础模板。
- `compactionPrompt.ts` / `templates/compaction.zh-CN.md`：上下文压缩模板。

新增或修改工具时，优先更新对应场景模板；只有所有任务都必须遵守的规则才进入 `templates/system/`。
