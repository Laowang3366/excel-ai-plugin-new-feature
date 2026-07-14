# 工具注册表层

职责：定义模型可见的工具 schema、风险等级、描述和参数结构。

模块说明：

- `toolDefinitions.ts`: 汇总所有工具定义并生成按名称索引的映射。
- `workbook.ts`: 工作簿检查、打开、创建、保存和切换工具定义。
- `range.ts`: 区域读取、写入、清除和选区工具定义。
- `formula.ts`: 公式上下文工具定义。
- `sheet.ts`: 工作表操作工具定义。
- `macro.ts`: Excel/WPS 内部 VBA/JSA 宏工具定义。
- `ui.ts`: Excel UI 控件和窗体工具定义。
- `file.ts`: 文件读取和写入工具定义。
- `knowledge.ts`: 知识库检索工具定义。
- `ocr.ts`: OCR 解析工具定义，供文本模型通过 MinerU 读取图片/PDF。
- `office.ts`: Word、PowerPoint、统一 Office action、多窗口文档选择和事务工作流工具定义。

关联模块：

- `../executors/createToolExecutors.ts`: 为注册表中的工具名装配执行器。
- `../../core/agentLoop/toolExecutor.ts`: 根据工具定义判断风险等级和审批策略。
- `../../prompts`: 提示词中的工具说明需要与注册表保持一致。

维护要求：

- 新增工具或 operation 时，必须同步更新 `prompts/templates/scenarios/office-tools.zh-CN.md`，避免模型退回高失败率的临时脚本。
- 统一 Office action 的 `operation` 描述要列出核心能力，例如 `createPresentation`、`deleteSlides`、`replaceText`、`styleTable`、`insertChart`。
- 风险等级需要与真实写入行为一致：只读检查为 `safe`，文件写入和另存为为 `moderate`。
