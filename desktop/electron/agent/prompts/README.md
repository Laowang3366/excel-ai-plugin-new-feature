# 提示词层

职责：组装系统提示词和按场景拆分的提示词片段，为模型请求提供稳定的行为约束和工具使用说明。

模块说明：

- `systemPrompt.ts`: 系统提示词总装配入口，按固定顺序组合 sections。
- `sections/modelPrompt.ts`: 模型身份、基础行为和输出约束。
- `sections/formulaAssistantPrompt.ts`: Excel 公式助手场景提示词。
- `sections/officeToolsPrompt.ts`: Word、PowerPoint、Excel 工具能力说明。
- `sections/permissionPrompt.ts`: 权限、审批和安全策略说明。
- `sections/scriptPrompt.ts`: 脚本执行与自动化代码约束。
- `sections/qualityPrompt.ts`: 输出质量、验证和错误处理要求。
- `sections/scenarioPrompt.ts`: 常见业务场景提示词。
- `sections/folderContextPrompt.ts`: 文件夹上下文注入提示词。
- `compactionPrompt.ts` / `templates/compaction.zh-CN.md`: 上下文压缩摘要提示词模板，供压缩摘要生成使用。

关联模块：

- `../core/agentLoop/buildStreamParams.ts`: 构建最终发送给模型的系统提示词。
- `../core/agentLoop/summaryGenerator.ts`: 读取压缩提示词模板生成摘要请求。
- `../tools/registry`: 工具定义需要与提示词中的能力描述保持一致。
