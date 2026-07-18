# 系统提示词架构说明

本项目参考 OpenAI Codex 源码提交 `5c19155cbd93bfa099016e7487259f61669823ff`，重构系统提示词的存储和组装方式。

## Codex 的处理方式

1. 基础指令独立存放。Codex 把默认基础指令放在 `codex-rs/protocol/src/prompts/base_instructions/default.md`，通过 `include_str!` 编译期加载，而不是把长提示词写进业务函数。
2. 模型配置持有提示词。`ModelInfo` 同时支持 `base_instructions`、`instructions_template` 和 `instructions_variables`；模板不可用时回退基础指令，避免复制整套模型提示词。
3. 动态上下文拆成片段。权限、协作模式、人格、技能、插件、环境和用户指令分别由 `codex-rs/core/src/context/` 下的模块渲染，再按确定顺序组装。
4. 静态前缀保持稳定。基础指令固定在请求前缀，发生模型或权限变化时追加差异片段；`prompt_caching.rs` 专门验证后续请求复用原有前缀和缓存键。
5. 行为由测试锁定。Codex 对人格替换、模型切换、上下文布局和提示词缓存都有单元测试或快照测试。

Codex 的工具参数不依赖基础提示词重复说明。`ToolSpec` 直接携带工具名称、用途和 JSON Schema，字段级 Schema 声明 description、required、enum、嵌套结构及 `additionalProperties:false`；基础提示词只保留通用调用纪律。工具数量较多时，Codex 通过 Direct/Deferred/Hidden 暴露策略和 tool search 控制本轮上下文，而不是删除参数定义后让模型猜测。参考：

- [shell_spec.rs](https://github.com/openai/codex/blob/main/codex-rs/core/src/tools/handlers/shell_spec.rs)
- [spec_plan.rs](https://github.com/openai/codex/blob/main/codex-rs/core/src/tools/spec_plan.rs)
- [default.md](https://github.com/openai/codex/blob/main/codex-rs/protocol/src/prompts/base_instructions/default.md)

## 本项目的对应实现

本项目需要同时兼容 OpenAI Responses、OpenAI Chat Completions 和 Anthropic。三类客户端当前都只接收一个 `systemPrompt`，因此不能直接复制 Codex 的多角色消息布局，但保留相同的分层原则：

1. `prompts/templates/system/` 存放稳定基础前缀。
2. `prompts/templates/scenarios/` 存放按本轮意图加载的场景规则。
3. `prompts/templates/runtime/` 存放 Office 连接、文件夹、功能开关和日期等可变事实。
4. `prompts/promptComposer.ts` 负责固定顺序、按 key 去重、清理空片段和严格变量替换。
5. `prompts/systemPrompt.ts` 只负责注册片段和判断触发条件，不再承载长篇提示词正文。
6. Office operation 的准确参数来自 `tools/registry` 中唯一的严格 Schema。执行前继续使用完整 discriminator 校验；发送给模型时，从同一份 Schema 自动生成紧凑参数目录，保留必填/可选、枚举、数组和嵌套结构，同时移除重复的顶层 `oneOf`。场景提示不再维护第二份参数清单。

最终顺序固定为：稳定基础前缀 -> 场景片段 -> 文件夹上下文 -> 运行时环境 -> 用户长期记忆。这样可以让支持前缀缓存的供应商复用静态部分，也能避免普通问答常驻全部 Office 场景规则。

## 维护规则

- 所有任务都必须遵守的规则才进入 `templates/system/`。
- 只对特定任务有用的规则进入 `templates/scenarios/`，并在 `systemPrompt.ts` 注册触发条件。
- 当前时间、连接状态、目录和用户设置只能进入 `templates/runtime/` 或长期记忆层。
- 不为不同模型复制整套提示词；确有模型差异时，优先增加小型变量或能力片段。
- 模板新增变量时必须通过 `renderPromptTemplate` 渲染，缺少变量应直接在测试或构建阶段暴露。
- 场景提示与模型可见工具必须共用同一份意图解析结果。Power Query、透视表、切片器等高级 Excel operation 默认不进入工具 Schema；只有当前用户消息明确表达外部/多来源可刷新 ETL 或交互式透视意图时才开放，执行层仍必须独立校验 `advancedIntent` 和 operation 前置条件。
- operation 参数新增或修改时，只改注册表严格 Schema，不在系统提示词或场景提示中复制参数格式。模型侧紧凑目录与执行侧校验必须从同一 Schema 派生，并用上下文预算测试防止重新膨胀。
