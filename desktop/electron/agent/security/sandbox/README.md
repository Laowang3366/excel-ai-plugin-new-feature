# 命令沙箱

职责：为 `shell.execute` 提供执行前安全评估、工作目录约束、审计记录和进程启动包装。

模块说明：

- `parseCommand.ts`: 将 shell 命令切分为可评估的子命令和 token。
- `defaultRules.ts`: 默认安全策略规则。
- `execPolicy.ts`: 策略评估、决策合并和工作目录白名单检查。
- `audit.ts`: 记录策略命中、拒绝和执行审计。
- `index.ts`: 对外提供 `evaluateCommand`、`runShellSpawn`、规则热更新和进程树终止能力。

关联模块：

- `../../core/agentLoop/toolExecutor.ts`: 在工具审批前评估 shell 命令风险。
- `../../tools/executors/shellExecutor.ts`: 在策略允许后执行命令。
- `../../../main-modules/ipcHandlers.ts`: 从设置页读取并应用用户自定义规则和可写根。
