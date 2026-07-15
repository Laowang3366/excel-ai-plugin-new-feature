# 贡献指南

本仓库包含两个独立 Node 项目：`desktop/` 是 Electron/React/.NET Office Worker 桌面应用，`product-site/` 是 Fastify 产品站。根目录没有 package workspace 或统一构建命令。

## 开始前

- 使用 Windows x64、Node.js 22.12+ 和根 `global.json` 指定的 .NET 8 SDK。
- 从运行代码和项目级 `package.json` 确认事实；当前架构与开发约束见 `docs/README.md`。
- 不提交 API Key、Token、私钥、生产数据库、真实文档或用户数据。
- 不修改与当前问题无关的本地文件或生成目录。

## 变更原则

1. 保持改动范围最小，并为故障边界补自动化测试。
2. Renderer 通过 `src/services/ipcApi.ts` 调用 IPC；IPC 输入必须经过共享 Zod Schema 校验。
3. Office COM/Open XML 实现放在 `.NET Worker`；不要重新引入任意 Shell、Python、PowerShell 或外部 JScript Office 自动化。
4. 修改模型可见 Office 工具时，同步更新注册、executor、契约、策略、Worker 桥、实现和提示词层。
5. 用户可见行为更新 `CHANGELOG.md`；架构、Office、构建、发布或数据边界更新对应当前文档。

## 本地验证

桌面端在 `desktop/` 运行：

```powershell
npm ci
npm run lint
npm run typecheck
npm test
npm run build
npm run office:test
```

修改 Worker 后至少运行 `office:test`。`electron:build` 会发布 Worker 并清空构建输出目录，只在需要安装包验证时执行。真实 Excel/WPS/Office 冒烟不是默认门禁，必须在隔离环境中按变更范围选择，且不得干扰用户已有进程。

产品站在 `product-site/` 运行：

```powershell
npm ci
npm audit --audit-level=high
npm test
```

## 安全敏感变更

以下范围视为安全敏感：CI/发布、Electron preload 与 IPC、Agent 工具和审批、凭据与出站策略、数据持久化、Office Worker、更新/签名、产品站认证/分析/部署。

合并前应满足：

- 至少两名人员批准，其中至少一名是 `.github/CODEOWNERS` 指定的所有者；
- 作者不得自行批准并合并；
- 所有必需 CI 检查通过，且没有未解决的安全审查意见；
- 危险行为变化带负向测试，外部验收项记录负责人和状态；
- 生产密钥轮换、证书、Environment approval、服务器 ACL 等动作不以单元测试替代。

仓库规则管理员应在 GitHub branch protection/ruleset 中实际强制上述审批和状态检查；文档本身不能提供平台级强制。

## Pull Request 清单

- [ ] 变更范围与风险已说明。
- [ ] 新行为和失败路径有自动化测试。
- [ ] 已运行适用的 lint、typecheck、test、build 和依赖审计。
- [ ] 当前文档和 `CHANGELOG.md` 已按需更新。
- [ ] 不包含真实秘密、用户数据或无关生成文件。
- [ ] 仍需真实 Office、生产环境或法律确认的事项已明确列出。

提交改动不代表项目授予额外许可。仓库的最终许可条款由项目所有者另行确认和发布。
