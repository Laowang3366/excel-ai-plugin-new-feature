import { defineConfig } from "vitest/config";
import { resolve } from "path";

export default defineConfig({
  test: {
    // 使用 node 环境（Electron 主进程代码 + 纯函数）
    environment: "node",
    // 匹配的测试文件
    include: ["src/**/*.test.ts", "electron/**/*.test.ts"],
    // 覆盖率配置
    coverage: {
      provider: "v8",
      include: [
        "src/utils/**",
        "src/store/**",
        "electron/agent/compaction.ts",
        "electron/agent/agentLoop/*.ts",
      ],
      reporter: ["text", "lcov"],
    },
    // 全局 setup
    globals: true,
    // 超时
    testTimeout: 10_000,
  },
  resolve: {
    alias: {
      "@": resolve(__dirname, "src"),
    },
  },
});
