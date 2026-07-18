import { describe, expect, it, vi } from "vitest";

import {
  applySettingRuntimeEffects,
  type SettingRuntimeEffectsDeps,
} from "./settingRuntimeEffects";

function createRuntimeDeps(): SettingRuntimeEffectsDeps {
  return {
    agents: [],
    mainWindow: null,
    getActiveAIConfig: vi.fn(() => ({ contextWindowSize: 64_000 }) as any),
    getActiveDataPath: vi.fn(() => "C:\\app-data"),
    getRuntimeSettingValue: vi.fn(() => undefined),
    refreshKnowledgeRuntime: vi.fn(async () => ({ ready: true })),
    applyWindowTheme: vi.fn(),
    applyWindowOpacity: vi.fn(),
    setDynamicArrayFunctionsEnabled: vi.fn(),
  };
}

function createRuntimeAgent() {
  return {
    updateAIConfig: vi.fn(),
    updateCompactionConfig: vi.fn(),
    updatePermissionMode: vi.fn(),
  };
}

describe("applySettingRuntimeEffects", () => {
  it("refreshes loaded agents and the knowledge runtime after provider changes", async () => {
    const deps = createRuntimeDeps();
    const agent = createRuntimeAgent();
    deps.agents = [agent as any];

    await applySettingRuntimeEffects("activeProvider", "provider-2", deps);

    expect(agent.updateAIConfig).toHaveBeenCalledWith({ contextWindowSize: 64_000 });
    expect(agent.updateCompactionConfig).toHaveBeenCalledTimes(1);
    expect(deps.refreshKnowledgeRuntime).toHaveBeenCalledTimes(1);
    expect(deps.refreshKnowledgeRuntime).toHaveBeenCalledWith(
      { contextWindowSize: 64_000 },
      "C:\\app-data",
    );
  });

  it("updates compaction without rebuilding unrelated runtime state", async () => {
    const deps = createRuntimeDeps();
    const agent = createRuntimeAgent();
    deps.agents = [agent as any];

    await applySettingRuntimeEffects("compactionConfig", { mode: "auto" }, deps);

    expect(agent.updateCompactionConfig).toHaveBeenCalledTimes(1);
    expect(agent.updateAIConfig).not.toHaveBeenCalled();
    expect(deps.refreshKnowledgeRuntime).not.toHaveBeenCalled();
  });

  it("routes permission and window-specific settings to their owners", async () => {
    const deps = createRuntimeDeps();
    const agent = createRuntimeAgent();
    const mainWindow = {} as any;
    deps.agents = [agent as any];
    deps.mainWindow = mainWindow;

    await applySettingRuntimeEffects("permissionMode", "confirm_all", deps);
    await applySettingRuntimeEffects("theme", "dark", deps);
    await applySettingRuntimeEffects("windowOpacity", 0.9, deps);
    await applySettingRuntimeEffects("dynamicArrayFunctionsEnabled", true, deps);

    expect(agent.updatePermissionMode).toHaveBeenCalledWith("confirm_all");
    expect(deps.applyWindowTheme).toHaveBeenCalledWith(mainWindow);
    expect(deps.applyWindowOpacity).toHaveBeenCalledWith(mainWindow);
    expect(deps.setDynamicArrayFunctionsEnabled).toHaveBeenCalledWith(true);
  });
});
