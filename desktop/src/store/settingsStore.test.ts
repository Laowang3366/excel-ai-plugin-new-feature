import { beforeEach, describe, expect, it, vi } from "vitest";

const mocks = vi.hoisted(() => ({
  get: vi.fn(),
  set: vi.fn(),
}));

vi.mock("../services/ipcApi", () => ({
  ipcApi: {
    settings: {
      get: mocks.get,
      set: mocks.set,
    },
  },
}));

import { useSettingsStore } from "./settingsStore";

describe("settingsStore compaction config persistence", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.setState({
      compactionEnabled: true,
      autoCompactThresholdPercent: 80,
    });
  });

  it("preserves remote compaction settings when updating health fields", async () => {
    mocks.get.mockResolvedValue({
      enabled: true,
      autoCompactThresholdPercent: 80,
      compactionProvider: "remote",
      remoteCompactUrl: "https://compact.example.test/v2",
      remoteCompactApiKey: "remote-key",
      remoteCompactModel: "compact-model",
    });
    mocks.set.mockResolvedValue(undefined);

    useSettingsStore.getState().setCompactionEnabled(false);
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(mocks.set).toHaveBeenCalledWith("compactionConfig", {
      enabled: false,
      autoCompactThresholdPercent: 80,
      compactionProvider: "remote",
      remoteCompactUrl: "https://compact.example.test/v2",
      remoteCompactApiKey: "remote-key",
      remoteCompactModel: "compact-model",
    });
  });
});

describe("settingsStore window opacity persistence", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.setState({
      windowOpacity: 1,
    });
  });

  it("persists normalized window opacity when the setting changes", async () => {
    mocks.set.mockResolvedValue(undefined);

    useSettingsStore.getState().setWindowOpacity(0.42);
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(useSettingsStore.getState().windowOpacity).toBe(0.55);
    expect(mocks.set).toHaveBeenCalledWith("windowOpacity", 0.55);
  });
});
