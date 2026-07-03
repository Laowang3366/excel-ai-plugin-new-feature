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
