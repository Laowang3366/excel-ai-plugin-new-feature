import { beforeEach, describe, expect, it, vi } from "vitest";
import type { AiProviderConfig } from "../electronApi";

const mocks = vi.hoisted(() => ({
  get: vi.fn(),
  set: vi.fn(),
  getAll: vi.fn(),
}));

vi.mock("../services/ipcApi", () => ({
  ipcApi: {
    settings: {
      get: mocks.get,
      set: mocks.set,
      getAll: mocks.getAll,
    },
  },
}));

import { useSettingsStore } from "./settingsStore";

function makeProvider(patch: Partial<AiProviderConfig> = {}): AiProviderConfig {
  return {
    id: "provider-a",
    name: "Provider A",
    provider: "custom",
    apiKey: "",
    baseUrl: "",
    model: "",
    apiFormat: "openai",
    ...patch,
  };
}

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

describe("settingsStore dynamic array function support", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.setState({
      dynamicArrayFunctionsEnabled: true,
      isLoading: true,
    });
  });

  it("defaults dynamic array support to enabled when loading settings without a saved value", async () => {
    mocks.getAll.mockResolvedValue({});

    await useSettingsStore.getState().loadSettings();

    expect(useSettingsStore.getState().dynamicArrayFunctionsEnabled).toBe(true);
  });

  it("persists dynamic array support when the setting changes", async () => {
    mocks.set.mockResolvedValue(undefined);

    useSettingsStore.getState().setDynamicArrayFunctionsEnabled(false);
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(useSettingsStore.getState().dynamicArrayFunctionsEnabled).toBe(false);
    expect(mocks.set).toHaveBeenCalledWith("dynamicArrayFunctionsEnabled", false);
  });
});

describe("settingsStore remote data processing", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    useSettingsStore.setState({
      remoteDataProcessingEnabled: false,
      isLoading: true,
    });
  });

  it("defaults to local-only mode when loading settings without a saved value", async () => {
    mocks.getAll.mockResolvedValue({});

    await useSettingsStore.getState().loadSettings();

    expect(useSettingsStore.getState().remoteDataProcessingEnabled).toBe(false);
  });

  it("persists the remote-processing decision", async () => {
    mocks.set.mockResolvedValue(undefined);

    useSettingsStore.getState().setRemoteDataProcessingEnabled(true);
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(mocks.set).toHaveBeenCalledWith("remoteDataProcessingEnabled", true);
  });
});

describe("settingsStore provider configuration state", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mocks.set.mockResolvedValue(undefined);
    useSettingsStore.setState({
      providers: {},
      activeProviderId: "",
      isConfigured: false,
    });
  });

  it("recomputes configuration from the updated active provider", () => {
    useSettingsStore.setState({
      providers: {
        active: makeProvider({ id: "active" }),
      },
      activeProviderId: "active",
      isConfigured: false,
    });

    useSettingsStore.getState().updateProvider("active", {
      apiKey: "sk-test",
      baseUrl: "https://api.example.test/v1",
      model: "model-a",
    });

    expect(useSettingsStore.getState().isConfigured).toBe(true);
  });

  it("does not mark settings configured when adding a configured inactive provider", () => {
    useSettingsStore.setState({
      providers: {
        active: makeProvider({ id: "active" }),
      },
      activeProviderId: "active",
      isConfigured: false,
    });

    useSettingsStore.getState().addProvider(makeProvider({
      id: "inactive",
      apiKey: "sk-test",
      baseUrl: "https://api.example.test/v1",
      model: "model-a",
    }));

    expect(useSettingsStore.getState().activeProviderId).toBe("active");
    expect(useSettingsStore.getState().isConfigured).toBe(false);
  });
});
