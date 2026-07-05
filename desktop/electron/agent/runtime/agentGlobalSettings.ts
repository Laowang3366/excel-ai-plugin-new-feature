export interface AgentGlobalSettings {
  dynamicArrayFunctionsEnabled: boolean;
}

const DEFAULT_AGENT_GLOBAL_SETTINGS: AgentGlobalSettings = {
  dynamicArrayFunctionsEnabled: true,
};

let currentSettings: AgentGlobalSettings = { ...DEFAULT_AGENT_GLOBAL_SETTINGS };

export function getAgentGlobalSettings(): AgentGlobalSettings {
  return { ...currentSettings };
}

export function setAgentGlobalSettings(patch: Partial<AgentGlobalSettings>): void {
  currentSettings = {
    ...currentSettings,
    ...patch,
  };
}

export function setDynamicArrayFunctionsEnabled(value: unknown): void {
  setAgentGlobalSettings({
    dynamicArrayFunctionsEnabled: typeof value === "boolean"
      ? value
      : DEFAULT_AGENT_GLOBAL_SETTINGS.dynamicArrayFunctionsEnabled,
  });
}
