import { ipcApi } from "../services/ipcApi";
import type { SettingsActions, SettingsState } from "./settingsStore";

const KEY_MAP: Partial<Record<keyof SettingsState, string>> = {
  providers: "aiProviders",
  activeProviderId: "activeProvider",
  permissionMode: "permissionMode",
  showReasoning: "showReasoning",
  language: "language",
  theme: "theme",
  closeToTray: "closeToTray",
  officeAutoCompactEnabled: "officeAutoCompactEnabled",
  dynamicArrayFunctionsEnabled: "dynamicArrayFunctionsEnabled",
  remoteDataProcessingEnabled: "remoteDataProcessingEnabled",
  windowOpacity: "windowOpacity",
  pinnedFolders: "pinnedFolders",
  knowledgeEnabled: "knowledgeEnabled",
};

const COMPACTION_FIELDS: (keyof SettingsState)[] = ["compactionEnabled", "autoCompactThresholdPercent"];

export async function savePartial(
  keys: (keyof SettingsState)[],
  get: () => SettingsState & SettingsActions
): Promise<Record<string, unknown>> {
  const state = get();
  const toWrite: Array<[string, unknown]> = [];
  let needCompaction = false;

  for (const key of keys) {
    if (COMPACTION_FIELDS.includes(key)) {
      needCompaction = true;
      continue;
    }
    const storeKey = KEY_MAP[key];
    if (storeKey) {
      toWrite.push([storeKey, state[key]]);
    }
  }

  if (needCompaction) {
    const existingCompaction = await ipcApi.settings.get("compactionConfig") as Record<string, unknown> | null;
    toWrite.push(["compactionConfig", {
      ...(existingCompaction && typeof existingCompaction === "object" ? existingCompaction : {}),
      enabled: state.compactionEnabled,
      autoCompactThresholdPercent: state.autoCompactThresholdPercent,
    }]);
  }

  const persisted = await Promise.all(
    toWrite.map(async ([key, value]) => [key, await ipcApi.settings.set(key, value)] as const)
  );
  return Object.fromEntries(persisted);
}
