import {
  DEFAULT_PERMISSION_MODE,
  normalizePermissionMode,
  type PermissionMode,
} from "./approvalPolicy";

export const PERMISSION_MODE_PERSISTENCE_KEY =
  "wengge.excel-addin.permission-mode";
export const PERMISSION_MODE_PERSISTENCE_VERSION = 1;

/** Non-secret local settings storage (same shape as provider persistence). */
export interface PermissionModePersistenceStorage {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
}

interface PersistedPayload {
  version: number;
  permissionMode: PermissionMode;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

export function loadPermissionMode(
  storage?: PermissionModePersistenceStorage | null,
): PermissionMode {
  if (!storage) return DEFAULT_PERMISSION_MODE;
  try {
    const raw = storage.getItem(PERMISSION_MODE_PERSISTENCE_KEY);
    if (!raw) return DEFAULT_PERMISSION_MODE;
    const parsed: unknown = JSON.parse(raw);
    // Accept bare string or versioned object for forward-compat.
    if (typeof parsed === "string") {
      return normalizePermissionMode(parsed);
    }
    if (!isRecord(parsed)) return DEFAULT_PERMISSION_MODE;
    if (typeof parsed.permissionMode === "string") {
      return normalizePermissionMode(parsed.permissionMode);
    }
    return DEFAULT_PERMISSION_MODE;
  } catch {
    return DEFAULT_PERMISSION_MODE;
  }
}

export function persistPermissionMode(
  storage: PermissionModePersistenceStorage | null | undefined,
  mode: PermissionMode,
): void {
  if (!storage) return;
  const payload: PersistedPayload = {
    version: PERMISSION_MODE_PERSISTENCE_VERSION,
    permissionMode: normalizePermissionMode(mode),
  };
  try {
    storage.setItem(PERMISSION_MODE_PERSISTENCE_KEY, JSON.stringify(payload));
  } catch {
    /* quota / private mode — keep in-memory only */
  }
}

export class PermissionModeStore {
  private mode: PermissionMode;

  constructor(
    private readonly persistence?: PermissionModePersistenceStorage | null,
  ) {
    this.mode = loadPermissionMode(persistence);
  }

  get(): PermissionMode {
    return this.mode;
  }

  set(mode: PermissionMode | string | undefined | null): PermissionMode {
    this.mode = normalizePermissionMode(mode);
    persistPermissionMode(this.persistence, this.mode);
    return this.mode;
  }
}

export function getBrowserPermissionModePersistenceStorage():
  | PermissionModePersistenceStorage
  | undefined {
  try {
    if (typeof globalThis.localStorage === "undefined") return undefined;
    return globalThis.localStorage;
  } catch {
    return undefined;
  }
}

let browserStore: PermissionModeStore | null = null;

/** Shared browser store for ChatController default + App UI selector. */
export function getBrowserPermissionModeStore(): PermissionModeStore {
  if (!browserStore) {
    browserStore = new PermissionModeStore(
      getBrowserPermissionModePersistenceStorage(),
    );
  }
  return browserStore;
}

/** Test seam: reset module singleton. */
export function resetBrowserPermissionModeStoreForTests(): void {
  browserStore = null;
}
