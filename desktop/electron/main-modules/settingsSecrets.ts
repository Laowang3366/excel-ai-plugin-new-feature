import { SETTINGS_SECRET_MASK } from "../shared/settingsSecretContract";

export { SETTINGS_SECRET_MASK } from "../shared/settingsSecretContract";
const ENCRYPTED_SECRET_PREFIX = "safe-storage:v1:";

export interface SettingsSecretCipher {
  isAvailable(): boolean;
  encrypt(value: string): string;
  decrypt(value: string): string;
}

type ProviderRecord = Record<string, unknown>;
type ProviderMap = Record<string, ProviderRecord>;

function asRecord(value: unknown): Record<string, unknown> {
  return value && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};
}

function isEncryptedSecret(value: string): boolean {
  return value.startsWith(ENCRYPTED_SECRET_PREFIX);
}

function encryptSecret(value: unknown, cipher: SettingsSecretCipher): string {
  if (typeof value !== "string" || !value) return "";
  if (isEncryptedSecret(value)) return value;
  if (!cipher.isAvailable()) throw new Error("secure_storage_unavailable");
  return `${ENCRYPTED_SECRET_PREFIX}${cipher.encrypt(value)}`;
}

function decryptSecret(value: unknown, cipher: SettingsSecretCipher): string {
  if (typeof value !== "string" || !value) return "";
  if (!isEncryptedSecret(value)) return value;
  if (!cipher.isAvailable()) throw new Error("secure_storage_unavailable");
  return cipher.decrypt(value.slice(ENCRYPTED_SECRET_PREFIX.length));
}

function sanitizeSecret(value: unknown): string {
  return typeof value === "string" && value ? SETTINGS_SECRET_MASK : "";
}

function protectHeaders(
  incoming: unknown,
  current: unknown,
  cipher: SettingsSecretCipher
): Record<string, string> | undefined {
  if (!incoming || typeof incoming !== "object" || Array.isArray(incoming)) return undefined;
  const currentHeaders = asRecord(current);
  return Object.fromEntries(
    Object.entries(incoming as Record<string, unknown>).map(([name, value]) => {
      const nextValue = value === SETTINGS_SECRET_MASK ? currentHeaders[name] : value;
      return [name, encryptSecret(nextValue, cipher)];
    })
  );
}

function decryptHeaders(
  headers: unknown,
  cipher: SettingsSecretCipher
): Record<string, string> | undefined {
  if (!headers || typeof headers !== "object" || Array.isArray(headers)) return undefined;
  return Object.fromEntries(
    Object.entries(headers as Record<string, unknown>).map(([name, value]) => [
      name,
      decryptSecret(value, cipher),
    ])
  );
}

function sanitizeHeaders(headers: unknown): Record<string, string> | undefined {
  if (!headers || typeof headers !== "object" || Array.isArray(headers)) return undefined;
  return Object.fromEntries(
    Object.entries(headers as Record<string, unknown>).map(([name, value]) => [
      name,
      sanitizeSecret(value),
    ])
  );
}

export function protectProviderMapForStorage(
  incoming: unknown,
  current: unknown,
  cipher: SettingsSecretCipher
): ProviderMap {
  const incomingProviders = asRecord(incoming) as ProviderMap;
  const currentProviders = asRecord(current) as ProviderMap;
  return Object.fromEntries(
    Object.entries(incomingProviders).map(([id, rawProvider]) => {
      const provider = asRecord(rawProvider);
      const currentProvider = asRecord(currentProviders[id]);
      const incomingApiKey = provider.apiKey === SETTINGS_SECRET_MASK
        ? currentProvider.apiKey
        : provider.apiKey;
      const customHeaders = protectHeaders(
        provider.customHeaders,
        currentProvider.customHeaders,
        cipher
      );
      return [id, {
        ...provider,
        apiKey: encryptSecret(incomingApiKey, cipher),
        ...(customHeaders ? { customHeaders } : {}),
      }];
    })
  );
}

export function migrateSettingsSecrets(
  settings: Record<string, unknown>,
  cipher: SettingsSecretCipher
): Record<string, unknown> {
  const providers = protectProviderMapForStorage(
    settings.aiProviders,
    settings.aiProviders,
    cipher
  );
  return {
    ...settings,
    aiProviders: providers,
    mineruApiToken: encryptSecret(settings.mineruApiToken, cipher),
    ...(settings.ocrMineruApiToken !== undefined
      ? { ocrMineruApiToken: encryptSecret(settings.ocrMineruApiToken, cipher) }
      : {}),
  };
}

export function protectSettingValueForStorage(
  key: string,
  incoming: unknown,
  current: unknown,
  cipher: SettingsSecretCipher
): unknown {
  if (key === "aiProviders") {
    return protectProviderMapForStorage(incoming, current, cipher);
  }
  if (key === "mineruApiToken" || key === "ocrMineruApiToken") {
    const nextValue = incoming === SETTINGS_SECRET_MASK ? current : incoming;
    return encryptSecret(nextValue, cipher);
  }
  return incoming;
}

export function sanitizeSettingsForRenderer(
  settings: Record<string, unknown>
): Record<string, unknown> {
  const providers = asRecord(settings.aiProviders) as ProviderMap;
  return {
    ...settings,
    aiProviders: Object.fromEntries(
      Object.entries(providers).map(([id, rawProvider]) => {
        const provider = asRecord(rawProvider);
        const customHeaders = sanitizeHeaders(provider.customHeaders);
        return [id, {
          ...provider,
          apiKey: sanitizeSecret(provider.apiKey),
          ...(customHeaders ? { customHeaders } : {}),
        }];
      })
    ),
    mineruApiToken: sanitizeSecret(settings.mineruApiToken),
    ...(settings.ocrMineruApiToken !== undefined
      ? { ocrMineruApiToken: sanitizeSecret(settings.ocrMineruApiToken) }
      : {}),
  };
}

export function decryptProviderForRuntime(
  provider: ProviderRecord,
  cipher: SettingsSecretCipher
): ProviderRecord {
  const customHeaders = decryptHeaders(provider.customHeaders, cipher);
  return {
    ...provider,
    apiKey: decryptSecret(provider.apiKey, cipher),
    ...(customHeaders ? { customHeaders } : {}),
  };
}

export function decryptSettingValueForRuntime(
  key: string,
  value: unknown,
  cipher: SettingsSecretCipher
): unknown {
  if (key === "mineruApiToken" || key === "ocrMineruApiToken") {
    return decryptSecret(value, cipher);
  }
  return value;
}
