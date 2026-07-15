import { createHash, randomBytes } from "node:crypto";
import * as fs from "node:fs";
import * as path from "node:path";
import { app, safeStorage } from "electron";
import { generateDataKey } from "./aesGcm";
import { hasProtectedLocalDataArtifacts } from "./protectedDataPresence";

const KEYSTORE_FILE = "local-data-keystore.json";
const WRAPPED_PREFIX = "safe-storage:ldp-key:v1:";

export interface DataKeyEntry {
  keyId: number;
  wrappedKey: string;
  createdAt: string;
  role: "active" | "pending" | "retired";
}

export interface DataKeystoreRecord {
  formatVersion: 1;
  installId: string;
  currentKeyId: number;
  pendingKeyId: number | null;
  keys: DataKeyEntry[];
}

export interface DataKeystoreCipher {
  isAvailable(): boolean;
  encrypt(value: string): string;
  decrypt(value: string): string;
}

export interface OpenDataKeystoreResult {
  keystore: DataKeystore;
  created: boolean;
}

function defaultCipher(): DataKeystoreCipher {
  return {
    isAvailable: () => safeStorage?.isEncryptionAvailable() === true,
    encrypt: (value) => safeStorage.encryptString(value).toString("base64"),
    decrypt: (value) => safeStorage.decryptString(Buffer.from(value, "base64")),
  };
}

function keystorePath(userDataPath = app.getPath("userData")): string {
  return path.join(userDataPath, KEYSTORE_FILE);
}

function wrapKey(key: Buffer, cipher: DataKeystoreCipher): string {
  if (!cipher.isAvailable()) throw new Error("secure_storage_unavailable");
  return `${WRAPPED_PREFIX}${cipher.encrypt(key.toString("base64"))}`;
}

function unwrapKey(wrapped: string, cipher: DataKeystoreCipher): Buffer {
  if (!wrapped.startsWith(WRAPPED_PREFIX)) throw new Error("invalid_wrapped_key");
  if (!cipher.isAvailable()) throw new Error("secure_storage_unavailable");
  return Buffer.from(cipher.decrypt(wrapped.slice(WRAPPED_PREFIX.length)), "base64");
}

function normalizeRecord(raw: DataKeystoreRecord): DataKeystoreRecord {
  if (raw.formatVersion !== 1 || !raw.installId || !Array.isArray(raw.keys)) {
    throw new Error("invalid_keystore_format");
  }
  return {
    formatVersion: 1,
    installId: raw.installId,
    currentKeyId: raw.currentKeyId,
    pendingKeyId: raw.pendingKeyId ?? null,
    keys: raw.keys.map((entry) => ({
      keyId: entry.keyId,
      wrappedKey: entry.wrappedKey,
      createdAt: entry.createdAt,
      role: entry.role ?? (entry.keyId === raw.currentKeyId ? "active" : "retired"),
    })),
  };
}

function loadRecord(filePath: string): DataKeystoreRecord | null {
  if (!fs.existsSync(filePath)) return null;
  return normalizeRecord(JSON.parse(fs.readFileSync(filePath, "utf8")) as DataKeystoreRecord);
}

function writeRecordAtomic(filePath: string, record: DataKeystoreRecord): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tempPath, `${JSON.stringify(record, null, 2)}\n`, "utf8");
  fs.renameSync(tempPath, filePath);
}

export class DataKeystore {
  private record: DataKeystoreRecord;
  private readonly filePath: string;
  private readonly cipher: DataKeystoreCipher;
  private readonly cache = new Map<number, Buffer>();

  constructor(record: DataKeystoreRecord, filePath: string, cipher: DataKeystoreCipher) {
    this.record = record;
    this.filePath = filePath;
    this.cipher = cipher;
  }

  get installId(): string {
    return this.record.installId;
  }

  get currentKeyId(): number {
    return this.record.currentKeyId;
  }

  get pendingKeyId(): number | null {
    return this.record.pendingKeyId;
  }

  getKey(keyId = this.record.currentKeyId): Buffer {
    const cached = this.cache.get(keyId);
    if (cached) return cached;
    const entry = this.record.keys.find((item) => item.keyId === keyId);
    if (!entry) throw new Error(`missing_data_key:${keyId}`);
    const key = unwrapKey(entry.wrappedKey, this.cipher);
    this.cache.set(keyId, key);
    return key;
  }

  listKeyIds(): number[] {
    return this.record.keys.map((item) => item.keyId).sort((a, b) => a - b);
  }

  createPendingKey(): number {
    if (this.record.pendingKeyId != null) return this.record.pendingKeyId;
    const nextKeyId = Math.max(this.record.currentKeyId, ...this.listKeyIds(), 0) + 1;
    const nextKey = generateDataKey();
    const nextRecord: DataKeystoreRecord = {
      ...this.record,
      pendingKeyId: nextKeyId,
      keys: [
        ...this.record.keys,
        {
          keyId: nextKeyId,
          wrappedKey: wrapKey(nextKey, this.cipher),
          createdAt: new Date().toISOString(),
          role: "pending",
        },
      ],
    };
    writeRecordAtomic(this.filePath, nextRecord);
    this.record = nextRecord;
    this.cache.set(nextKeyId, nextKey);
    return nextKeyId;
  }

  /**
   * Promote pending → current. Disk write first; keep retired keys until purgeRetiredKeys().
   */
  commitPendingKey(): { previousKeyId: number; nextKeyId: number } {
    const nextKeyId = this.record.pendingKeyId;
    if (nextKeyId == null) throw new Error("no_pending_data_key");
    const previousKeyId = this.record.currentKeyId;
    const nextRecord: DataKeystoreRecord = {
      ...this.record,
      currentKeyId: nextKeyId,
      pendingKeyId: null,
      keys: this.record.keys.map((entry) => {
        if (entry.keyId === nextKeyId) return { ...entry, role: "active" as const };
        if (entry.keyId === previousKeyId) return { ...entry, role: "retired" as const };
        return entry;
      }),
    };
    writeRecordAtomic(this.filePath, nextRecord);
    this.record = nextRecord;
    return { previousKeyId, nextKeyId };
  }

  /** Drop retired keys only after rotation fully succeeds (data + restore). */
  purgeRetiredKeys(): void {
    const nextRecord: DataKeystoreRecord = {
      ...this.record,
      keys: this.record.keys.filter((entry) => entry.role !== "retired"),
    };
    writeRecordAtomic(this.filePath, nextRecord);
    this.record = nextRecord;
    this.cache.clear();
  }

  discardPendingKey(): void {
    if (this.record.pendingKeyId == null) return;
    const pending = this.record.pendingKeyId;
    const nextRecord: DataKeystoreRecord = {
      ...this.record,
      pendingKeyId: null,
      keys: this.record.keys.filter((entry) => entry.keyId !== pending),
    };
    writeRecordAtomic(this.filePath, nextRecord);
    this.record = nextRecord;
    this.cache.delete(pending);
  }

  /** After failed post-commit restore: put previous key back as current and drop the failed new key. */
  revertToKeyId(keyId: number): void {
    if (!this.record.keys.some((entry) => entry.keyId === keyId)) {
      throw new Error(`missing_data_key:${keyId}`);
    }
    const nextRecord: DataKeystoreRecord = {
      ...this.record,
      currentKeyId: keyId,
      pendingKeyId: null,
      keys: this.record.keys
        .filter(
          (entry) =>
            entry.keyId === keyId ||
            entry.role === "retired" ||
            entry.keyId === this.record.currentKeyId,
        )
        .map((entry) => ({
          ...entry,
          role: entry.keyId === keyId ? ("active" as const) : ("retired" as const),
        }))
        .filter((entry) => entry.keyId === keyId),
    };
    writeRecordAtomic(this.filePath, nextRecord);
    this.record = nextRecord;
    this.cache.clear();
  }

  /**
   * Erase-path atomic swap: promote pending to sole active key and drop all prior keys
   * in a single writeRecordAtomic. On write failure, in-memory state is unchanged.
   */
  commitPendingKeyAndPurgePriorKeys(): {
    replacementKeyId: number;
    destroyedKeyIds: number[];
  } {
    const replacementKeyId = this.record.pendingKeyId;
    if (replacementKeyId == null) throw new Error("no_pending_data_key");
    const replacement = this.record.keys.find((entry) => entry.keyId === replacementKeyId);
    if (!replacement) throw new Error(`missing_data_key:${replacementKeyId}`);
    // Resolve material before disk write so unwrap failure never leaves sole-key disk + old memory.
    const replacementMaterial =
      this.cache.get(replacementKeyId) ?? unwrapKey(replacement.wrappedKey, this.cipher);
    const destroyedKeyIds = this.listKeyIds().filter((id) => id !== replacementKeyId);
    const nextRecord: DataKeystoreRecord = {
      formatVersion: 1,
      installId: this.record.installId,
      currentKeyId: replacementKeyId,
      pendingKeyId: null,
      keys: [
        {
          keyId: replacementKeyId,
          wrappedKey: replacement.wrappedKey,
          createdAt: replacement.createdAt,
          role: "active",
        },
      ],
    };
    writeRecordAtomic(this.filePath, nextRecord);
    this.record = nextRecord;
    this.cache.clear();
    this.cache.set(replacementKeyId, replacementMaterial);
    return { replacementKeyId, destroyedKeyIds };
  }

  /** After empty-keystore open, mint a fresh active key while reusing installId. */
  reseedActiveKey(): number {
    if (!this.cipher.isAvailable()) throw new Error("secure_storage_unavailable");
    const nextKeyId = 1;
    const nextKey = generateDataKey();
    const nextRecord: DataKeystoreRecord = {
      formatVersion: 1,
      installId: this.record.installId,
      currentKeyId: nextKeyId,
      pendingKeyId: null,
      keys: [
        {
          keyId: nextKeyId,
          wrappedKey: wrapKey(nextKey, this.cipher),
          createdAt: new Date().toISOString(),
          role: "active",
        },
      ],
    };
    writeRecordAtomic(this.filePath, nextRecord);
    this.record = nextRecord;
    this.cache.clear();
    this.cache.set(nextKeyId, nextKey);
    return nextKeyId;
  }

  installIdDigest(): string {
    return createHash("sha256").update(this.record.installId, "utf8").digest("hex");
  }
}

export function openOrCreateDataKeystore(options: {
  userDataPath?: string;
  cipher?: DataKeystoreCipher;
  /**
   * Required for production init. When set, empty reseed and brand-new keystore creation
   * are refused if hasProtectedLocalDataArtifacts(dataRoot, userDataPath) is true.
   * Callers cannot pass a boolean to claim the disk is clean.
   */
  dataRoot: string;
}): OpenDataKeystoreResult {
  const cipher = options.cipher ?? defaultCipher();
  const userDataPath = options.userDataPath ?? app.getPath("userData");
  const filePath = keystorePath(userDataPath);
  const blocked = hasProtectedLocalDataArtifacts(options.dataRoot, userDataPath);
  const existing = loadRecord(filePath);
  if (existing) {
    const keystore = new DataKeystore(existing, filePath, cipher);
    if (existing.keys.length === 0 || existing.currentKeyId === 0) {
      if (blocked) {
        throw new Error("empty_keystore_with_protected_data");
      }
      keystore.reseedActiveKey();
      return { keystore, created: false };
    }
    return { keystore, created: false };
  }
  if (blocked) {
    throw new Error("empty_keystore_with_protected_data");
  }
  if (!cipher.isAvailable()) throw new Error("secure_storage_unavailable");
  const key = generateDataKey();
  const record: DataKeystoreRecord = {
    formatVersion: 1,
    installId: randomBytes(16).toString("hex"),
    currentKeyId: 1,
    pendingKeyId: null,
    keys: [
      {
        keyId: 1,
        wrappedKey: wrapKey(key, cipher),
        createdAt: new Date().toISOString(),
        role: "active",
      },
    ],
  };
  writeRecordAtomic(filePath, record);
  const keystore = new DataKeystore(record, filePath, cipher);
  keystore.getKey(1);
  return { keystore, created: true };
}

export function dataKeystoreFilePath(userDataPath?: string): string {
  return keystorePath(userDataPath);
}
