import { createHash } from "node:crypto";
import * as fs from "node:fs";
import * as path from "node:path";
import { app } from "electron";
import { normalizePathForCompare } from "../settingsDataPath";

const REGISTRY_FILE = "managed-data-replicas.json";

export type ManagedReplicaCategory = "active_root" | "old_root" | "privacy_export";

export interface ManagedReplicaEntry {
  id: string;
  category: ManagedReplicaCategory;
  absolutePath: string;
  pathDigest: string;
  registeredAt: string;
  lastSeenAt: string;
  status: "active" | "pending_erase" | "erased";
  notes?: string;
}

export interface ManagedReplicaRegistryRecord {
  formatVersion: 1;
  installId: string;
  replicas: ManagedReplicaEntry[];
}

function registryPath(userDataPath = app.getPath("userData")): string {
  return path.join(userDataPath, REGISTRY_FILE);
}

function pathDigest(absolutePath: string): string {
  return createHash("sha256").update(normalizePathForCompare(absolutePath), "utf8").digest("hex");
}

function writeAtomic(filePath: string, record: ManagedReplicaRegistryRecord): void {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tempPath, `${JSON.stringify(record, null, 2)}\n`, "utf8");
  fs.renameSync(tempPath, filePath);
}

export class ManagedReplicaRegistry {
  private record: ManagedReplicaRegistryRecord;
  private readonly filePath: string;

  constructor(record: ManagedReplicaRegistryRecord, filePath: string) {
    this.record = record;
    this.filePath = filePath;
  }

  get installId(): string {
    return this.record.installId;
  }

  list(): ManagedReplicaEntry[] {
    return [...this.record.replicas];
  }

  listErasable(): ManagedReplicaEntry[] {
    return this.record.replicas.filter((entry) => entry.status !== "erased");
  }

  upsert(input: {
    category: ManagedReplicaCategory;
    absolutePath: string;
    status?: ManagedReplicaEntry["status"];
    notes?: string;
  }): ManagedReplicaEntry {
    const absolutePath = path.resolve(input.absolutePath);
    const digest = pathDigest(absolutePath);
    const now = new Date().toISOString();
    const existing = this.record.replicas.find(
      (entry) =>
        entry.pathDigest === digest ||
        normalizePathForCompare(entry.absolutePath) === normalizePathForCompare(absolutePath),
    );
    if (existing) {
      existing.category = input.category;
      existing.absolutePath = absolutePath;
      existing.lastSeenAt = now;
      existing.status = input.status ?? existing.status;
      if (input.notes !== undefined) existing.notes = input.notes;
      writeAtomic(this.filePath, this.record);
      return existing;
    }
    const entry: ManagedReplicaEntry = {
      id: digest.slice(0, 16),
      category: input.category,
      absolutePath,
      pathDigest: digest,
      registeredAt: now,
      lastSeenAt: now,
      status: input.status ?? "active",
      notes: input.notes,
    };
    this.record.replicas.push(entry);
    writeAtomic(this.filePath, this.record);
    return entry;
  }

  markStatus(absolutePath: string, status: ManagedReplicaEntry["status"], notes?: string): void {
    const normalized = normalizePathForCompare(absolutePath);
    const entry = this.record.replicas.find(
      (item) => normalizePathForCompare(item.absolutePath) === normalized,
    );
    if (!entry) return;
    entry.status = status;
    entry.lastSeenAt = new Date().toISOString();
    if (notes !== undefined) entry.notes = notes;
    writeAtomic(this.filePath, this.record);
  }

  /** Drop only successfully erased entries; keep pending_erase / active. */
  removeErasedEntries(): void {
    this.record = {
      ...this.record,
      replicas: this.record.replicas.filter((entry) => entry.status !== "erased"),
    };
    writeAtomic(this.filePath, this.record);
  }
}

export function openOrCreateManagedReplicaRegistry(options: {
  installId: string;
  userDataPath?: string;
}): ManagedReplicaRegistry {
  const filePath = registryPath(options.userDataPath);
  if (fs.existsSync(filePath)) {
    const raw = JSON.parse(fs.readFileSync(filePath, "utf8")) as ManagedReplicaRegistryRecord;
    if (raw.formatVersion !== 1 || !Array.isArray(raw.replicas) || !raw.installId) {
      throw new Error("invalid_replica_registry");
    }
    if (raw.installId !== options.installId) {
      throw new Error("replica_registry_install_id_mismatch");
    }
    return new ManagedReplicaRegistry(raw, filePath);
  }
  const record: ManagedReplicaRegistryRecord = {
    formatVersion: 1,
    installId: options.installId,
    replicas: [],
  };
  writeAtomic(filePath, record);
  return new ManagedReplicaRegistry(record, filePath);
}

export function managedReplicaRegistryPath(userDataPath?: string): string {
  return registryPath(userDataPath);
}
