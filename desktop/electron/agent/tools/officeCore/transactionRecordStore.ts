import { randomUUID } from "node:crypto";
import { mkdir, readFile, readdir, rename, rm, writeFile } from "node:fs/promises";
import path from "node:path";

import type { OfficeTransactionRecord } from "./transactionTypes";

export async function saveOfficeTransaction(
  root: string,
  record: OfficeTransactionRecord,
): Promise<void> {
  record.updatedAt = new Date().toISOString();
  const directory = transactionDirectory(root, record.id);
  await mkdir(directory, { recursive: true });
  const destination = path.join(directory, "transaction.json");
  const temporary = path.join(directory, `.transaction.${randomUUID()}.tmp`);
  await writeFile(temporary, `${JSON.stringify(record, null, 2)}\n`, "utf8");
  await rm(destination, { force: true });
  await rename(temporary, destination);
}

export async function getOfficeTransaction(
  root: string,
  id: string,
): Promise<OfficeTransactionRecord> {
  validateRecordId(id);
  const record = JSON.parse(
    await readFile(path.join(transactionDirectory(root, id), "transaction.json"), "utf8"),
  ) as OfficeTransactionRecord;
  if (record.id !== id || !Array.isArray(record.steps) || !Array.isArray(record.snapshots)) {
    throw new Error("Office 事务记录已损坏");
  }
  return record;
}

export async function listOfficeTransactions(root: string): Promise<OfficeTransactionRecord[]> {
  let names: string[];
  try {
    names = await readdir(root);
  } catch {
    return [];
  }
  const records = await Promise.all(
    names.map(async (name) => {
      try {
        return await getOfficeTransaction(root, name);
      } catch {
        return undefined;
      }
    }),
  );
  return records
    .filter((record): record is OfficeTransactionRecord => Boolean(record))
    .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt));
}

export function transactionDirectory(root: string, id: string): string {
  validateRecordId(id);
  return path.join(path.resolve(root), id);
}

export function validateRecordId(id: string): void {
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(id)) {
    throw new Error("Office 事务 ID 无效");
  }
}
