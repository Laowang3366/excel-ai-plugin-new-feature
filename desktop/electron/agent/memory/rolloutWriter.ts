/**
 * Rollout 异步写入队列。
 *
 * 关联模块：
 * - sessionStore.ts: 组装 rollout JSONL 行后交给本模块排队落盘。
 *
 * 设计参考 Codex 的 rollout writer：主流程只提交写入命令，后台按文件批量追加；
 * 需要读取或退出前可调用 flush()，等待所有已排队写入完成。
 */

import * as fs from "fs";
import * as path from "path";

export type RolloutWriteBatch = (filePath: string, content: string) => Promise<void>;

export interface AsyncRolloutWriterOptions {
  writeBatch?: RolloutWriteBatch;
  scheduleDrain?: (run: () => void) => void;
  maxQueuedLines?: number;
}

export class AsyncRolloutWriter {
  private pending = new Map<string, string[]>();
  private pendingLineCount = 0;
  private inFlightLineCount = 0;
  private capacityWaiters: Array<() => void> = [];
  private scheduled = false;
  private currentDrain: Promise<void> | null = null;
  private lastError: Error | null = null;
  private readonly writeBatch: RolloutWriteBatch;
  private readonly scheduleDrain: (run: () => void) => void;
  private readonly maxQueuedLines: number;

  constructor(options: AsyncRolloutWriterOptions = {}) {
    this.writeBatch = options.writeBatch ?? defaultWriteBatch;
    this.scheduleDrain = options.scheduleDrain ?? defaultScheduleDrain;
    this.maxQueuedLines = Math.max(1, Math.floor(options.maxQueuedLines ?? 4096));
  }

  async enqueue(filePath: string, lines: string[]): Promise<void> {
    if (lines.length === 0) return;
    if (this.lastError) {
      throw this.lastError;
    }
    await this.waitForCapacity(lines.length);

    const existing = this.pending.get(filePath);
    if (existing) {
      existing.push(...lines);
    } else {
      this.pending.set(filePath, [...lines]);
    }
    this.pendingLineCount += lines.length;

    this.scheduleDrainIfIdle();
  }

  async flush(): Promise<void> {
    if (this.lastError) {
      this.lastError = null;
    }

    await this.startDrain();

    if (this.lastError) {
      const error = this.lastError;
      this.lastError = null;
      throw error;
    }
  }

  private startDrain(): Promise<void> {
    if (this.currentDrain) return this.currentDrain;

    this.scheduled = false;
    this.currentDrain = this.drainLoop().finally(() => {
      this.currentDrain = null;
      this.scheduleDrainIfIdle();
    });
    return this.currentDrain;
  }

  private async drainLoop(): Promise<void> {
    while (this.pending.size > 0) {
      const batches = this.pending;
      const batchLineCount = this.pendingLineCount;
      this.pending = new Map();
      this.pendingLineCount = 0;
      this.inFlightLineCount += batchLineCount;

      try {
        const entries = [...batches.entries()];
        const results = await Promise.allSettled(
          entries.map(([filePath, lines]) => this.writeBatch(filePath, lines.join(""))),
        );
        const failed: Array<[string, string[]]> = [];
        let firstError: unknown;
        for (let index = 0; index < results.length; index++) {
          const result = results[index];
          if (result.status === "rejected") {
            failed.push(entries[index]);
            firstError ??= result.reason;
          }
        }
        for (const [filePath, lines] of failed.reverse()) {
          const newerLines = this.pending.get(filePath) ?? [];
          this.pending.set(filePath, [...lines, ...newerLines]);
          this.pendingLineCount += lines.length;
        }
        if (firstError) {
          this.lastError = normalizeError(firstError);
          throw this.lastError;
        }
      } finally {
        this.inFlightLineCount = Math.max(0, this.inFlightLineCount - batchLineCount);
        this.notifyCapacityWaiters();
      }
    }
  }

  private async waitForCapacity(lineCount: number): Promise<void> {
    while (this.isOverCapacity(lineCount)) {
      await new Promise<void>((resolve) => {
        this.capacityWaiters.push(resolve);
      });
      if (this.lastError) {
        throw this.lastError;
      }
    }
  }

  private isOverCapacity(incomingLineCount: number): boolean {
    const queuedLineCount = this.pendingLineCount + this.inFlightLineCount;
    if (queuedLineCount === 0) return false;
    return queuedLineCount + incomingLineCount > this.maxQueuedLines;
  }

  private notifyCapacityWaiters(): void {
    const waiters = this.capacityWaiters;
    this.capacityWaiters = [];
    for (const resolve of waiters) {
      resolve();
    }
  }

  private scheduleDrainIfIdle(): void {
    if (this.lastError || this.scheduled || this.currentDrain || this.pending.size === 0) return;
    this.scheduled = true;
    this.scheduleDrain(() => {
      this.scheduled = false;
      void this.startDrain().catch((error) => {
        this.lastError = normalizeError(error);
        this.notifyCapacityWaiters();
      });
    });
  }
}

async function defaultWriteBatch(filePath: string, content: string): Promise<void> {
  await fs.promises.mkdir(path.dirname(filePath), { recursive: true });
  await fs.promises.writeFile(filePath, content, { flag: "a", encoding: "utf-8" });
}

function defaultScheduleDrain(run: () => void): void {
  if (typeof setImmediate === "function") {
    setImmediate(run);
  } else {
    setTimeout(run, 0);
  }
}

function normalizeError(error: unknown): Error {
  return error instanceof Error ? error : new Error(String(error));
}
