import { spawn, type ChildProcessWithoutNullStreams } from "node:child_process";
import { existsSync } from "node:fs";
import path from "node:path";
import { createLogger } from "../../shared/logger";

const log = createLogger("OfficeWorkerClient");
const PROTOCOL_VERSION = 1;
const DEFAULT_TIMEOUT_MS = 120_000;
const MAX_STDOUT_BUFFER = 64 * 1024 * 1024;

interface WorkerResponse {
  id: string;
  result?: unknown;
  error?: { code: string; message: string; data?: unknown };
}

interface PendingRequest {
  method: string;
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  timer: NodeJS.Timeout;
  probeTimer?: NodeJS.Timeout;
}

export class OfficeWorkerError extends Error {
  constructor(
    public readonly code: string,
    message: string,
    public readonly details?: unknown,
  ) {
    super(message);
    this.name = "OfficeWorkerError";
  }
}

export class OfficeWorkerClient {
  private process: ChildProcessWithoutNullStreams | null = null;
  private starting: Promise<void> | null = null;
  private stdoutBuffer = "";
  private requestSequence = 0;
  private readonly pending = new Map<string, PendingRequest>();

  async invoke<T>(method: string, params: Record<string, unknown> = {}, timeoutMs = DEFAULT_TIMEOUT_MS): Promise<T> {
    traceSmoke(`invoke.start:${method}:timeout=${timeoutMs}`);
    await this.ensureStarted();
    const worker = this.process;
    if (!worker || worker.killed || !worker.stdin.writable) {
      throw new OfficeWorkerError("worker_unavailable", "Office Worker 未运行");
    }

    const id = String(++this.requestSequence);
    return new Promise<T>((resolve, reject) => {
      const startedAt = Date.now();
      const probeTimer = process.env.WENGGE_OFFICE_SMOKE === "1"
        ? setInterval(() => traceSmoke(`invoke.waiting:${method}:elapsed=${Date.now() - startedAt}`), 10_000)
        : undefined;
      probeTimer?.unref();
      const timer = setTimeout(() => {
        if (probeTimer) clearInterval(probeTimer);
        traceSmoke(`invoke.timeout:${method}`);
        this.pending.delete(id);
        const error = new OfficeWorkerError(
          "worker_timeout_unknown_outcome",
          `Office Worker 操作超时并已终止；操作可能已经部分执行，请检查目标文件或使用事务恢复: ${method}`,
        );
        reject(error);
        this.failWorker(worker, error);
      }, Math.max(1_000, timeoutMs));
      this.pending.set(id, { method, resolve: value => resolve(value as T), reject, timer, probeTimer });
      worker.stdin.write(`${JSON.stringify({ id, method, params })}\n`, "utf8", error => {
        if (!error) return;
        clearTimeout(timer);
        if (probeTimer) clearInterval(probeTimer);
        this.pending.delete(id);
        reject(new OfficeWorkerError("worker_write_failed", error.message));
      });
    });
  }

  async dispose(): Promise<void> {
    const worker = this.process;
    this.process = null;
    this.starting = null;
    this.rejectAll(new OfficeWorkerError("worker_stopped", "Office Worker 已停止"));
    if (!worker || worker.killed) return;
    worker.stdin.end();
    await new Promise<void>(resolve => {
      const timer = setTimeout(() => {
        if (!worker.killed) worker.kill();
        resolve();
      }, 2_000);
      worker.once("exit", () => {
        clearTimeout(timer);
        resolve();
      });
    });
  }

  private async ensureStarted(): Promise<void> {
    if (this.process && !this.process.killed) return;
    if (this.starting) return this.starting;
    this.starting = this.start();
    try {
      await this.starting;
    } finally {
      this.starting = null;
    }
  }

  private async start(): Promise<void> {
    const executablePath = resolveOfficeWorkerPath();
    const worker = spawn(executablePath, [], {
      windowsHide: true,
      stdio: ["pipe", "pipe", "pipe"],
      env: { ...process.env, DOTNET_EnableDiagnostics: "0" },
    });
    this.stdoutBuffer = "";
    this.process = worker;
    worker.stdout.setEncoding("utf8");
    worker.stderr.setEncoding("utf8");
    worker.stdout.on("data", chunk => this.onStdout(worker, String(chunk)));
    worker.stderr.on("data", chunk => {
      const message = String(chunk).trim();
      if (process.env.WENGGE_OFFICE_SMOKE === "1") {
        process.stderr.write(`${message}\n`);
      } else {
        log.warn("Worker stderr", { message });
      }
    });
    worker.on("error", error => this.onWorkerExit(worker, error));
    worker.on("exit", (code, signal) => this.onWorkerExit(worker,
      new OfficeWorkerError("worker_exited", `Office Worker 已退出: code=${code ?? "null"}, signal=${signal ?? "null"}`),
    ));

    const health = await this.invokeWithoutStart<{
      ready: boolean;
      protocolVersion: number;
      workerVersion?: string;
    }>("worker.health", {}, 10_000);
    if (!health.ready || health.protocolVersion !== PROTOCOL_VERSION) {
      await this.dispose();
      throw new OfficeWorkerError(
        "protocol_mismatch",
        `Office Worker 协议不兼容: 应用=${PROTOCOL_VERSION}, Worker=${health.protocolVersion}`,
      );
    }
    log.info("Office Worker ready", health);
  }

  private invokeWithoutStart<T>(method: string, params: Record<string, unknown>, timeoutMs: number): Promise<T> {
    const worker = this.process;
    if (!worker) return Promise.reject(new OfficeWorkerError("worker_unavailable", "Office Worker 启动失败"));
    const id = String(++this.requestSequence);
    return new Promise<T>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        const error = new OfficeWorkerError("worker_timeout", `Office Worker 握手超时: ${method}`);
        reject(error);
        this.failWorker(worker, error);
      }, timeoutMs);
      this.pending.set(id, { method, resolve: value => resolve(value as T), reject, timer });
      worker.stdin.write(`${JSON.stringify({ id, method, params })}\n`, "utf8");
    });
  }

  private onStdout(worker: ChildProcessWithoutNullStreams, chunk: string): void {
    if (this.process !== worker) return;
    this.stdoutBuffer += chunk;
    if (this.stdoutBuffer.length > MAX_STDOUT_BUFFER) {
      this.failWorker(worker, new OfficeWorkerError("protocol_overflow", "Office Worker 输出超过限制"));
      return;
    }

    while (true) {
      const newline = this.stdoutBuffer.indexOf("\n");
      if (newline < 0) return;
      const line = this.stdoutBuffer.slice(0, newline).trim();
      this.stdoutBuffer = this.stdoutBuffer.slice(newline + 1);
      if (!line) continue;
      this.handleResponse(line);
    }
  }

  private handleResponse(line: string): void {
    let response: WorkerResponse;
    try {
      response = JSON.parse(line) as WorkerResponse;
    } catch (error) {
      log.error("Invalid Worker JSON", { line: line.slice(0, 500), error: String(error) });
      return;
    }
    const pending = this.pending.get(response.id);
    if (!pending) return;
    clearTimeout(pending.timer);
    if (pending.probeTimer) clearInterval(pending.probeTimer);
    this.pending.delete(response.id);
    if (response.error) {
      traceSmoke(`invoke.error:${pending.method}:${response.error.code}`);
      pending.reject(new OfficeWorkerError(response.error.code, response.error.message, response.error.data));
    } else {
      traceSmoke(`invoke.done:${pending.method}`);
      pending.resolve(response.result);
    }
  }

  private onWorkerExit(worker: ChildProcessWithoutNullStreams, error: Error): void {
    if (this.process !== worker) return;
    this.process = null;
    this.stdoutBuffer = "";
    this.rejectAll(error);
    log.error("Office Worker stopped", { message: error.message });
  }

  private failWorker(worker: ChildProcessWithoutNullStreams, error: Error): void {
    if (this.process !== worker) return;
    this.onWorkerExit(worker, error);
    if (!worker.killed) worker.kill();
  }

  private rejectAll(error: Error): void {
    for (const pending of this.pending.values()) {
      clearTimeout(pending.timer);
      if (pending.probeTimer) clearInterval(pending.probeTimer);
      pending.reject(error);
    }
    this.pending.clear();
  }
}

function traceSmoke(message: string): void {
  if (process.env.WENGGE_OFFICE_SMOKE === "1") {
    process.stderr.write(`[office-smoke] worker-client:${message}\n`);
  }
}

export function resolveOfficeWorkerPath(): string {
  const executable = "Wengge.OfficeWorker.exe";
  const candidates = [
    process.env.WENGGE_OFFICE_WORKER_PATH,
    process.resourcesPath ? path.join(process.resourcesPath, "office-worker", executable) : undefined,
    path.resolve(process.cwd(), "dotnet", "publish", "win-x64", executable),
    path.resolve(process.cwd(), "desktop", "dotnet", "publish", "win-x64", executable),
    path.resolve(__dirname, "../../../../dotnet/publish/win-x64", executable),
  ].filter((candidate): candidate is string => Boolean(candidate));
  const match = candidates.find(existsSync);
  if (match) return match;
  throw new OfficeWorkerError("worker_not_found", `找不到 Office Worker: ${candidates.join(", ")}`);
}

let sharedClient: OfficeWorkerClient | null = null;

export function getOfficeWorkerClient(): OfficeWorkerClient {
  sharedClient ??= new OfficeWorkerClient();
  return sharedClient;
}

export async function disposeOfficeWorkerClient(): Promise<void> {
  const client = sharedClient;
  sharedClient = null;
  await client?.dispose();
}
