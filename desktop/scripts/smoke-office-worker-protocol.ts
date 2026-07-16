/**
 * H-10: new client (PROTOCOL_VERSION=2) against an intentional v1 Worker binary.
 *
 * Requires:
 *   WENGGE_OFFICE_WORKER_PATH=<absolute path to historical Wengge.OfficeWorker.exe>
 *
 * Build the fixture outside this script (e.g. checkout baseline 660c0597 and
 * `npm run office:publish`). Do not run office:publish here — that would replace
 * the fixture with the current v2 Worker.
 */
import { access } from "node:fs/promises";
import path from "node:path";

import {
  OfficeWorkerClient,
  OfficeWorkerError,
} from "../electron/agent/officeWorker/officeWorkerClient";

async function main(): Promise<void> {
  process.env.WENGGE_OFFICE_SMOKE = "1";

  const workerPath = process.env.WENGGE_OFFICE_WORKER_PATH?.trim();
  if (!workerPath) {
    throw new Error(
      "H-10 协议不匹配冒烟要求设置 WENGGE_OFFICE_WORKER_PATH 为历史 v1 Worker 的绝对路径（勿使用当前 office:publish 产物）。",
    );
  }
  if (!path.isAbsolute(workerPath)) {
    throw new Error(`WENGGE_OFFICE_WORKER_PATH 必须是绝对路径，收到相对路径: ${workerPath}`);
  }

  try {
    await access(workerPath);
  } catch {
    throw new Error(`WENGGE_OFFICE_WORKER_PATH 不可访问: ${workerPath}`);
  }
  process.env.WENGGE_OFFICE_WORKER_PATH = workerPath;

  const client = new OfficeWorkerClient();
  let mismatch: OfficeWorkerError | undefined;
  try {
    await client.invoke("worker.health", {}, 15_000);
  } catch (error) {
    if (error instanceof OfficeWorkerError && error.code === "protocol_mismatch") {
      mismatch = error;
    } else {
      throw error;
    }
  } finally {
    await client.dispose();
  }

  if (!mismatch) {
    throw new Error(
      `期望 OfficeWorkerError.code === "protocol_mismatch"，但 Worker 握手成功（路径可能指向当前 v2 Worker）: ${workerPath}`,
    );
  }

  process.stdout.write(
    `${JSON.stringify(
      {
        ok: true,
        scenario: "protocol_mismatch",
        code: mismatch.code,
        message: mismatch.message,
        workerPath,
      },
      null,
      2,
    )}\n`,
  );
}

main().catch((error) => {
  process.stderr.write(
    `${error instanceof Error ? error.stack || error.message : String(error)}\n`,
  );
  process.exitCode = 1;
});
