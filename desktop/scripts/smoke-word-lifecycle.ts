import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { OfficeComActionBridge } from "../electron/agent/tools/implementations/office/officeComActionBridge";
import { applyWordAdvancedAction } from "../electron/agent/tools/implementations/officeOpenXml/advancedWord";

async function main(): Promise<void> {
  const tempDir = await mkdtemp(path.join(os.tmpdir(), "wengge-word-lifecycle-"));
  const filePath = path.join(tempDir, "reopen.docx");

  try {
    const fixture = await applyWordAdvancedAction({
      operation: "createDocument",
      filePath,
      params: { title: "Word 生命周期测试", paragraphs: ["同一个文件需要连续打开两次。"] },
    });
    if (fixture.status !== "done") throw new Error(fixture.error || fixture.summary);

    const bridge = new OfficeComActionBridge();
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      const result = await bridge.executeAction({
        app: "word",
        action: "inspect",
        operation: "inspectDocumentFormatting",
        filePath,
      });
      if (result.status !== "done") {
        throw new Error(`第 ${attempt} 次打开失败: ${result.error || result.summary}`);
      }
    }

    process.stdout.write(`${JSON.stringify({ ok: true, attempts: 2, filePath }, null, 2)}\n`);
  } finally {
    await rm(tempDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  }
}

void main().catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack || error.message : String(error)}\n`);
  process.exitCode = 1;
});
