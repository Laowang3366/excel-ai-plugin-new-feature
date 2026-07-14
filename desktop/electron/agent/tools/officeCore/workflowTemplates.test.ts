import { mkdtemp, rm } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { afterEach, describe, expect, it } from "vitest";

import {
  deleteOfficeWorkflowTemplate,
  getOfficeWorkflowTemplate,
  listOfficeWorkflowTemplates,
  saveOfficeWorkflowTemplate,
} from "./workflowTemplates";

describe("Office workflow templates", () => {
  const roots: string[] = [];

  afterEach(async () => {
    await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
  });

  it("saves, updates, resolves by name, lists, and deletes templates", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-template-"));
    roots.push(root);
    const steps = [{
      app: "excel" as const,
      action: "inspect" as const,
      operation: "inspectCharts",
      filePath: "{{vars.sourcePath}}",
    }];

    const created = await saveOfficeWorkflowTemplate({ root, name: "月报", description: "生成月报", steps });
    const updated = await saveOfficeWorkflowTemplate({ root, id: created.id, name: "月报", description: "更新说明", steps });

    expect(updated.id).toBe(created.id);
    expect(updated.createdAt).toBe(created.createdAt);
    expect((await getOfficeWorkflowTemplate(root, "月报")).id).toBe(created.id);
    expect(await listOfficeWorkflowTemplates(root)).toEqual([updated]);
    expect(await deleteOfficeWorkflowTemplate(root, created.id)).toBe(true);
    expect(await listOfficeWorkflowTemplates(root)).toEqual([]);
    expect(await deleteOfficeWorkflowTemplate(root, created.id)).toBe(false);
  });

  it("rejects invalid IDs before resolving template paths", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "office-template-"));
    roots.push(root);
    await expect(deleteOfficeWorkflowTemplate(root, "../../outside")).rejects.toThrow("模板 ID 无效");
  });
});
