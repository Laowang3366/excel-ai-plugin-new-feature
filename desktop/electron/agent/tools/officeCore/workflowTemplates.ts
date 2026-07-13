import { randomUUID } from "node:crypto";
import { mkdir, readFile, readdir, rename, rm, writeFile } from "node:fs/promises";
import path from "node:path";

import type { OfficeWorkflowStepInput } from "./workflowStepExecution";

export interface OfficeWorkflowTemplate {
  id: string;
  name: string;
  description?: string;
  createdAt: string;
  updatedAt: string;
  steps: OfficeWorkflowStepInput[];
}

export async function saveOfficeWorkflowTemplate(input: {
  root: string;
  id?: string;
  name: string;
  description?: string;
  steps: OfficeWorkflowStepInput[];
}): Promise<OfficeWorkflowTemplate> {
  const name = input.name.trim();
  if (!name) throw new Error("工作流模板名称不能为空");
  if (input.steps.length === 0) throw new Error("工作流模板至少需要一个步骤");
  const existing = input.id ? await getOfficeWorkflowTemplate(input.root, input.id) : undefined;
  const now = new Date().toISOString();
  const template: OfficeWorkflowTemplate = {
    id: existing?.id || randomUUID(),
    name,
    description: input.description?.trim() || undefined,
    createdAt: existing?.createdAt || now,
    updatedAt: now,
    steps: input.steps,
  };
  await writeTemplate(input.root, template);
  return template;
}

export async function getOfficeWorkflowTemplate(root: string, idOrName: string): Promise<OfficeWorkflowTemplate> {
  const byId = isUuid(idOrName) ? await readTemplate(root, idOrName) : undefined;
  if (byId) return byId;
  const byName = (await listOfficeWorkflowTemplates(root)).find((item) => item.name === idOrName);
  if (!byName) throw new Error(`找不到 Office 工作流模板: ${idOrName}`);
  return byName;
}

export async function listOfficeWorkflowTemplates(root: string): Promise<OfficeWorkflowTemplate[]> {
  const directory = templateRoot(root);
  let names: string[];
  try { names = await readdir(directory); } catch { return []; }
  const templates = await Promise.all(names.filter((name) => name.endsWith(".json")).map(async (name) => {
    try { return await readTemplate(root, path.basename(name, ".json")); } catch { return undefined; }
  }));
  return templates
    .filter((item): item is OfficeWorkflowTemplate => Boolean(item))
    .sort((left, right) => right.updatedAt.localeCompare(left.updatedAt));
}

export async function deleteOfficeWorkflowTemplate(root: string, id: string): Promise<boolean> {
  validateId(id);
  const filePath = templatePath(root, id);
  try { await rm(filePath); return true; } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") return false;
    throw error;
  }
}

async function writeTemplate(root: string, template: OfficeWorkflowTemplate): Promise<void> {
  const directory = templateRoot(root);
  await mkdir(directory, { recursive: true });
  const destination = templatePath(root, template.id);
  const temporary = `${destination}.${randomUUID()}.tmp`;
  await writeFile(temporary, `${JSON.stringify(template, null, 2)}\n`, "utf8");
  await rm(destination, { force: true });
  await rename(temporary, destination);
}

async function readTemplate(root: string, id: string): Promise<OfficeWorkflowTemplate | undefined> {
  validateId(id);
  try {
    const template = JSON.parse(await readFile(templatePath(root, id), "utf8")) as OfficeWorkflowTemplate;
    if (template.id !== id || !template.name || !Array.isArray(template.steps)) throw new Error("Office 工作流模板已损坏");
    return template;
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") return undefined;
    throw error;
  }
}

function templateRoot(root: string): string {
  return path.join(path.resolve(root), "templates");
}

function templatePath(root: string, id: string): string {
  validateId(id);
  return path.join(templateRoot(root), `${id}.json`);
}

function validateId(id: string): void {
  if (!isUuid(id)) throw new Error("Office 工作流模板 ID 无效");
}

function isUuid(value: string): boolean {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(value);
}
