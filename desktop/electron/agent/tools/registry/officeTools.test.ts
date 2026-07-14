import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS, TOOL_DEFINITIONS_MAP } from "./toolDefinitions";

type ObjectToolParameters = {
  required?: string[];
  properties: Record<string, { description?: string; enum?: string[] }>;
};

function tool(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((candidate) => candidate.name === name);
  expect(definition, `missing tool definition: ${name}`).toBeDefined();
  return definition!;
}

function parameters(name: string): ObjectToolParameters {
  return tool(name).parameters as ObjectToolParameters;
}

describe("Office tool definitions", () => {
  it("keeps canonical names unique and exposes the supported surface", () => {
    const names = ALL_TOOL_DEFINITIONS.map((definition) => definition.name);

    expect(new Set(names).size).toBe(names.length);
    expect(names).toEqual(
      expect.arrayContaining([
        "workbook.inspect",
        "range.read",
        "range.write",
        "macro.write",
        "macro.run",
        "knowledge.search",
        "knowledge.write",
        "memory.write",
        "memory.delete",
        "office.connection.status",
        "office.documents.list",
        "office.workflow.run",
        "office.workflow.status",
        "office.objects.list",
        "office.transaction.inspect",
        "office.transaction.undo",
        "office.transaction.redo",
        "word.open",
        "word.insertHeading",
        "word.replaceText",
        "presentation.open",
        "presentation.addSlide",
        "presentation.replaceText",
        "office.action.inspect",
        "office.action.apply",
        "office.action.validate",
      ]),
    );
    expect(names).not.toEqual(
      expect.arrayContaining([
        "range_read",
        "office_action_apply",
        "script.execute",
        "vba.writeModule",
        "office.file.inspect",
        "office.visual.snapshot",
        "shell.execute",
        "python.execute",
      ]),
    );
  });

  it("requires an explicit action and file path for Office writes", () => {
    const apply = tool("office.action.apply");
    const applyParameters = parameters("office.action.apply");

    expect(applyParameters.required).toEqual(["app", "action", "operation", "filePath"]);
    expect(applyParameters.properties.action.description).toContain("必填");
    expect(apply.description).toContain("必须提供 filePath");
    expect(apply.description).toContain("当前活动窗口");
    expect(applyParameters.properties.operation.description).toContain("snapshot");
    expect(parameters("office.action.inspect").properties.operation.description).not.toContain(
      "snapshot",
    );
  });

  it("documents linked reports, resumable workflows, and transactions", () => {
    const apply = parameters("office.action.apply");
    const inspect = parameters("office.action.inspect");

    expect(inspect.properties.operation.description).toContain("inspectLinkedOfficeContent");
    expect(apply.properties.operation.description).toContain("refreshLinkedOfficeContent");
    expect(apply.properties.operation.description).toContain("exportRangeToWord");
    expect(apply.properties.params.description).toContain("linked:true");
    for (const name of [
      "office.workflow.cancel",
      "office.workflow.template.save",
      "office.objects.activate",
      "office.transaction.list",
    ])
      expect(tool(name)).toBeDefined();
  });

  it("exposes dynamic-array spill expansion on range.read", () => {
    const read = tool("range.read");
    const readParameters = parameters("range.read");

    expect(readParameters.properties.expand.enum).toEqual([
      "none",
      "spill",
      "currentArray",
      "currentRegion",
    ]);
    expect(read.description).toContain('expand:"spill"');
    expect(read.description).toContain("省略 expand 时会自动探测");
  });

  it("uses one language-parameterized tool for internal macro code", () => {
    const write = tool("macro.write");

    expect(parameters("macro.write").required).toEqual(["language", "code", "entryPoint"]);
    expect(parameters("macro.write").properties.language.enum).toEqual(["vba", "javascript"]);
    expect(write.description).toContain("内部宏工程");
    expect(write.description).not.toContain("Python");
    expect(parameters("macro.run").properties.language.enum).toEqual(["vba"]);
  });

  it("keeps knowledge search focused on stored business context", () => {
    const search = tool("knowledge.search");
    const searchParameters = parameters("knowledge.search");

    expect(search.description).toContain("项目资料、业务口径、模板规范和历史规则");
    expect(search.description).toContain("仅在任务依赖已沉淀知识");
    expect(searchParameters.properties.query.description).toContain("结构化搜索词");
    expect(searchParameters.properties.scope).toBeUndefined();
  });

  it("keeps public memory schemas stable", () => {
    expect(parameters("memory.write").properties.kind.enum).toEqual([
      "preference",
      "constraint",
      "correction",
      "style_preference",
      "operation_preference",
      "file_impression",
    ]);
    expect(parameters("memory.delete").required).toEqual(["memoryId"]);
    expect(tool("memory.delete").description).toContain("memory.list");
  });

  it("maps common underscore aliases to canonical definitions", () => {
    expect(TOOL_DEFINITIONS_MAP.get("range_read")?.name).toBe("range.read");
    expect(TOOL_DEFINITIONS_MAP.get("range_read")?.riskLevel).toBe("safe");
    expect(TOOL_DEFINITIONS_MAP.get("office.action_apply")?.name).toBe("office.action.apply");
    expect(TOOL_DEFINITIONS_MAP.get("office_action_apply")?.name).toBe("office.action.apply");
  });
});
