import { describe, expect, it } from "vitest";

import { ALL_TOOL_DEFINITIONS } from "./toolDefinitions";
import { parseAndValidateToolArguments } from "./toolSchema";

function parameters(name: string) {
  const definition = ALL_TOOL_DEFINITIONS.find((tool) => tool.name === name);
  if (!definition) throw new Error(`missing tool definition: ${name}`);
  return definition.parameters;
}

function apply(operation: string, params: Record<string, unknown>) {
  return {
    app: "word",
    action: "edit",
    operation,
    filePath: "C:/documents/report.docx",
    params,
  };
}

describe("Word reference and review parameter schemas", () => {
  it("accepts reference inspection and all Worker-supported reference commands", () => {
    const inspect = {
      app: "word",
      operation: "inspectReferences",
      filePath: "C:/documents/report.docx",
      params: { host: "word", countPath: "bookmarkCount", minCount: 1 },
    };
    expect(
      parseAndValidateToolArguments(JSON.stringify(inspect), parameters("office.action.inspect"))
        .error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(JSON.stringify(inspect), parameters("office.action.validate"))
        .error,
    ).toBeUndefined();

    const cases = [
      { command: "createBookmark", name: "Summary", start: 0, end: 10 },
      { command: "addBookmark", name: "Details", bookmark: "Summary" },
      { command: "deleteBookmark", name: "Obsolete" },
      { command: "addFootnote", text: "Source note" },
      { command: "addEndnote", text: "Appendix note" },
      { command: "addCaption", label: "Figure", title: "Revenue" },
      { command: "addCrossReference", referenceType: "bookmark", item: "Summary" },
      { command: "addTableOfFigures", label: "Figure" },
      { command: "updateFields" },
    ];

    for (const params of cases) {
      const args = apply("manageReferences", params);
      expect(
        parseAndValidateToolArguments(JSON.stringify(args), parameters("office.action.apply"))
          .error,
      ).toBeUndefined();
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({ steps: [args] }),
          parameters("office.workflow.run"),
        ).error,
      ).toBeUndefined();
    }
  });

  it("rejects undocumented reference commands and cross-command fields", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      { command: "updateAll" },
      { command: "addBookmark" },
      { command: "addFootnote", text: "note", targetType: "table" },
      { command: "addCrossReference", item: "Summary", rule: { authors: ["AI"] } },
      { command: "updateFields", name: "unexpected" },
    ];

    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("manageReferences", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("accepts only the implemented revision management commands", () => {
    const schema = parameters("office.action.apply");
    const valid = [
      { command: "acceptAll" },
      { command: "rejectAll" },
      { command: "accept", author: "Reviewer", revisionType: 1 },
      { command: "reject", revisionType: 2 },
      { command: "track", enabled: true },
    ];
    for (const params of valid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("manageRevisions", params)), schema)
          .error,
      ).toBeUndefined();
    }

    for (const params of [
      { command: "acceptMatching", rule: { authors: ["Reviewer"] } },
      { command: "deleteComments" },
      { command: "track" },
      { command: "accept", revisionType: 100 },
    ]) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("manageRevisions", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("supports the canonical and legacy compare path without accepting fictional options", () => {
    const schema = parameters("office.action.apply");
    for (const params of [
      { comparePath: "C:/documents/revised.docx" },
      { revisedFilePath: "C:/documents/revised.docx" },
    ]) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify({
            ...apply("compareDocuments", params),
            outputPath: "C:/documents/comparison.docx",
          }),
          schema,
        ).error,
      ).toBeUndefined();
    }

    for (const params of [
      {},
      { comparePath: "C:/revised.docx", revisedFilePath: "C:/other.docx" },
      { comparePath: "C:/revised.docx", author: "AI" },
      { comparePath: "C:/revised.docx", granularity: "paragraph" },
    ]) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("compareDocuments", params)), schema)
          .error,
      ).toBeDefined();
    }
  });

  it("accepts bounded tracked-change variants and the legacy edits alias", () => {
    const changes = [
      { command: "insert", text: "Introduction", position: "start" },
      { command: "replace", find: "draft", replace: "final", matchCase: true, all: false },
      { command: "delete", find: "obsolete", all: true },
      { command: "replaceBookmark", name: "Summary", text: "Updated summary" },
      { command: "replaceContentControl", tag: "customer", text: "Contoso" },
      { command: "replaceContentControl", title: "Approver", text: "Reviewer" },
      {
        command: "replaceContentControl",
        tag: "status",
        title: "Status",
        text: "Approved",
      },
    ];
    const schema = parameters("office.action.apply");

    for (const params of [
      { changes, keepTracking: true },
      { edits: changes, restoreTracking: false },
    ]) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("applyTrackedChanges", params)), schema)
          .error,
      ).toBeUndefined();
    }
  });

  it("rejects ambiguous or unsafe tracked-change structures", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      { changes: [{ command: "replaceContentControl", text: "matches every control" }] },
      { changes: [{ command: "replace", replace: "missing find" }] },
      { changes: [{ command: "insert", text: "x", position: -1 }] },
      { changes: [{ command: "delete", find: "x", replace: "unexpected" }] },
      { changes: [{ find: "implicit command", replace: "not allowed" }] },
      { changes: [{ command: "replaceBookmark", name: "B", text: "x", shell: "whoami" }] },
      {
        changes: [{ command: "insert", text: "x" }],
        edits: [{ command: "insert", text: "y" }],
      },
    ];

    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(apply("applyTrackedChanges", params)), schema)
          .error,
      ).toBeDefined();
    }
  });
});
