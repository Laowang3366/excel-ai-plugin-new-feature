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
    filePath: "C:/documents/template.docx",
    params,
  };
}

describe("Word mail merge and content control parameter schemas", () => {
  it("accepts content-control inspection validation", () => {
    const args = {
      app: "word",
      operation: "inspectContentControls",
      filePath: "C:/documents/template.docx",
      params: { host: "word", countPath: "controlCount", minCount: 1 },
    };
    for (const tool of ["office.action.inspect", "office.action.validate"]) {
      expect(
        parseAndValidateToolArguments(JSON.stringify(args), parameters(tool)).error,
      ).toBeUndefined();
    }
  });

  it("models the actual append-only mail merge template fields", () => {
    const schema = parameters("office.action.apply");
    expect(
      parseAndValidateToolArguments(
        JSON.stringify(
          apply("prepareMailMergeTemplate", {
            fields: ["CustomerName", { name: "InvoiceNumber" }],
          }),
        ),
        schema,
      ).error,
    ).toBeUndefined();

    for (const params of [
      { fields: [] },
      { fields: [{ placeholder: "{{name}}", field: "CustomerName" }] },
      { fields: [{ name: "CustomerName", format: "upper" }] },
    ]) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(apply("prepareMailMergeTemplate", params)),
          schema,
        ).error,
      ).toBeDefined();
    }
  });

  it("accepts bounded content-control values and rejects invented mappings", () => {
    const schema = parameters("office.action.apply");
    const valid = apply("populateContentControls", {
      dateFormat: "yyyy-MM-dd",
      values: {
        customer: "Contoso",
        approved: true,
        amount: 1200.5,
        dueDate: { value: "2026-08-01", dateFormat: "yyyy年M月d日" },
      },
    });
    expect(parseAndValidateToolArguments(JSON.stringify(valid), schema).error).toBeUndefined();

    for (const params of [
      { values: {} },
      { values: { customer: { value: "Contoso", format: "bold" } } },
      { values: { customer: "Contoso" }, fieldMap: { customer: "客户" } },
      { values: { picture: { path: "C:/logo.png", width: 100 } } },
    ]) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(apply("populateContentControls", params)),
          schema,
        ).error,
      ).toBeDefined();
    }
  });

  it("accepts implemented mail merge inputs in direct and workflow calls", () => {
    const args = apply("batchMailMerge", {
      dataSourcePath: "C:/data/customers.xlsx",
      outputFormat: "both",
      outputDirectory: "C:/output/contracts",
      fileNamePattern: "{CustomerName}-{index}",
      conditions: [
        {
          placeholder: "{{vipText}}",
          field: "Tier",
          operator: "eq",
          value: "VIP",
          trueText: "Priority",
          falseText: "Standard",
        },
      ],
      imageFields: [{ placeholder: "{{logo}}", field: "LogoPath", width: 120 }],
    });

    expect(
      parseAndValidateToolArguments(JSON.stringify(args), parameters("office.action.apply")).error,
    ).toBeUndefined();
    expect(
      parseAndValidateToolArguments(
        JSON.stringify({ steps: [args] }),
        parameters("office.workflow.run"),
      ).error,
    ).toBeUndefined();
  });

  it("rejects unsupported mail merge sources and options", () => {
    const schema = parameters("office.action.apply");
    const invalid = [
      apply("mailMerge", { dataSourcePath: "C:/data/customers.csv" }),
      apply("mailMerge", {
        dataSourcePath: "C:/data/customers.xlsx",
        outputFormat: "txt",
      }),
      apply("mailMerge", {
        dataSourcePath: "C:/data/customers.xlsx",
        outputDirectory: "C:/output",
      }),
      apply("batchMailMerge", {
        dataSourcePath: "C:/data/customers.xlsx",
        fileNameField: "CustomerName",
      }),
      apply("batchMailMerge", {
        dataSourcePath: "C:/data/customers.xlsx",
        firstRecord: 1,
        lastRecord: 10,
      }),
      apply("batchMailMerge", {
        dataSourcePath: "C:/data/customers.xlsx",
        imageFields: [{ placeholder: "{{logo}}", field: "LogoPath", height: 80 }],
      }),
    ];

    for (const args of invalid) {
      expect(parseAndValidateToolArguments(JSON.stringify(args), schema).error).toBeDefined();
    }
  });

  it("supports single and batch content-control creation by real type", () => {
    const schema = parameters("office.action.apply");
    const valid = [
      apply("manageContentControls", {
        command: "add",
        type: "text",
        title: "Customer",
        tag: "customer",
        placeholder: "Enter customer",
        start: 0,
        end: 0,
      }),
      apply("manageContentControls", {
        command: "add",
        type: "dropdown",
        tag: "status",
        entries: [
          { text: "Draft", value: "draft" },
          { text: "Approved", value: "approved" },
        ],
      }),
      apply("manageContentControls", {
        command: "add",
        controls: [
          { type: "checkbox", tag: "approved", title: "Approved" },
          { type: "date", tag: "dueDate", placeholder: "yyyy-MM-dd" },
        ],
      }),
    ];

    for (const args of valid) {
      expect(parseAndValidateToolArguments(JSON.stringify(args), schema).error).toBeUndefined();
    }

    for (const params of [
      { command: "add", type: "toggle", tag: "approved" },
      { command: "add", type: "text", entries: [{ text: "Unexpected" }] },
      { command: "add", type: "date", defaultValue: "2026-08-01" },
      { command: "add", controls: [{ type: "dropdown", entries: [{ value: "missing text" }] }] },
    ]) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(apply("manageContentControls", params)),
          schema,
        ).error,
      ).toBeDefined();
    }
  });

  it("requires deterministic selectors for delete and update", () => {
    const schema = parameters("office.action.apply");
    const valid = [
      { command: "delete", id: "123", deleteContents: false },
      { command: "delete", tag: "obsolete", deleteContents: true },
      { command: "update", id: "123", title: "Customer", tag: "customer" },
      { command: "update", title: "Customer", lockContents: true },
      { command: "update", tag: "customer", lockContents: true, lockControl: false },
    ];
    for (const params of valid) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(apply("manageContentControls", params)),
          schema,
        ).error,
      ).toBeUndefined();
    }

    const invalid = [
      { command: "delete" },
      { command: "delete", id: "123", tag: "customer" },
      { command: "update", id: "123" },
      { command: "update", title: "Customer", tag: "renamed" },
      { command: "setValue", tag: "customer", value: "Contoso" },
      { command: "addListEntry", tag: "status", text: "Approved" },
    ];
    for (const params of invalid) {
      expect(
        parseAndValidateToolArguments(
          JSON.stringify(apply("manageContentControls", params)),
          schema,
        ).error,
      ).toBeDefined();
    }
  });
});
