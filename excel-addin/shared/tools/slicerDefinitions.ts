import type { ToolDefinition } from "./types";

const SOURCE_TYPES = ["table", "pivotTable"] as const;
const SORT_BY = ["dataSourceOrder", "ascending", "descending"] as const;

const nonEmptyString = {
  type: "string",
  minLength: 1,
  maxLength: 255,
} as const;

const optionalLayout = {
  top: { type: "number", minimum: 0 },
  left: { type: "number", minimum: 0 },
  width: { type: "number", exclusiveMinimum: 0 },
  height: { type: "number", exclusiveMinimum: 0 },
  style: { type: "string", minLength: 1, maxLength: 255 },
  sortBy: { type: "string", enum: [...SORT_BY] },
  caption: { type: "string", maxLength: 255 },
} as const;

export const SLICER_TOOL_DEFINITIONS: ToolDefinition[] = [
  {
    name: "slicer.list",
    description:
      "List workbook slicers (Office.js ExcelApi 1.10). Optional sheetName filter. Returns host readback: name,id,caption,sheet,layout,sortBy,style,isFilterCleared. No source readback (stable API has none).",
    riskLevel: "safe",
    parameters: {
      type: "object",
      additionalProperties: false,
      properties: {
        sheetName: { type: "string", minLength: 1, maxLength: 255 },
      },
    },
  },
  {
    name: "slicer.create",
    description:
      "Create a slicer from a Table or PivotTable (Office.js ExcelApi 1.10). Requires advancedIntent=interactive-pivot, sourceType, sourceName, sourceField, destinationSheet. sourceField is Table column name/ID or PivotField ID (host-resolved). Write→sync→load→sync; requestedSource is input-only (no host source readback).",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      additionalProperties: false,
      required: [
        "advancedIntent",
        "sourceType",
        "sourceName",
        "sourceField",
        "destinationSheet",
      ],
      properties: {
        advancedIntent: { type: "string", const: "interactive-pivot" },
        sourceType: { type: "string", enum: [...SOURCE_TYPES] },
        sourceName: nonEmptyString,
        sourceField: nonEmptyString,
        destinationSheet: nonEmptyString,
        name: nonEmptyString,
        ...optionalLayout,
      },
    },
  },
  {
    name: "slicer.update",
    description:
      "Update stable writable slicer fields (caption/name/top/left/width/height/style/sortBy). Empty update rejected. Write→sync→load→sync.",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      minProperties: 2,
      properties: {
        name: nonEmptyString,
        newName: nonEmptyString,
        ...optionalLayout,
      },
    },
  },
  {
    name: "slicer.delete",
    description: "Delete a slicer by exact name and confirm it is gone from the collection.",
    riskLevel: "dangerous",
    parameters: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        name: nonEmptyString,
      },
    },
  },
  {
    name: "slicer.filter.get",
    description:
      "Read slicer filter state: isFilterCleared, selectedKeys (official getSelectedItems), item summary. verified may be false if selection cannot be confirmed.",
    riskLevel: "safe",
    parameters: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        name: nonEmptyString,
      },
    },
  },
  {
    name: "slicer.filter.apply",
    description:
      "Select slicer items by keys via selectItems. Empty keys[] = select all (official), not select-none. Max 500 keys. Write→sync→readback.",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      additionalProperties: false,
      required: ["name", "keys"],
      properties: {
        name: nonEmptyString,
        keys: {
          type: "array",
          maxItems: 500,
          items: { type: "string", minLength: 1, maxLength: 255 },
        },
      },
    },
  },
  {
    name: "slicer.filter.clear",
    description: "Clear slicer filters (clearFilters) and verify isFilterCleared===true.",
    riskLevel: "moderate",
    parameters: {
      type: "object",
      additionalProperties: false,
      required: ["name"],
      properties: {
        name: nonEmptyString,
      },
    },
  },
];
