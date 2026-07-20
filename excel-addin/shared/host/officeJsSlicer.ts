/**
 * Office.js slicer.list / create / delete (ExcelApi 1.10).
 */
import {
  applyWritableSlicerFields,
  queueLoadSlicer,
  readSlicerSnapshot,
  withSlicerExcel,
} from "./officeJsSlicerShared";
import { requireHostString } from "./officeJsSlicerReadback";
import type {
  ExcelRequestContextWithSlicer,
  ExcelSlicerCollection,
  ExcelTableLike,
  ExcelWorksheetWithSlicers,
} from "./officeJsSlicerTypes";
import type {
  SlicerCreateInfo,
  SlicerCreateInput,
  SlicerDeleteInfo,
  SlicerDeleteInput,
  SlicerInfo,
  SlicerListInfo,
  SlicerListInput,
} from "./slicerTypes";
import type { HostResult } from "./types";
import { fail } from "./types";

const NO_SOURCE_READBACK =
  "Stable Excel.Slicer has no sourceName/sourceType/sourceField readback; requestedSource is create input only and cannot be host-verified";

const FIELD_NOTE =
  "sourceField: Table column name/ID or PivotField ID per Office.js add(); host resolves strings; failure is ordinary";

function getWorkbookSlicers(context: ExcelRequestContextWithSlicer): ExcelSlicerCollection {
  const coll = context.workbook.slicers;
  if (!coll || typeof coll.load !== "function") {
    throw new Error("Workbook.slicers is not available");
  }
  return coll;
}

async function loadSlicerCollection(
  context: ExcelRequestContextWithSlicer,
  coll: ExcelSlicerCollection,
): Promise<SlicerInfo[]> {
  coll.load("items");
  await context.sync();
  const items = Array.isArray(coll.items) ? coll.items : null;
  if (!items) throw new Error("SlicerCollection.items must be an array after load");
  for (const s of items) queueLoadSlicer(s);
  await context.sync();
  return items.map((s) => readSlicerSnapshot(s, [NO_SOURCE_READBACK]));
}

async function resolveSource(
  context: ExcelRequestContextWithSlicer,
  input: SlicerCreateInput,
): Promise<object> {
  if (input.sourceType === "table") {
    const tables = context.workbook.tables;
    if (!tables || typeof tables.getItem !== "function") {
      throw new Error("Workbook.tables is not available to resolve table source");
    }
    const table = tables.getItem(input.sourceName) as ExcelTableLike;
    if (!table || typeof table.load !== "function") {
      throw new Error("Table.load is not available");
    }
    table.load("name");
    await context.sync();
    const name = requireHostString("table.name", table.name);
    if (name !== input.sourceName) {
      throw new Error(`table name mismatch: requested=${input.sourceName} got=${name}`);
    }
    return table as object;
  }
  const pivots = context.workbook.pivotTables;
  if (!pivots || typeof pivots.getItem !== "function") {
    throw new Error("Workbook.pivotTables is not available to resolve pivotTable source");
  }
  const pivot = pivots.getItem(input.sourceName);
  if (!pivot || typeof pivot.load !== "function") {
    throw new Error("PivotTable.load is not available");
  }
  pivot.load("name");
  await context.sync();
  const name = requireHostString("pivotTable.name", pivot.name);
  if (name !== input.sourceName) {
    throw new Error(`pivotTable name mismatch: requested=${input.sourceName} got=${name}`);
  }
  return pivot as object;
}

export async function officeJsListSlicers(
  input: SlicerListInput = {},
): Promise<HostResult<SlicerListInfo>> {
  return withSlicerExcel("slicer.list", async (context) => {
    const limitations = [
      NO_SOURCE_READBACK,
      "Filter item details use slicer.filter.get; list only includes scalar layout snapshot",
    ];

    if (input.sheetName !== undefined) {
      const sheet = context.workbook.worksheets.getItem(
        input.sheetName,
      ) as ExcelWorksheetWithSlicers;
      if (typeof sheet.load !== "function") {
        throw new Error("Worksheet.load is not available");
      }
      sheet.load("name");
      await context.sync();
      const sheetName = requireHostString("worksheet.name", sheet.name);
      if (sheetName !== input.sheetName) {
        throw new Error(`sheet name mismatch: requested=${input.sheetName} got=${sheetName}`);
      }
      if (!sheet.slicers || typeof sheet.slicers.load !== "function") {
        throw new Error("Worksheet.slicers is not available");
      }
      const slicers = await loadSlicerCollection(context, sheet.slicers);
      slicers.sort((a, b) => a.name.localeCompare(b.name));
      return { slicers, limitations };
    }

    const slicers = await loadSlicerCollection(context, getWorkbookSlicers(context));
    slicers.sort((a, b) => {
      const sheet = a.sheetName.localeCompare(b.sheetName);
      return sheet !== 0 ? sheet : a.name.localeCompare(b.name);
    });
    return { slicers, limitations };
  });
}

export async function officeJsCreateSlicer(
  input: SlicerCreateInput,
): Promise<HostResult<SlicerCreateInfo>> {
  if (input.advancedIntent !== "interactive-pivot") {
    return fail(
      "slicer.create",
      "office-js",
      "advancedIntent must be interactive-pivot",
    );
  }
  if (input.sourceType !== "table" && input.sourceType !== "pivotTable") {
    return fail("slicer.create", "office-js", "sourceType must be table or pivotTable");
  }

  return withSlicerExcel("slicer.create", async (context) => {
    const source = await resolveSource(context, input);
    const destSheet = context.workbook.worksheets.getItem(
      input.destinationSheet,
    ) as ExcelWorksheetWithSlicers;
    if (typeof destSheet.load !== "function") {
      throw new Error("Worksheet.load is not available");
    }
    destSheet.load("name");
    await context.sync();
    const destName = requireHostString("destinationSheet.name", destSheet.name);
    if (destName !== input.destinationSheet) {
      throw new Error(
        `destination sheet name mismatch: requested=${input.destinationSheet} got=${destName}`,
      );
    }

    const coll = getWorkbookSlicers(context);
    if (typeof coll.add !== "function") {
      throw new Error("Workbook.slicers.add is not available");
    }

    const slicer = coll.add(source, input.sourceField, destSheet);
    if (!slicer) throw new Error("slicers.add returned no object");

    applyWritableSlicerFields(slicer, {
      name: input.name,
      caption: input.caption,
      top: input.top,
      left: input.left,
      width: input.width,
      height: input.height,
      style: input.style,
      sortBy: input.sortBy,
    });

    await context.sync();
    queueLoadSlicer(slicer);
    await context.sync();

    const baseLimitations = [NO_SOURCE_READBACK, FIELD_NOTE];
    const snapshot = readSlicerSnapshot(slicer, baseLimitations);
    const checks: SlicerCreateInfo["verification"]["checks"] = [];
    const objectExists = true;
    checks.push({ name: "objectExists", ok: objectExists, message: snapshot.id });
    const nameMatches = !input.name || snapshot.name === input.name;
    checks.push({ name: "name", ok: nameMatches, message: snapshot.name });
    const sheetOk = snapshot.sheetName === destName;
    checks.push({ name: "destinationSheet", ok: sheetOk, message: snapshot.sheetName });
    if (input.caption !== undefined) {
      checks.push({
        name: "caption",
        ok: snapshot.caption === input.caption,
        message: snapshot.caption,
      });
    }
    if (input.top !== undefined) {
      checks.push({ name: "top", ok: snapshot.top === input.top, message: String(snapshot.top) });
    }
    if (input.left !== undefined) {
      checks.push({
        name: "left",
        ok: snapshot.left === input.left,
        message: String(snapshot.left),
      });
    }
    if (input.width !== undefined) {
      checks.push({
        name: "width",
        ok: snapshot.width === input.width,
        message: String(snapshot.width),
      });
    }
    if (input.height !== undefined) {
      checks.push({
        name: "height",
        ok: snapshot.height === input.height,
        message: String(snapshot.height),
      });
    }
    if (input.style !== undefined) {
      checks.push({
        name: "style",
        ok: snapshot.style === input.style,
        message: snapshot.style,
      });
    }
    if (input.sortBy !== undefined) {
      checks.push({
        name: "sortBy",
        ok: snapshot.sortBy === input.sortBy,
        message: snapshot.sortBy,
      });
    }

    if (!checks.every((c) => c.ok)) {
      throw new Error(
        `slicer.create readback mismatch: ${checks
          .filter((c) => !c.ok)
          .map((c) => c.name)
          .join(", ")}`,
      );
    }

    return {
      ...snapshot,
      requestedSource: {
        sourceType: input.sourceType,
        sourceName: input.sourceName,
        sourceField: input.sourceField,
      },
      verification: {
        ok: true,
        objectExists,
        nameMatches,
        checks,
      },
    };
  });
}

async function confirmSlicerGone(
  context: ExcelRequestContextWithSlicer,
  coll: ExcelSlicerCollection,
  names: string[],
): Promise<void> {
  const want = new Set(names);
  if (typeof coll.getItemOrNullObject === "function") {
    const gone = coll.getItemOrNullObject(names[0]!);
    if (typeof gone.load === "function") gone.load("name");
    await context.sync();
    if (gone.isNullObject === true) return;
  }

  if (typeof coll.load !== "function") {
    throw new Error("SlicerCollection.load is not available for delete confirmation");
  }
  coll.load("items");
  await context.sync();
  const items = Array.isArray(coll.items) ? coll.items : null;
  if (!items) throw new Error("SlicerCollection.items must be an array after load");
  for (const s of items) {
    if (typeof s.load !== "function") throw new Error("Slicer.load is not available");
    s.load("name");
  }
  await context.sync();
  for (const s of items) {
    const n = requireHostString("slicer.name", s.name);
    if (want.has(n)) {
      throw new Error(`slicer still present after delete: ${n}`);
    }
  }
}

export async function officeJsDeleteSlicer(
  input: SlicerDeleteInput,
): Promise<HostResult<SlicerDeleteInfo>> {
  return withSlicerExcel("slicer.delete", async (context) => {
    const coll = getWorkbookSlicers(context);
    if (typeof coll.getItem !== "function") {
      throw new Error("Workbook.slicers.getItem is not available");
    }
    const slicer = coll.getItem(input.name);
    queueLoadSlicer(slicer);
    await context.sync();
    const deletedName = requireHostString("slicer.name", slicer.name);
    if (typeof slicer.delete !== "function") {
      throw new Error("Slicer.delete is not available");
    }
    slicer.delete();
    await context.sync();
    await confirmSlicerGone(context, coll, [deletedName, input.name]);
    return { deleted: deletedName };
  });
}
