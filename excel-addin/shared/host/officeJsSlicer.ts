/**
 * Office.js slicer.list / create / delete (ExcelApi 1.10).
 */
import {
  applyWritableSlicerFields,
  queueLoadSlicer,
  readSlicerSnapshot,
  withSlicerExcel,
} from "./officeJsSlicerShared";
import type {
  ExcelRequestContextWithSlicer,
  ExcelSlicer,
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

function getWorkbookSlicers(context: ExcelRequestContextWithSlicer): ExcelSlicerCollection {
  const coll = context.workbook.slicers;
  if (!coll || typeof coll.load !== "function") {
    throw new Error("Workbook.slicers is not available");
  }
  return coll;
}

async function resolveSource(
  context: ExcelRequestContextWithSlicer,
  input: SlicerCreateInput,
): Promise<object> {
  if (input.sourceType === "table") {
    const tables = context.workbook.tables;
    if (tables && typeof tables.getItem === "function") {
      const table = tables.getItem(input.sourceName) as ExcelTableLike;
      if (table && typeof table.load === "function") table.load("name");
      await context.sync();
      return table as object;
    }
    throw new Error("Workbook.tables is not available to resolve table source");
  }
  const pivots = context.workbook.pivotTables;
  if (pivots && typeof pivots.getItem === "function") {
    const pivot = pivots.getItem(input.sourceName);
    if (pivot && typeof pivot.load === "function") pivot.load("name");
    await context.sync();
    return pivot as object;
  }
  throw new Error("Workbook.pivotTables is not available to resolve pivotTable source");
}

export async function officeJsListSlicers(
  input: SlicerListInput = {},
): Promise<HostResult<SlicerListInfo>> {
  return withSlicerExcel("slicer.list", async (context) => {
    const coll = getWorkbookSlicers(context);
    coll.load("items");
    await context.sync();

    const items = Array.isArray(coll.items) ? coll.items : [];
    for (const s of items) queueLoadSlicer(s);
    await context.sync();

    const limitations = [
      NO_SOURCE_READBACK,
      "Filter item details use slicer.filter.get; list only includes scalar layout snapshot",
    ];
    let slicers: SlicerInfo[] = items.map((s) => readSlicerSnapshot(s));
    if (input.sheetName) {
      const want = input.sheetName;
      slicers = slicers.filter((s) => s.sheetName === want);
    }
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
    destSheet.load("name");
    await context.sync();

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

    const baseLimitations = [NO_SOURCE_READBACK];
    const snapshot = readSlicerSnapshot(slicer, baseLimitations);
    const expectedName = input.name?.trim() || snapshot.name;
    const checks: SlicerCreateInfo["verification"]["checks"] = [];
    const objectExists = Boolean(snapshot.id || snapshot.name);
    checks.push({ name: "objectExists", ok: objectExists, message: snapshot.id });
    const nameMatches = !input.name || snapshot.name === input.name;
    checks.push({ name: "name", ok: nameMatches, message: snapshot.name });
    const sheetOk = snapshot.sheetName === String(destSheet.name);
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

    const verificationOk = checks.every((c) => c.ok);
    if (!verificationOk) {
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
        nameMatches: nameMatches || expectedName === snapshot.name,
        checks,
      },
    };
  });
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
    const deletedName = String(slicer.name);
    if (typeof slicer.delete !== "function") {
      throw new Error("Slicer.delete is not available");
    }
    slicer.delete();
    await context.sync();

    // Confirm gone
    if (typeof coll.getItemOrNullObject === "function") {
      const gone = coll.getItemOrNullObject(input.name);
      gone.load("name");
      await context.sync();
      if (gone.isNullObject !== true) {
        // Some hosts need id; also try collection scan
        coll.load("items");
        await context.sync();
        const still = (coll.items ?? []).some(
          (s: ExcelSlicer) => String(s.name) === deletedName || String(s.name) === input.name,
        );
        if (still) throw new Error(`slicer still present after delete: ${input.name}`);
      }
    } else {
      coll.load("items");
      await context.sync();
      for (const s of coll.items ?? []) queueLoadSlicer(s);
      await context.sync();
      const still = (coll.items ?? []).some(
        (s) => String(s.name) === deletedName || String(s.name) === input.name,
      );
      if (still) throw new Error(`slicer still present after delete: ${input.name}`);
    }

    return { deleted: deletedName };
  });
}
