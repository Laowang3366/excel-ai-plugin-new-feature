import type { ExcelRequestContext } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import type {
  HostResult,
  NamedRangeInfo,
  NamedRangeScope,
  SheetProtectionInfo,
  SheetVisibility,
  SheetVisibilityInfo,
} from "./types";

function mapVisibility(raw: string | undefined | null): SheetVisibility {
  const v = String(raw ?? "").toLowerCase().replace(/\s+/g, "");
  if (v.includes("veryhidden")) return "veryHidden";
  if (v.includes("hidden")) return "hidden";
  return "visible";
}

function toOfficeVisibility(v: SheetVisibility): string {
  if (v === "veryHidden") return "VeryHidden";
  if (v === "hidden") return "Hidden";
  return "Visible";
}

function readNamedItem(
  item: {
    name: string;
    formula?: string;
    visible?: boolean;
  },
  scope: NamedRangeScope,
  sheetName?: string,
): NamedRangeInfo {
  return {
    name: item.name,
    refersTo: item.formula ?? "",
    scope,
    // workbook-scoped results never carry sheetName
    sheetName: scope === "worksheet" ? sheetName : undefined,
    visible: item.visible,
  };
}

export async function officeJsGetSheetVisibility(
  sheetName: string,
): Promise<HostResult<SheetVisibilityInfo>> {
  return withExcel("sheet.visibility.get", async (context: ExcelRequestContext) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name,visibility");
    await context.sync();
    return { sheetName: sheet.name, visibility: mapVisibility(sheet.visibility) };
  });
}

export async function officeJsSetSheetVisibility(
  sheetName: string,
  visibility: SheetVisibility,
): Promise<HostResult<SheetVisibilityInfo>> {
  return withExcel("sheet.visibility.set", async (context: ExcelRequestContext) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.visibility = toOfficeVisibility(visibility);
    sheet.load("name,visibility");
    await context.sync();
    return { sheetName: sheet.name, visibility: mapVisibility(sheet.visibility) };
  });
}

export async function officeJsGetSheetProtection(
  sheetName: string,
): Promise<HostResult<SheetProtectionInfo>> {
  return withExcel("sheet.protection.get", async (context: ExcelRequestContext) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name");
    sheet.protection.load("protected");
    await context.sync();
    return { sheetName: sheet.name, protected: sheet.protection.protected === true };
  });
}

export async function officeJsProtectSheet(
  sheetName: string,
  password?: string,
): Promise<HostResult<SheetProtectionInfo>> {
  return withExcel("sheet.protection.protect", async (context: ExcelRequestContext) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name");
    sheet.protection.load("protected");
    await context.sync();
    if (sheet.protection.protected) {
      throw new Error("sheet is already protected");
    }
    // Official: protect(options?, password?) — password is 2nd arg, request-scoped only.
    if (password != null && password !== "") {
      sheet.protection.protect({}, password);
    } else {
      sheet.protection.protect();
    }
    sheet.protection.load("protected");
    await context.sync();
    return { sheetName: sheet.name, protected: Boolean(sheet.protection.protected) };
  });
}

export async function officeJsUnprotectSheet(
  sheetName: string,
  password?: string,
): Promise<HostResult<SheetProtectionInfo>> {
  return withExcel("sheet.protection.unprotect", async (context: ExcelRequestContext) => {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.load("name");
    sheet.protection.load("protected");
    await context.sync();
    if (!sheet.protection.protected) {
      throw new Error("sheet is not protected");
    }
    if (password != null && password !== "") {
      sheet.protection.unprotect(password);
    } else {
      sheet.protection.unprotect();
    }
    sheet.protection.load("protected");
    await context.sync();
    return { sheetName: sheet.name, protected: Boolean(sheet.protection.protected) };
  });
}

function namesCollection(context: ExcelRequestContext, scope: NamedRangeScope, sheetName?: string) {
  if (scope === "worksheet") {
    if (!sheetName) throw new Error("sheetName is required for worksheet scope");
    return context.workbook.worksheets.getItem(sheetName).names;
  }
  return context.workbook.names;
}

export async function officeJsListNamedRanges(input?: {
  scope?: NamedRangeScope;
  sheetName?: string;
}): Promise<HostResult<NamedRangeInfo[]>> {
  return withExcel("namedRange.list", async (context: ExcelRequestContext) => {
    const scope = input?.scope ?? "workbook";
    const names = namesCollection(context, scope, input?.sheetName);
    names.load("items/name,items/formula,items/visible");
    await context.sync();
    return names.items.map((item) => readNamedItem(item, scope, input?.sheetName));
  });
}

export async function officeJsCreateNamedRange(input: {
  name: string;
  refersTo: string;
  scope: NamedRangeScope;
  sheetName?: string;
  visible?: boolean;
}): Promise<HostResult<NamedRangeInfo>> {
  return withExcel("namedRange.create", async (context: ExcelRequestContext) => {
    const names = namesCollection(context, input.scope, input.sheetName);
    const formula = input.refersTo.startsWith("=") ? input.refersTo : `=${input.refersTo}`;
    const item = names.add(input.name, formula);
    if (input.visible != null) item.visible = input.visible;
    item.load("name,formula,visible");
    await context.sync();
    return readNamedItem(item, input.scope, input.sheetName);
  });
}

export async function officeJsUpdateNamedRange(input: {
  name: string;
  scope: NamedRangeScope;
  sheetName?: string;
  newName?: string;
  refersTo?: string;
  visible?: boolean;
}): Promise<HostResult<NamedRangeInfo>> {
  return withExcel("namedRange.update", async (context: ExcelRequestContext) => {
    const names = namesCollection(context, input.scope, input.sheetName);
    const item = names.getItem(input.name);
    item.load("name,formula,visible");
    await context.sync();
    // NamedItem.name is readonly. Rename: add new first (fail keeps old), then delete old.
    if (input.newName != null) {
      const newName = input.newName.trim();
      if (newName === "") throw new Error("newName must be non-empty");
      const sameName =
        newName.localeCompare(item.name, undefined, { sensitivity: "accent" }) === 0;
      if (!sameName) {
        const oldName = item.name;
        names.load("items/name");
        await context.sync();
        const conflict = names.items.some(
          (entry) => entry.name.localeCompare(newName, undefined, { sensitivity: "accent" }) === 0,
        );
        if (conflict) {
          throw new Error(`named range already exists: ${newName}`);
        }
        const formula =
          input.refersTo != null
            ? input.refersTo.startsWith("=")
              ? input.refersTo
              : `=${input.refersTo}`
            : item.formula;
        const visible = input.visible ?? item.visible;
        // Add first so a failed names.add leaves the original name intact.
        const created = names.add(newName, formula);
        created.visible = visible;
        created.load("name,formula,visible");
        await context.sync();
        const old = names.getItem(oldName);
        old.delete();
        await context.sync();
        return readNamedItem(created, input.scope, input.sheetName);
      }
      // sameName after trim: no-op rename, continue formula/visible updates.
    }
    if (input.refersTo != null) {
      item.formula = input.refersTo.startsWith("=") ? input.refersTo : `=${input.refersTo}`;
    }
    if (input.visible != null) item.visible = input.visible;
    item.load("name,formula,visible");
    await context.sync();
    return readNamedItem(item, input.scope, input.sheetName);
  });
}

export async function officeJsDeleteNamedRange(input: {
  name: string;
  scope: NamedRangeScope;
  sheetName?: string;
}): Promise<HostResult<{ deleted: string }>> {
  return withExcel("namedRange.delete", async (context: ExcelRequestContext) => {
    const names = namesCollection(context, input.scope, input.sheetName);
    const item = names.getItem(input.name);
    item.delete();
    await context.sync();
    return { deleted: input.name };
  });
}

