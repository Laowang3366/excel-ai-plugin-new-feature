import { requireWorkbook } from "./wpsJsaRuntime";
import type { HostResult, NamedRangeInfo, NamedRangeScope } from "./types";
import { fail, ok, unsupported } from "./types";
import {
  NAMED_RANGE_EVIDENCE as EVIDENCE,
  findByName,
  hasName,
  messageOf,
  normalizeRefersTo,
  readNamedRangeInfo as readInfo,
  requireNamesCollection,
  sameName,
} from "./wpsJsaNamedRangeHelpers";

export async function wpsListNamedRanges(input?: {
  scope?: NamedRangeScope;
  sheetName?: string;
}): Promise<HostResult<NamedRangeInfo[]>> {
  const scope = input?.scope ?? "workbook";
  const workbookResult = requireWorkbook("namedRange.list");
  if (!workbookResult.ok) return workbookResult;
  const namesResult = requireNamesCollection(
    "namedRange.list",
    workbookResult.data,
    scope,
    input?.sheetName,
  );
  if (!namesResult.ok) return namesResult;
  const names = namesResult.data;
  try {
    const out: NamedRangeInfo[] = [];
    const count = names.Count ?? 0;
    for (let i = 1; i <= count; i += 1) {
      const item = names.Item?.(i);
      if (!item) continue;
      out.push(readInfo(item, scope, input?.sheetName));
    }
    return ok(out);
  } catch (error) {
    return fail("namedRange.list", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

export async function wpsCreateNamedRange(input: {
  name: string;
  refersTo: string;
  scope: NamedRangeScope;
  sheetName?: string;
  visible?: boolean;
}): Promise<HostResult<NamedRangeInfo>> {
  const name = input.name?.trim() ?? "";
  if (name === "") {
    return fail("namedRange.create", "wps-jsa", "name must be non-empty", EVIDENCE);
  }
  const workbookResult = requireWorkbook("namedRange.create");
  if (!workbookResult.ok) return workbookResult;
  const namesResult = requireNamesCollection(
    "namedRange.create",
    workbookResult.data,
    input.scope,
    input.sheetName,
  );
  if (!namesResult.ok) return namesResult;
  const names = namesResult.data;
  if (typeof names.Add !== "function") {
    return unsupported(
      "namedRange.create",
      "wps-jsa",
      "Names.Add is unavailable",
      EVIDENCE,
    );
  }
  if (hasName(names, name)) {
    return fail("namedRange.create", "wps-jsa", `named range already exists: ${name}`, EVIDENCE);
  }
  try {
    const created = names.Add(name, normalizeRefersTo(input.refersTo));
    if (!created || created.Name == null) {
      return fail("namedRange.create", "wps-jsa", "Names.Add returned empty name", EVIDENCE);
    }
    if (input.visible != null && "Visible" in created) {
      created.Visible = input.visible;
    }
    const readBack = findByName(names, name) ?? created;
    const info = readInfo(readBack, input.scope, input.sheetName);
    if (!info.name) {
      return fail("namedRange.create", "wps-jsa", "created name missing Name", EVIDENCE);
    }
    return ok(info);
  } catch (error) {
    return fail("namedRange.create", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

export async function wpsUpdateNamedRange(input: {
  name: string;
  scope: NamedRangeScope;
  sheetName?: string;
  newName?: string;
  refersTo?: string;
  visible?: boolean;
}): Promise<HostResult<NamedRangeInfo>> {
  const workbookResult = requireWorkbook("namedRange.update");
  if (!workbookResult.ok) return workbookResult;
  const namesResult = requireNamesCollection(
    "namedRange.update",
    workbookResult.data,
    input.scope,
    input.sheetName,
  );
  if (!namesResult.ok) return namesResult;
  const names = namesResult.data;
  const item = findByName(names, input.name);
  if (!item) {
    return unsupported(
      "namedRange.update",
      "wps-jsa",
      `named range not found: ${input.name}`,
      EVIDENCE,
    );
  }
  try {
    if (input.newName != null) {
      const newName = input.newName.trim();
      if (newName === "") {
        return fail("namedRange.update", "wps-jsa", "newName must be non-empty", EVIDENCE);
      }
      const currentName = String(item.Name ?? input.name);
      if (!sameName(newName, currentName)) {
        if (hasName(names, newName)) {
          return fail(
            "namedRange.update",
            "wps-jsa",
            `named range already exists: ${newName}`,
            EVIDENCE,
          );
        }
        if (typeof names.Add !== "function") {
          return unsupported(
            "namedRange.update",
            "wps-jsa",
            "Names.Add is unavailable for rename",
            EVIDENCE,
          );
        }
        const formula =
          input.refersTo != null
            ? normalizeRefersTo(input.refersTo)
            : String(item.RefersTo ?? "");
        const visible = input.visible ?? item.Visible;
        // Add first so a failed Add leaves the original name intact.
        const created = names.Add(newName, formula);
        if (!created || created.Name == null) {
          return fail(
            "namedRange.update",
            "wps-jsa",
            "rename Add failed; original name retained",
            EVIDENCE,
          );
        }
        if (visible != null && "Visible" in created) created.Visible = visible;
        if (typeof item.Delete !== "function") {
          try {
            created.Delete?.();
          } catch {
            // ignore cleanup failure
          }
          return unsupported(
            "namedRange.update",
            "wps-jsa",
            "Name.Delete is unavailable; original name retained (new name rolled back when possible)",
            EVIDENCE,
          );
        }
        try {
          item.Delete();
        } catch (error) {
          try {
            created.Delete?.();
          } catch {
            // best-effort rollback
          }
          return fail(
            "namedRange.update",
            "wps-jsa",
            `rename failed deleting original; attempted rollback of new name: ${messageOf(error)}`,
            EVIDENCE,
          );
        }
        if (hasName(names, currentName) && !sameName(currentName, newName)) {
          return fail(
            "namedRange.update",
            "wps-jsa",
            "rename incomplete: original name still present after Delete",
            EVIDENCE,
          );
        }
        if (!hasName(names, newName)) {
          return fail(
            "namedRange.update",
            "wps-jsa",
            "rename incomplete: new name missing after Add+Delete",
            EVIDENCE,
          );
        }
        const readBack = findByName(names, newName) ?? created;
        return ok(readInfo(readBack, input.scope, input.sheetName));
      }
    }
    if (input.refersTo != null) {
      if (!("RefersTo" in item)) {
        return unsupported(
          "namedRange.update",
          "wps-jsa",
          "Name.RefersTo is unavailable",
          EVIDENCE,
        );
      }
      item.RefersTo = normalizeRefersTo(input.refersTo);
    }
    if (input.visible != null) {
      if (!("Visible" in item)) {
        return unsupported(
          "namedRange.update",
          "wps-jsa",
          "Name.Visible is unavailable",
          EVIDENCE,
        );
      }
      item.Visible = input.visible;
    }
    const current = String(item.Name ?? input.name);
    const readBack = findByName(names, current) ?? item;
    return ok(readInfo(readBack, input.scope, input.sheetName));
  } catch (error) {
    return fail("namedRange.update", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

export async function wpsDeleteNamedRange(input: {
  name: string;
  scope: NamedRangeScope;
  sheetName?: string;
}): Promise<HostResult<{ deleted: string }>> {
  const workbookResult = requireWorkbook("namedRange.delete");
  if (!workbookResult.ok) return workbookResult;
  const namesResult = requireNamesCollection(
    "namedRange.delete",
    workbookResult.data,
    input.scope,
    input.sheetName,
  );
  if (!namesResult.ok) return namesResult;
  const names = namesResult.data;
  const item = findByName(names, input.name);
  if (!item) {
    return unsupported(
      "namedRange.delete",
      "wps-jsa",
      `named range not found: ${input.name}`,
      EVIDENCE,
    );
  }
  if (typeof item.Delete !== "function") {
    return unsupported(
      "namedRange.delete",
      "wps-jsa",
      "Name.Delete is unavailable",
      EVIDENCE,
    );
  }
  try {
    const deleted = String(item.Name ?? input.name);
    item.Delete();
    if (hasName(names, deleted)) {
      return fail(
        "namedRange.delete",
        "wps-jsa",
        "Delete completed but name is still present",
        EVIDENCE,
      );
    }
    return ok({ deleted });
  } catch (error) {
    return fail("namedRange.delete", "wps-jsa", messageOf(error), EVIDENCE);
  }
}
