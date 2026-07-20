import { afterEach, describe, expect, it } from "vitest";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { sanitizeProtectionMessage } from "../shared/host/wpsJsaSheetProtection";
import {
  mapPublicToVisible,
  mapVisibleToPublic,
} from "../shared/host/wpsJsaSheetVisibility";

type NameRec = { Name: string; RefersTo: string; Visible: boolean };

function installStructureWps(options?: {
  withVisible?: boolean;
  withProtect?: boolean;
  withNames?: boolean;
  withInsertDelete?: boolean;
  protectThrows?: string;
  /** When true, Name.Delete is a no-op (simulates incomplete rename). */
  nameDeleteNoop?: boolean;
  /** When set, Name.Delete throws this message. */
  nameDeleteThrows?: string;
  /** Limit Delete fail/no-op to these names (others delete normally, for rollback). */
  nameDeleteFailFor?: string[];
}) {
  const withVisible = options?.withVisible ?? true;
  const withProtect = options?.withProtect ?? true;
  const withNames = options?.withNames ?? true;
  const withInsertDelete = options?.withInsertDelete ?? true;

  const sheets = new Map<
    string,
    {
      Name: string;
      Index: number;
      Visible: number;
      ProtectContents: boolean;
      password?: string;
    }
  >();
  sheets.set("Sheet1", {
    Name: "Sheet1",
    Index: 1,
    Visible: -1,
    ProtectContents: false,
  });
  sheets.set("Sheet2", {
    Name: "Sheet2",
    Index: 2,
    Visible: -1,
    ProtectContents: false,
  });

  const workbookNames: NameRec[] = [];

  function namesApi(list: NameRec[]) {
    return {
      get Count() {
        return list.length;
      },
      Item(indexOrName: number | string) {
        if (typeof indexOrName === "number") {
          const item = list[indexOrName - 1];
          if (!item) throw new Error("index");
          return nameProxy(item, list);
        }
        const found = list.find(
          (n) => n.Name.localeCompare(String(indexOrName), undefined, { sensitivity: "accent" }) === 0,
        );
        if (!found) throw new Error("not found");
        return nameProxy(found, list);
      },
      Add(name: string, refersTo?: string) {
        if (
          list.some(
            (n) => n.Name.localeCompare(name, undefined, { sensitivity: "accent" }) === 0,
          )
        ) {
          throw new Error("exists");
        }
        const rec: NameRec = {
          Name: name,
          RefersTo: refersTo ?? "",
          Visible: true,
        };
        list.push(rec);
        return nameProxy(rec, list);
      },
    };
  }

  function nameProxy(rec: NameRec, list: NameRec[]) {
    return {
      get Name() {
        return rec.Name;
      },
      get RefersTo() {
        return rec.RefersTo;
      },
      set RefersTo(v: string) {
        rec.RefersTo = v;
      },
      get Visible() {
        return rec.Visible;
      },
      set Visible(v: boolean) {
        rec.Visible = v;
      },
      Delete() {
        // Only fail deletes for names that existed before the current operation
        // when nameDeleteFailFor is set; allows rollback Delete of newly-added names.
        if (
          options?.nameDeleteFailFor &&
          options.nameDeleteFailFor.some(
            (n) => n.localeCompare(rec.Name, undefined, { sensitivity: "accent" }) === 0,
          )
        ) {
          if (options.nameDeleteThrows) throw new Error(options.nameDeleteThrows);
          if (options.nameDeleteNoop) return;
        } else if (options?.nameDeleteThrows && !options.nameDeleteFailFor) {
          throw new Error(options.nameDeleteThrows);
        } else if (options?.nameDeleteNoop && !options.nameDeleteFailFor) {
          return;
        }
        const idx = list.indexOf(rec);
        if (idx >= 0) list.splice(idx, 1);
      },
    };
  }

  function sheetApi(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing ${name}`);
    const api: Record<string, unknown> = {
      get Name() {
        return sheet.Name;
      },
      get Index() {
        return sheet.Index;
      },
      Range(address: string) {
        const range: Record<string, unknown> = {
          Address: `${sheet.Name}!${address}`,
          Value2: [[1]],
          Formula: [["1"]],
        };
        if (withInsertDelete) {
          range.Insert = (shift: number) => {
            return { Address: `${sheet.Name}!INS:${address}:${shift}` };
          };
          range.Delete = (_shift: number) => {
            // no-op mock
          };
        }
        return range;
      },
      Delete() {
        sheets.delete(sheet.Name);
      },
      UsedRange: { Address: "A1:B2" },
    };
    if (withVisible) {
      Object.defineProperty(api, "Visible", {
        get() {
          return sheet.Visible;
        },
        set(v: number) {
          sheet.Visible = v;
        },
        enumerable: true,
        configurable: true,
      });
    }
    if (withProtect) {
      Object.defineProperty(api, "ProtectContents", {
        get() {
          return sheet.ProtectContents;
        },
        enumerable: true,
        configurable: true,
      });
      api.Protect = (password?: string) => {
        if (options?.protectThrows) throw new Error(options.protectThrows);
        sheet.ProtectContents = true;
        sheet.password = password;
      };
      api.Unprotect = (password?: string) => {
        if (sheet.password && password !== sheet.password) {
          throw new Error(`bad password ${password}`);
        }
        sheet.ProtectContents = false;
        sheet.password = undefined;
      };
    }
    if (withNames) {
      // worksheet-scoped names not stored separately in this mock
      api.Names = namesApi([]);
    }
    return api;
  }

  const workbook = {
    Name: "Book1.xlsx",
    get ActiveSheet() {
      return sheetApi("Sheet1");
    },
    Worksheets: {
      get Count() {
        return sheets.size;
      },
      Item(indexOrName: number | string) {
        if (typeof indexOrName === "number") {
          const name = [...sheets.keys()][indexOrName - 1];
          if (!name) throw new Error("index");
          return sheetApi(name);
        }
        return sheetApi(indexOrName);
      },
      Add() {
        const name = `Sheet${sheets.size + 1}`;
        sheets.set(name, {
          Name: name,
          Index: sheets.size + 1,
          Visible: -1,
          ProtectContents: false,
        });
        return sheetApi(name);
      },
    },
    ...(withNames ? { Names: namesApi(workbookNames) } : {}),
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Application: unknown }).Application = {
    Name: "WPS 表格",
    ActiveWorkbook: workbook,
  };

  return { sheets, workbookNames };
}

describe("WPS visibility mapping helpers", () => {
  it("maps COM Visible values", () => {
    expect(mapVisibleToPublic(-1)).toBe("visible");
    expect(mapVisibleToPublic(0)).toBe("hidden");
    expect(mapVisibleToPublic(2)).toBe("veryHidden");
    expect(mapPublicToVisible("hidden")).toBe(0);
    expect(mapPublicToVisible("veryHidden")).toBe(2);
  });
});

describe("WPS sheet visibility / protection / namedRange / insert-delete", () => {
  afterEach(() => {
    delete (globalThis as { Application?: unknown }).Application;
  });

  it("gets and sets visibility with writeback", async () => {
    installStructureWps();
    const adapter = new WpsJsaAdapter();
    const got = await adapter.getSheetVisibility("Sheet1");
    expect(got.ok).toBe(true);
    if (got.ok) expect(got.data.visibility).toBe("visible");

    const set = await adapter.setSheetVisibility("Sheet1", "veryHidden");
    expect(set.ok).toBe(true);
    if (set.ok) expect(set.data.visibility).toBe("veryHidden");

    const again = await adapter.getSheetVisibility("Sheet1");
    expect(again.ok).toBe(true);
    if (again.ok) expect(again.data.visibility).toBe("veryHidden");
  });

  it("returns typed unsupported when Visible is missing", async () => {
    installStructureWps({ withVisible: false });
    const adapter = new WpsJsaAdapter();
    const result = await adapter.getSheetVisibility("Sheet1");
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.unsupported).toBe(true);
  });

  it("rejects invalid visibility with fail", async () => {
    installStructureWps();
    const adapter = new WpsJsaAdapter();
    const result = await adapter.setSheetVisibility("Sheet1", "ghost" as never);
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.unsupported).not.toBe(true);
  });

  it("protects and unprotects with state verification", async () => {
    installStructureWps();
    const adapter = new WpsJsaAdapter();
    const before = await adapter.getSheetProtection("Sheet1");
    expect(before.ok).toBe(true);
    if (before.ok) expect(before.data.protected).toBe(false);

    const protectedResult = await adapter.protectSheet("Sheet1", "s3cret!");
    expect(protectedResult.ok).toBe(true);
    if (protectedResult.ok) {
      expect(protectedResult.data.protected).toBe(true);
      expect(JSON.stringify(protectedResult)).not.toContain("s3cret!");
    }

    const twice = await adapter.protectSheet("Sheet1");
    expect(twice.ok).toBe(false);
    if (!twice.ok) expect(twice.reason).toMatch(/already protected/i);

    const open = await adapter.unprotectSheet("Sheet1", "s3cret!");
    expect(open.ok).toBe(true);
    if (open.ok) {
      expect(open.data.protected).toBe(false);
      expect(JSON.stringify(open)).not.toContain("s3cret!");
    }
  });

  it("does not leak password when host throws", async () => {
    installStructureWps();
    const adapter = new WpsJsaAdapter();
    await adapter.protectSheet("Sheet1", "hunter2");
    const bad = await adapter.unprotectSheet("Sheet1", "wrong-pass");
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(bad.reason).not.toContain("wrong-pass");
      expect(bad.reason).not.toContain("hunter2");
      expect(JSON.stringify(bad)).not.toContain("wrong-pass");
      expect(JSON.stringify(bad)).not.toContain("hunter2");
    }
    expect(sanitizeProtectionMessage("fail hunter2", "hunter2")).toBe("fail [redacted]");
  });

  it("returns typed unsupported when Protect members are missing", async () => {
    installStructureWps({ withProtect: false });
    const adapter = new WpsJsaAdapter();
    const get = await adapter.getSheetProtection("Sheet1");
    expect(get.ok).toBe(false);
    if (!get.ok) expect(get.unsupported).toBe(true);
    const protect = await adapter.protectSheet("Sheet1");
    expect(protect.ok).toBe(false);
    if (!protect.ok) expect(protect.unsupported).toBe(true);
  });

  it("lists creates updates deletes named ranges with rename-add-first", async () => {
    const fake = installStructureWps();
    const adapter = new WpsJsaAdapter();
    const created = await adapter.createNamedRange({
      name: "Sales",
      refersTo: "Sheet1!$A$1",
      scope: "workbook",
      visible: true,
    });
    expect(created.ok).toBe(true);
    if (created.ok) {
      expect(created.data.name).toBe("Sales");
      expect(created.data.refersTo).toBe("=Sheet1!$A$1");
    }

    const listed = await adapter.listNamedRanges({ scope: "workbook" });
    expect(listed.ok).toBe(true);
    if (listed.ok) expect(listed.data.map((n) => n.name)).toContain("Sales");

    // conflict case-insensitive
    const conflict = await adapter.createNamedRange({
      name: "sales",
      refersTo: "Sheet1!$B$1",
      scope: "workbook",
    });
    expect(conflict.ok).toBe(false);

    const renamed = await adapter.updateNamedRange({
      name: "Sales",
      scope: "workbook",
      newName: "Revenue",
      refersTo: "Sheet1!$C$1",
    });
    expect(renamed.ok).toBe(true);
    if (renamed.ok) {
      expect(renamed.data.name).toBe("Revenue");
      expect(renamed.data.refersTo).toBe("=Sheet1!$C$1");
    }
    expect(fake.workbookNames.map((n) => n.Name)).toEqual(["Revenue"]);

    const deleted = await adapter.deleteNamedRange({
      name: "Revenue",
      scope: "workbook",
    });
    expect(deleted.ok).toBe(true);
    if (deleted.ok) expect(deleted.data.deleted).toBe("Revenue");
    expect(fake.workbookNames).toHaveLength(0);
  });

  it("keeps original name when rename target conflicts", async () => {
    installStructureWps();
    const adapter = new WpsJsaAdapter();
    await adapter.createNamedRange({
      name: "A",
      refersTo: "Sheet1!$A$1",
      scope: "workbook",
    });
    await adapter.createNamedRange({
      name: "B",
      refersTo: "Sheet1!$B$1",
      scope: "workbook",
    });
    const result = await adapter.updateNamedRange({
      name: "A",
      scope: "workbook",
      newName: "b",
    });
    expect(result.ok).toBe(false);
    const listed = await adapter.listNamedRanges({ scope: "workbook" });
    expect(listed.ok).toBe(true);
    if (listed.ok) {
      expect(listed.data.map((n) => n.name).sort()).toEqual(["A", "B"]);
    }
  });

  it("rolls back new name when original Delete fails on rename", async () => {
    installStructureWps({
      nameDeleteThrows: "delete denied",
      nameDeleteFailFor: ["OldName"],
    });
    const adapter = new WpsJsaAdapter();
    await adapter.createNamedRange({
      name: "OldName",
      refersTo: "=Sheet1!$A$1",
      scope: "workbook",
    });
    const renamed = await adapter.updateNamedRange({
      name: "OldName",
      newName: "NewName",
      scope: "workbook",
    });
    expect(renamed.ok).toBe(false);
    if (!renamed.ok) {
      expect(renamed.reason).toMatch(/rollback|delete/i);
    }
    const list = await adapter.listNamedRanges({ scope: "workbook" });
    expect(list.ok).toBe(true);
    if (list.ok) {
      const names = list.data.map((n) => n.name).sort();
      expect(names).toContain("OldName");
      expect(names).not.toContain("NewName");
    }
  });

  it("fails closed when original name still present after Delete no-op", async () => {
    installStructureWps({ nameDeleteNoop: true, nameDeleteFailFor: ["KeepMe"] });
    const adapter = new WpsJsaAdapter();
    await adapter.createNamedRange({
      name: "KeepMe",
      refersTo: "=Sheet1!$B$1",
      scope: "workbook",
    });
    const renamed = await adapter.updateNamedRange({
      name: "KeepMe",
      newName: "Ghost",
      scope: "workbook",
    });
    expect(renamed.ok).toBe(false);
    if (!renamed.ok) {
      expect(renamed.reason).toMatch(/incomplete|still present/i);
    }
    const list = await adapter.listNamedRanges({ scope: "workbook" });
    expect(list.ok).toBe(true);
    if (list.ok) {
      // Original retained; new name may also exist after incomplete Delete — surface failure, no silent dual-success
      expect(list.data.some((n) => n.name === "KeepMe")).toBe(true);
    }
  });

  it("returns typed unsupported when Names is missing", async () => {
    installStructureWps({ withNames: false });
    const adapter = new WpsJsaAdapter();
    const list = await adapter.listNamedRanges();
    expect(list.ok).toBe(false);
    if (!list.ok) expect(list.unsupported).toBe(true);
  });

  it("inserts and deletes range with mapped shift constants", async () => {
    installStructureWps();
    const adapter = new WpsJsaAdapter();
    const inserted = await adapter.insertRange({
      sheetName: "Sheet1",
      address: "A1:B2",
      shift: "down",
    });
    expect(inserted.ok).toBe(true);
    if (inserted.ok) {
      expect(inserted.data.operation).toBe("insert");
      expect(inserted.data.shift).toBe("down");
      expect(inserted.data.address).toContain("-4121");
    }

    const deleted = await adapter.deleteRange({
      sheetName: "Sheet1",
      address: "C3",
      shift: "left",
    });
    expect(deleted.ok).toBe(true);
    if (deleted.ok) {
      expect(deleted.data.operation).toBe("delete");
      expect(deleted.data.shift).toBe("left");
      expect(deleted.data.address).toContain("C3");
    }
  });

  it("rejects invalid shift and missing Insert/Delete members", async () => {
    installStructureWps();
    const adapter = new WpsJsaAdapter();
    const bad = await adapter.insertRange({
      sheetName: "Sheet1",
      address: "A1",
      shift: "up" as never,
    });
    expect(bad.ok).toBe(false);
    if (!bad.ok) expect(bad.unsupported).not.toBe(true);

    installStructureWps({ withInsertDelete: false });
    const missing = await new WpsJsaAdapter().insertRange({
      sheetName: "Sheet1",
      address: "A1",
      shift: "right",
    });
    expect(missing.ok).toBe(false);
    if (!missing.ok) {
      expect(missing.unsupported).toBe(true);
      expect(missing.reason).toMatch(/Insert/i);
    }
  });

  it("does not throw raw TypeError without Application", async () => {
    delete (globalThis as { Application?: unknown }).Application;
    const adapter = new WpsJsaAdapter();
    for (const result of [
      await adapter.getSheetVisibility("Sheet1"),
      await adapter.setSheetVisibility("Sheet1", "hidden"),
      await adapter.getSheetProtection("Sheet1"),
      await adapter.protectSheet("Sheet1", "x"),
      await adapter.unprotectSheet("Sheet1", "x"),
      await adapter.listNamedRanges(),
      await adapter.createNamedRange({
        name: "N",
        refersTo: "A1",
        scope: "workbook",
      }),
      await adapter.insertRange({
        sheetName: "Sheet1",
        address: "A1",
        shift: "down",
      }),
      await adapter.deleteRange({
        sheetName: "Sheet1",
        address: "A1",
        shift: "up",
      }),
    ]) {
      expect(result.ok).toBe(false);
      if (!result.ok) {
        expect(result.unsupported).toBe(true);
        expect(JSON.stringify(result)).not.toContain('"x"');
      }
    }
  });
});
