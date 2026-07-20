import { afterEach, describe, expect, it } from "vitest";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";
import { hexFromOleColor, oleColorFromHex } from "../shared/host/wpsJsaFormat";

type SheetState = {
  Name: string;
  Index: number;
  cells: Map<string, { values: unknown[][]; formulas: string[][] }>;
  format: {
    fontName: string | null;
    fontSize: number | null;
    fontBold: boolean | null;
    fontColor: number | null;
    fillColor: number | null;
    numberFormat: string | null;
    horizontalAlignment: number | null;
    verticalAlignment: number | null;
    wrapText: boolean | null;
    columnWidth: number | null;
    rowHeight: number | null;
  };
};

function blankFormat(): SheetState["format"] {
  return {
    fontName: null,
    fontSize: null,
    fontBold: null,
    fontColor: null,
    fillColor: null,
    numberFormat: "General",
    horizontalAlignment: 1,
    verticalAlignment: -4107,
    wrapText: false,
    columnWidth: 8.43,
    rowHeight: 15,
  };
}

function installRichWps(options?: {
  withCurrentRegion?: boolean;
  withCopyMove?: boolean;
  withFormat?: boolean;
  withAutofit?: boolean;
}) {
  const withCurrentRegion = options?.withCurrentRegion ?? true;
  const withCopyMove = options?.withCopyMove ?? true;
  const withFormat = options?.withFormat ?? true;
  const withAutofit = options?.withAutofit ?? true;

  const sheets = new Map<string, SheetState>();
  sheets.set("Sheet1", {
    Name: "Sheet1",
    Index: 1,
    cells: new Map([
      [
        "A1",
        {
          values: [[1]],
          formulas: [["1"]],
        },
      ],
      [
        "B2",
        {
          values: [
            [3, 4],
            [5, 6],
          ],
          formulas: [
            ["=1+2", ""],
            ["", "=6"],
          ],
        },
      ],
    ]),
    format: blankFormat(),
  });
  sheets.set("Sheet2", {
    Name: "Sheet2",
    Index: 2,
    cells: new Map(),
    format: blankFormat(),
  });

  let order = ["Sheet1", "Sheet2"];
  let activeName = "Sheet1";

  function reindex() {
    order.forEach((name, i) => {
      const sheet = sheets.get(name);
      if (sheet) sheet.Index = i + 1;
    });
  }

  function sheetApi(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing ${name}`);

    const api: Record<string, unknown> = {
      get Name() {
        return sheet.Name;
      },
      set Name(next: string) {
        const idx = order.indexOf(sheet.Name);
        sheets.delete(sheet.Name);
        sheet.Name = next;
        sheets.set(next, sheet);
        if (idx >= 0) order[idx] = next;
        if (activeName === name) activeName = next;
      },
      get Index() {
        return sheet.Index;
      },
      Range(address: string) {
        const key = address.includes("!") ? address.split("!")[1]! : address;
        const range: Record<string, unknown> = {
          get Address() {
            return `${sheet.Name}!${key}`;
          },
          get Value2() {
            return sheet.cells.get(key)?.values ?? [[null]];
          },
          set Value2(next: unknown[][]) {
            const prev = sheet.cells.get(key);
            sheet.cells.set(key, {
              values: next,
              formulas: prev?.formulas ?? next.map((row) => row.map(() => "")),
            });
          },
          get Formula() {
            const formulas = sheet.cells.get(key)?.formulas ?? [[""]];
            return formulas.length === 1 && formulas[0]?.length === 1
              ? formulas[0][0]
              : formulas;
          },
          set Formula(next: string | string[][]) {
            const matrix = typeof next === "string" ? [[next]] : next;
            const prev = sheet.cells.get(key);
            sheet.cells.set(key, {
              values: prev?.values ?? matrix.map((row) => row.map(() => null)),
              formulas: matrix,
            });
          },
          Clear() {
            sheet.cells.delete(key);
          },
        };

        if (withCurrentRegion) {
          Object.defineProperty(range, "CurrentRegion", {
            get() {
              return {
                Address: `${sheet.Name}!A1:C3`,
                Value2: [
                  [1, 2, 3],
                  [4, 5, 6],
                  [7, 8, 9],
                ],
                Formula: [
                  ["1", "2", "3"],
                  ["4", "5", "6"],
                  ["7", "8", "9"],
                ],
              };
            },
          });
        }

        if (withFormat) {
          Object.defineProperty(range, "Font", {
            get() {
              return {
                get Name() {
                  return sheet.format.fontName;
                },
                set Name(v: string | null) {
                  sheet.format.fontName = v;
                },
                get Size() {
                  return sheet.format.fontSize;
                },
                set Size(v: number | null) {
                  sheet.format.fontSize = v;
                },
                get Bold() {
                  return sheet.format.fontBold;
                },
                set Bold(v: boolean | null) {
                  sheet.format.fontBold = v;
                },
                get Color() {
                  return sheet.format.fontColor;
                },
                set Color(v: number | null) {
                  sheet.format.fontColor = v;
                },
              };
            },
          });
          Object.defineProperty(range, "Interior", {
            get() {
              return {
                get Color() {
                  return sheet.format.fillColor;
                },
                set Color(v: number | null) {
                  sheet.format.fillColor = v;
                },
              };
            },
          });
          Object.defineProperty(range, "NumberFormat", {
            get() {
              return sheet.format.numberFormat;
            },
            set(v: string) {
              sheet.format.numberFormat = v;
            },
          });
          Object.defineProperty(range, "HorizontalAlignment", {
            get() {
              return sheet.format.horizontalAlignment;
            },
            set(v: number) {
              sheet.format.horizontalAlignment = v;
            },
          });
          Object.defineProperty(range, "VerticalAlignment", {
            get() {
              return sheet.format.verticalAlignment;
            },
            set(v: number) {
              sheet.format.verticalAlignment = v;
            },
          });
          Object.defineProperty(range, "WrapText", {
            get() {
              return sheet.format.wrapText;
            },
            set(v: boolean) {
              sheet.format.wrapText = v;
            },
          });
        }

        if (withAutofit) {
          Object.defineProperty(range, "Columns", {
            get() {
              return {
                AutoFit() {
                  sheet.format.columnWidth = 12.5;
                },
              };
            },
          });
          Object.defineProperty(range, "Rows", {
            get() {
              return {
                AutoFit() {
                  sheet.format.rowHeight = 22;
                },
              };
            },
          });
          Object.defineProperty(range, "ColumnWidth", {
            get() {
              return sheet.format.columnWidth;
            },
          });
          Object.defineProperty(range, "RowHeight", {
            get() {
              return sheet.format.rowHeight;
            },
          });
        }

        return range;
      },
      Delete() {
        sheets.delete(sheet.Name);
        order = order.filter((n) => n !== sheet.Name);
        reindex();
        if (activeName === sheet.Name) activeName = order[0] ?? "";
      },
      UsedRange: { Address: "A1:B2" },
    };

    if (withCopyMove) {
      api.Copy = (_before?: unknown, _after?: unknown) => {
        const copyName = `${sheet.Name}_Copy`;
        const cloned: SheetState = {
          Name: copyName,
          Index: order.length + 1,
          cells: new Map(sheet.cells),
          format: { ...sheet.format },
        };
        sheets.set(copyName, cloned);
        order.push(copyName);
        reindex();
        activeName = copyName;
      };
      api.Move = (before?: { Name?: string }, after?: { Name?: string }) => {
        order = order.filter((n) => n !== sheet.Name);
        if (before?.Name) {
          const idx = order.indexOf(before.Name);
          order.splice(idx < 0 ? order.length : idx, 0, sheet.Name);
        } else if (after?.Name) {
          const idx = order.indexOf(after.Name);
          order.splice(idx < 0 ? order.length : idx + 1, 0, sheet.Name);
        } else {
          order.push(sheet.Name);
        }
        reindex();
      };
    }

    return api;
  }

  const workbook = {
    Name: "Book1.xlsx",
    get ActiveSheet() {
      return sheetApi(activeName || order[0]!);
    },
    Worksheets: {
      get Count() {
        return order.length;
      },
      Item(indexOrName: number | string) {
        if (typeof indexOrName === "number") {
          const name = order[indexOrName - 1];
          if (!name) throw new Error("index");
          return sheetApi(name);
        }
        return sheetApi(indexOrName);
      },
      Add() {
        const name = `Sheet${order.length + 1}`;
        sheets.set(name, {
          Name: name,
          Index: order.length + 1,
          cells: new Map(),
          format: blankFormat(),
        });
        order.push(name);
        reindex();
        return sheetApi(name);
      },
    },
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Application: unknown }).Application = {
    Name: "WPS 表格",
    ActiveWorkbook: workbook,
    Selection: (workbook.ActiveSheet as { Range: (a: string) => unknown }).Range("A1"),
  };

  return { sheets, order: () => [...order], active: () => activeName };
}

describe("wpsJsa color helpers", () => {
  it("round-trips OLE BGR and #RRGGBB", () => {
    expect(oleColorFromHex("#FF0000")).toBe(0x0000ff);
    expect(hexFromOleColor(0x0000ff)).toBe("#FF0000");
    expect(oleColorFromHex("not-a-color")).toBeNull();
  });
});

describe("WPS JSA currentRegion / sheet copy-move / format / autofit", () => {
  afterEach(() => {
    delete (globalThis as { Application?: unknown }).Application;
  });

  it("expands currentRegion when member is present", async () => {
    installRichWps();
    const adapter = new WpsJsaAdapter();
    const result = await adapter.readRange("Sheet1", "B2", "currentRegion");
    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.data.expandMode).toBe("currentRegion");
      expect(result.data.expanded).toBe(true);
      expect(result.data.address).toContain("A1:C3");
      expect(result.data.values).toEqual([
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
      ]);
    }
  });

  it("returns typed unsupported when CurrentRegion is missing", async () => {
    installRichWps({ withCurrentRegion: false });
    const adapter = new WpsJsaAdapter();
    const result = await adapter.readRange("Sheet1", "A1", "currentRegion");
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      expect(result.reason).toMatch(/CurrentRegion/i);
    }
  });

  it("keeps spill/currentArray typed unsupported", async () => {
    installRichWps();
    const adapter = new WpsJsaAdapter();
    for (const mode of ["spill", "currentArray"] as const) {
      const result = await adapter.readRange("Sheet1", "A1", mode);
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
    const defaultSpill = await adapter.readRange("Sheet1", "A1");
    expect(defaultSpill.ok).toBe(false);
  });

  it("copies a sheet after the last sheet and renames", async () => {
    const fake = installRichWps();
    const adapter = new WpsJsaAdapter();
    const copied = await adapter.copySheet("Sheet1", "Dup");
    expect(copied.ok).toBe(true);
    if (copied.ok) {
      expect(copied.data.name).toBe("Dup");
      expect(copied.data.index).toBe(3);
    }
    expect(fake.order()).toEqual(["Sheet1", "Sheet2", "Dup"]);
  });

  it("returns typed unsupported when Copy member is missing", async () => {
    installRichWps({ withCopyMove: false });
    const adapter = new WpsJsaAdapter();
    const result = await adapter.copySheet("Sheet1");
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      expect(result.reason).toMatch(/Copy/i);
    }
  });

  it("moves a sheet with 1-based position and rejects bad position", async () => {
    installRichWps();
    const adapter = new WpsJsaAdapter();
    const bad = await adapter.moveSheet("Sheet2", 0);
    expect(bad.ok).toBe(false);
    if (!bad.ok) {
      expect(bad.unsupported).not.toBe(true);
      expect(bad.reason).toMatch(/1-based/);
    }

    const moved = await adapter.moveSheet("Sheet2", 1);
    expect(moved.ok).toBe(true);
    if (moved.ok) expect(moved.data.name).toBe("Sheet2");
  });

  it("reads and writes basic format subset", async () => {
    installRichWps();
    const adapter = new WpsJsaAdapter();
    const written = await adapter.writeFormat("Sheet1", "A1", {
      fontBold: true,
      fontName: "微软雅黑",
      fontSize: 12,
      fontColor: "#FF0000",
      fillColor: "#00FF00",
      numberFormat: "0.00",
      horizontalAlignment: "center",
      verticalAlignment: "top",
      wrapText: true,
    });
    expect(written.ok).toBe(true);
    if (written.ok) {
      expect(written.data.format.fontBold).toBe(true);
      expect(written.data.format.fontName).toBe("微软雅黑");
      expect(written.data.format.fontSize).toBe(12);
      expect(written.data.format.fontColor).toBe("#FF0000");
      expect(written.data.format.fillColor).toBe("#00FF00");
      expect(written.data.format.numberFormat).toBe("0.00");
      expect(written.data.format.horizontalAlignment).toBe("center");
      expect(written.data.format.verticalAlignment).toBe("top");
      expect(written.data.format.wrapText).toBe(true);
    }

    const read = await adapter.readFormat("Sheet1", "A1");
    expect(read.ok).toBe(true);
    if (read.ok) expect(read.data.format.fontBold).toBe(true);
  });

  it("returns typed unsupported when format members are missing", async () => {
    installRichWps({ withFormat: false });
    const adapter = new WpsJsaAdapter();
    const read = await adapter.readFormat("Sheet1", "A1");
    expect(read.ok).toBe(false);
    if (!read.ok) expect(read.unsupported).toBe(true);
    const write = await adapter.writeFormat("Sheet1", "A1", { fontBold: true });
    expect(write.ok).toBe(false);
    if (!write.ok) expect(write.unsupported).toBe(true);
  });

  it("rejects invalid color and alignment with fail (not unsupported)", async () => {
    installRichWps();
    const adapter = new WpsJsaAdapter();
    const color = await adapter.writeFormat("Sheet1", "A1", { fillColor: "red" });
    expect(color.ok).toBe(false);
    if (!color.ok) expect(color.unsupported).not.toBe(true);

    const align = await adapter.writeFormat("Sheet1", "A1", {
      horizontalAlignment: "diagonal",
    });
    expect(align.ok).toBe(false);
    if (!align.ok) expect(align.unsupported).not.toBe(true);
  });

  it("autofits columns/rows and reads dimensions back", async () => {
    installRichWps();
    const adapter = new WpsJsaAdapter();
    const both = await adapter.autofitRange({
      sheetName: "Sheet1",
      address: "A1",
      direction: "both",
    });
    expect(both.ok).toBe(true);
    if (both.ok) {
      expect(both.data.columnWidth).toBe(12.5);
      expect(both.data.rowHeight).toBe(22);
      expect(both.data.direction).toBe("both");
    }
  });

  it("returns typed unsupported when AutoFit members are missing", async () => {
    installRichWps({ withAutofit: false });
    const adapter = new WpsJsaAdapter();
    const result = await adapter.autofitRange({
      sheetName: "Sheet1",
      address: "A1",
      direction: "rows",
    });
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.unsupported).toBe(true);
      expect(result.reason).toMatch(/AutoFit/i);
    }
  });

  it("does not throw raw TypeError when Application is absent", async () => {
    delete (globalThis as { Application?: unknown }).Application;
    const adapter = new WpsJsaAdapter();
    for (const result of [
      await adapter.readRange("Sheet1", "A1", "currentRegion"),
      await adapter.copySheet("Sheet1"),
      await adapter.moveSheet("Sheet1", 1),
      await adapter.readFormat("Sheet1", "A1"),
      await adapter.writeFormat("Sheet1", "A1", { fontBold: true }),
      await adapter.autofitRange({ sheetName: "Sheet1", address: "A1", direction: "both" }),
    ]) {
      expect(result.ok).toBe(false);
      if (!result.ok) expect(result.unsupported).toBe(true);
    }
  });
});
