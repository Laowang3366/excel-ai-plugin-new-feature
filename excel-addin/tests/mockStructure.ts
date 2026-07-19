import { toA1 } from "../shared/host/a1Address";
import { ok } from "../shared/host/types";
import type {
  NamedRangeInfo,
  NamedRangeScope,
  SheetDisplayInfo,
  SheetDisplayUpdateInput,
  SheetPageLayoutInfo,
  SheetPageLayoutUpdateInput,
  SheetVisibility,
} from "../shared/host/types";

export function createMockStructureState() {
  const visibility = new Map<string, SheetVisibility>();
  const protectedSheets = new Set<string>();
  let namedRanges: NamedRangeInfo[] = [];
  const display = new Map<string, SheetDisplayInfo>();
  const freeze = new Map<
    string,
    { sheetName: string; address: string | null; rowCount: number; columnCount: number }
  >();
  const pageLayouts = new Map<string, SheetPageLayoutInfo>();

  function defaultDisplay(sheetName: string): SheetDisplayInfo {
    return (
      display.get(sheetName) ?? {
        sheetName,
        tabColor: "",
        showGridlines: true,
        showHeadings: true,
      }
    );
  }

  function defaultPageLayout(sheetName: string): SheetPageLayoutInfo {
    return (
      pageLayouts.get(sheetName) ?? {
        sheetName,
        orientation: "portrait",
        centerHorizontally: false,
        centerVertically: false,
        printGridlines: false,
        printHeadings: false,
        blackAndWhite: false,
        draft: false,
        pageOrder: "downThenOver",
        firstPageNumber: null,
        margins: { top: 72, bottom: 72, left: 72, right: 72 },
        zoomScale: 100,
        paperSize: "letter",
        fitToPagesWide: null,
        fitToPagesTall: null,
        printArea: null,
        printTitleRows: null,
        printTitleColumns: null,
      }
    );
  }

  return {
    async getSheetVisibility(sheetName: string) {
      return ok({
        sheetName,
        visibility: visibility.get(sheetName) ?? "visible",
      });
    },
    async setSheetVisibility(sheetName: string, next: SheetVisibility) {
      visibility.set(sheetName, next);
      return ok({ sheetName, visibility: next });
    },
    async getSheetProtection(sheetName: string) {
      return ok({ sheetName, protected: protectedSheets.has(sheetName) });
    },
    async protectSheet(sheetName: string, _password?: string) {
      if (protectedSheets.has(sheetName)) throw new Error("sheet is already protected");
      protectedSheets.add(sheetName);
      return ok({ sheetName, protected: true });
    },
    async unprotectSheet(sheetName: string, _password?: string) {
      if (!protectedSheets.has(sheetName)) throw new Error("sheet is not protected");
      protectedSheets.delete(sheetName);
      return ok({ sheetName, protected: false });
    },
    async listNamedRanges(input?: { scope?: NamedRangeScope; sheetName?: string }) {
      let list = namedRanges;
      if (input?.scope) list = list.filter((item) => item.scope === input.scope);
      if (input?.sheetName) list = list.filter((item) => item.sheetName === input.sheetName);
      return ok(list);
    },
    async createNamedRange(input: {
      name: string;
      refersTo: string;
      scope: NamedRangeScope;
      sheetName?: string;
      visible?: boolean;
    }) {
      const info: NamedRangeInfo = {
        name: input.name,
        refersTo: input.refersTo.startsWith("=") ? input.refersTo : `=${input.refersTo}`,
        scope: input.scope,
        sheetName: input.sheetName,
        visible: input.visible ?? true,
      };
      namedRanges.push(info);
      return ok(info);
    },
    async updateNamedRange(input: {
      name: string;
      scope: NamedRangeScope;
      sheetName?: string;
      newName?: string;
      refersTo?: string;
      visible?: boolean;
    }) {
      const item = namedRanges.find(
        (entry) =>
          entry.name === input.name &&
          entry.scope === input.scope &&
          (input.sheetName == null || entry.sheetName === input.sheetName),
      );
      if (!item) throw new Error("named range not found");
      if (input.newName) item.name = input.newName;
      if (input.refersTo) {
        item.refersTo = input.refersTo.startsWith("=") ? input.refersTo : `=${input.refersTo}`;
      }
      if (input.visible != null) item.visible = input.visible;
      return ok(item);
    },
    async deleteNamedRange(input: {
      name: string;
      scope: NamedRangeScope;
      sheetName?: string;
    }) {
      namedRanges = namedRanges.filter(
        (entry) =>
          !(
            entry.name === input.name &&
            entry.scope === input.scope &&
            (input.sheetName == null || entry.sheetName === input.sheetName)
          ),
      );
      return ok({ deleted: input.name });
    },
    async getSheetDisplay(sheetName: string) {
      return ok(defaultDisplay(sheetName));
    },
    async setSheetDisplay(input: SheetDisplayUpdateInput) {
      const current = defaultDisplay(input.sheetName);
      const next: SheetDisplayInfo = {
        sheetName: input.sheetName,
        tabColor: input.tabColor !== undefined ? input.tabColor : current.tabColor,
        showGridlines:
          input.showGridlines !== undefined ? input.showGridlines : current.showGridlines,
        showHeadings:
          input.showHeadings !== undefined ? input.showHeadings : current.showHeadings,
      };
      display.set(input.sheetName, next);
      return ok(next);
    },
    async getSheetFreeze(sheetName: string) {
      return ok(freeze.get(sheetName) ?? { sheetName, address: null, rowCount: 0, columnCount: 0 });
    },
    async setSheetFreeze(input: {
      sheetName: string;
      command: "rows" | "columns" | "at" | "clear";
      count?: number;
      address?: string;
    }) {
      if (input.command === "clear") {
        const cleared = { sheetName: input.sheetName, address: null, rowCount: 0, columnCount: 0 };
        freeze.set(input.sheetName, cleared);
        return ok(cleared);
      }
      if (input.command === "rows") {
        const next = {
          sheetName: input.sheetName,
          address: `A${(input.count ?? 1) + 1}`,
          rowCount: input.count ?? 1,
          columnCount: 0,
        };
        freeze.set(input.sheetName, next);
        return ok(next);
      }
      if (input.command === "columns") {
        const count = input.count ?? 1;
        const next = {
          sheetName: input.sheetName,
          address: toA1(0, count - 1),
          rowCount: 0,
          columnCount: count,
        };
        freeze.set(input.sheetName, next);
        return ok(next);
      }
      const next = {
        sheetName: input.sheetName,
        address: input.address ?? "B2",
        rowCount: 1,
        columnCount: 1,
      };
      freeze.set(input.sheetName, next);
      return ok(next);
    },
    async getSheetPageLayout(sheetName: string) {
      return ok(defaultPageLayout(sheetName));
    },
    async setSheetPageLayout(input: SheetPageLayoutUpdateInput) {
      const current = defaultPageLayout(input.sheetName);
      const next: SheetPageLayoutInfo = {
        ...current,
        sheetName: input.sheetName,
        orientation: input.orientation ?? current.orientation,
        centerHorizontally: input.centerHorizontally ?? current.centerHorizontally,
        centerVertically: input.centerVertically ?? current.centerVertically,
        printGridlines: input.printGridlines ?? current.printGridlines,
        printHeadings: input.printHeadings ?? current.printHeadings,
        blackAndWhite: input.blackAndWhite ?? current.blackAndWhite,
        draft: input.draft !== undefined ? input.draft : current.draft,
        pageOrder: input.pageOrder ?? current.pageOrder,
        firstPageNumber:
          input.firstPageNumber !== undefined ? input.firstPageNumber : current.firstPageNumber,
        margins: {
          top: input.margins?.top ?? current.margins.top,
          bottom: input.margins?.bottom ?? current.margins.bottom,
          left: input.margins?.left ?? current.margins.left,
          right: input.margins?.right ?? current.margins.right,
        },
        zoomScale:
          input.fitToPagesWide !== undefined || input.fitToPagesTall !== undefined
            ? null
            : (input.zoomScale ?? current.zoomScale),
        paperSize: input.paperSize ?? current.paperSize,
        fitToPagesWide:
          input.zoomScale !== undefined
            ? null
            : (input.fitToPagesWide ?? current.fitToPagesWide),
        fitToPagesTall:
          input.zoomScale !== undefined
            ? null
            : (input.fitToPagesTall ?? current.fitToPagesTall),
        printArea: input.printArea ?? current.printArea,
        printTitleRows: input.printTitleRows ?? current.printTitleRows,
        printTitleColumns: input.printTitleColumns ?? current.printTitleColumns,
      };
      pageLayouts.set(input.sheetName, next);
      return ok(next);
    },
  };
}
