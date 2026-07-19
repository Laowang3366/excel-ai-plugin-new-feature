/** Fake Excel.run matching official protect(options?, password?) and readonly NamedItem.name. */
export function installStructureExcel() {
  type Named = { name: string; formula: string; visible: boolean };

  const sheets = new Map<
    string,
    {
      name: string;
      visibility: string;
      protected: boolean;
      names: Named[];
    }
  >();
  sheets.set("Sheet1", {
    name: "Sheet1",
    visibility: "Visible",
    protected: false,
    names: [],
  });
  const workbookNames: Named[] = [];
  let lastProtectCall: { options: unknown; password: unknown } | null = null;

  function makeNames(list: Named[]) {
    return {
      get items() {
        return list.map((item) => makeNamed(item, list));
      },
      load() {},
      add(name: string, formula: string) {
        // Approximate Excel rules: reject empty, pure cell-like A1/AB12, and reserved R1C1.
        if (
          name.trim() === "" ||
          /^[A-Za-z]{1,3}\d{1,7}$/.test(name) ||
          /^R\d+C\d+$/i.test(name)
        ) {
          throw new Error(`invalid named range name: ${name}`);
        }
        const clash = list.some(
          (entry) => entry.name.localeCompare(name, undefined, { sensitivity: "accent" }) === 0,
        );
        if (clash) throw new Error(`named range already exists: ${name}`);
        const item = { name, formula, visible: true };
        list.push(item);
        return makeNamed(item, list);
      },
      getItem(name: string) {
        const found = list.find((item) => item.name === name);
        if (!found) throw new Error(`missing name ${name}`);
        return makeNamed(found, list);
      },
    };
  }

  function makeNamed(state: Named, list: Named[]) {
    return {
      get name() {
        return state.name;
      },
      // name is readonly in Office.js — no setter
      get formula() {
        return state.formula;
      },
      set formula(next: string) {
        state.formula = next;
      },
      get visible() {
        return state.visible;
      },
      set visible(next: boolean) {
        state.visible = next;
      },
      load() {},
      delete() {
        const idx = list.indexOf(state);
        if (idx >= 0) list.splice(idx, 1);
      },
    };
  }

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    return {
      get name() {
        return sheet.name;
      },
      get visibility() {
        return sheet.visibility;
      },
      set visibility(next: string) {
        sheet.visibility = next;
      },
      protection: {
        get protected() {
          return sheet.protected;
        },
        load() {},
        protect(options?: object, password?: string) {
          lastProtectCall = { options, password };
          // Do not persist password in sheet state
          sheet.protected = true;
        },
        unprotect(password?: string) {
          void password;
          sheet.protected = false;
        },
      },
      names: makeNames(sheet.names),
      load() {},
    };
  }

  const context = {
    workbook: {
      names: makeNames(workbookNames),
      worksheets: {
        getItem(name: string) {
          return makeSheet(name);
        },
      },
    },
    async sync() {},
  };

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: typeof context) => Promise<T>) => fn(context),
  };

  return {
    getVisibility(name: string) {
      return sheets.get(name)?.visibility;
    },
    isProtected(name: string) {
      return sheets.get(name)?.protected ?? false;
    },
    getLastProtectCall() {
      return lastProtectCall;
    },
    getSheetState(name: string) {
      return sheets.get(name);
    },
    workbookNames,
  };
}
