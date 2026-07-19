/** Fake Excel.run for Worksheet.tabColor / showGridlines / showHeadings. */
export function installDisplayExcel() {
  const sheets = new Map<
    string,
    {
      name: string;
      tabColor: string;
      showGridlines: boolean;
      showHeadings: boolean;
    }
  >();
  sheets.set("Sheet1", {
    name: "Sheet1",
    tabColor: "",
    showGridlines: true,
    showHeadings: true,
  });

  function makeSheet(name: string) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`missing sheet ${name}`);
    return {
      get name() {
        return sheet.name;
      },
      get tabColor() {
        return sheet.tabColor;
      },
      set tabColor(next: string) {
        sheet.tabColor = next;
      },
      get showGridlines() {
        return sheet.showGridlines;
      },
      set showGridlines(next: boolean) {
        sheet.showGridlines = next;
      },
      get showHeadings() {
        return sheet.showHeadings;
      },
      set showHeadings(next: boolean) {
        sheet.showHeadings = next;
      },
      load() {},
    };
  }

  const context = {
    workbook: {
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
    getState(name: string) {
      return sheets.get(name);
    },
  };
}
