/** @MOCK_INTERFACE — Excel.run formula protection double (formulas + locked + sheet protect). */

export type FormulaProtectionFakeOptions = {
  excelApi12?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  hasProtection?: boolean;
  hostSheetName?: string;
  /** 2D formulas; cells starting with = are formula cells */
  formulas?: string[][];
  /** initial locked matrix (same shape as formulas) */
  locked?: boolean[][];
  sheetProtected?: boolean;
  failVerifyProtect?: boolean;
};

export function installFormulaProtectionExcel(options: FormulaProtectionFakeOptions = {}) {
  const excelApi12 = options.excelApi12 !== false;
  const hasProtection = options.hasProtection !== false;
  const hostSheetName = options.hostSheetName ?? "HostSheet";
  const formulas = options.formulas ?? [
    ["H1", "H2"],
    ["=A1+1", "input"],
    ["=B2*2", "x"],
  ];
  let locked: boolean[][] =
    options.locked ??
    formulas.map((row) => row.map((cell) => (typeof cell === "string" && cell.startsWith("=") ? true : true)));
  let sheetProtected = options.sheetProtected === true;
  let protectCalls = 0;
  let unprotectCalls = 0;
  let lastPassword: string | undefined;
  const lockWrites: Array<{ r: number; c: number; locked: boolean }> = [];

  type CellProxy = {
    format: {
      protection: {
        locked: boolean;
        load: (p: string) => void;
        _flushLoad: () => void;
      };
    };
  };

  function makeCell(r: number, c: number): CellProxy {
    let snap: boolean | undefined;
    let pendingLoad: boolean | undefined;
    const protection = {
      get locked() {
        if (snap === undefined) throw new Error("protection.locked not loaded");
        return snap;
      },
      set locked(v: boolean) {
        if (!hasProtection) throw new Error("Range.format.protection.locked missing");
        lockWrites.push({ r, c, locked: v });
        if (!locked[r]) locked[r] = [];
        locked[r]![c] = v;
      },
      load(_p: string) {
        pendingLoad = locked[r]?.[c] === true;
        snap = undefined;
      },
      _flushLoad() {
        if (pendingLoad !== undefined) {
          snap = pendingLoad;
          pendingLoad = undefined;
        }
      },
    };
    return { format: { protection } };
  }

  let addressSnap: string | undefined;
  let addressPending: string | undefined;
  let formulasSnap: string[][] | undefined;
  let formulasPending: string[][] | undefined;
  let nameSnap: string | undefined;
  let namePending: string | undefined;
  let protectedSnap: boolean | undefined;
  let protectedPending: boolean | undefined;

  const range = {
    get address() {
      if (addressSnap === undefined) throw new Error("address not loaded");
      return addressSnap;
    },
    get formulas() {
      if (formulasSnap === undefined) throw new Error("formulas not loaded");
      return formulasSnap;
    },
    get rowCount() {
      return formulas.length;
    },
    get columnCount() {
      return formulas[0]?.length ?? 0;
    },
    format: {
      protection: {
        get locked() {
          return false;
        },
        set locked(v: boolean) {
          if (!hasProtection) throw new Error("Range.format.protection.locked missing");
          for (let r = 0; r < locked.length; r++) {
            for (let c = 0; c < (locked[r]?.length ?? 0); c++) {
              locked[r]![c] = v;
              lockWrites.push({ r, c, locked: v });
            }
          }
        },
        load() {},
      },
    },
    load(props: string) {
      if (props.includes("address")) {
        addressPending = `${hostSheetName}!A1:B3`;
        addressSnap = undefined;
      }
      if (props.includes("formulas")) {
        formulasPending = formulas.map((row) => [...row]);
        formulasSnap = undefined;
      }
    },
    getCell(r: number, c: number) {
      return makeCell(r, c);
    },
    _flushLoad() {
      if (addressPending !== undefined) {
        addressSnap = addressPending;
        addressPending = undefined;
      }
      if (formulasPending !== undefined) {
        formulasSnap = formulasPending;
        formulasPending = undefined;
      }
    },
    _flushCells() {
      // no-op; cells flush via shared locked state on next load
    },
  };

  // Track live cell proxies for flush after sync
  const liveCells: CellProxy[] = [];
  const origGetCell = range.getCell.bind(range);
  range.getCell = (r: number, c: number) => {
    const cell = origGetCell(r, c);
    liveCells.push(cell);
    return cell;
  };

  const sheet = {
    get name() {
      if (nameSnap === undefined) throw new Error("Worksheet.name not loaded");
      return nameSnap;
    },
    load(props: string) {
      if (props.includes("name")) {
        namePending = hostSheetName;
        nameSnap = undefined;
      }
    },
    protection: {
      get protected() {
        if (protectedSnap === undefined) throw new Error("protection.protected not loaded");
        return protectedSnap;
      },
      load(_p: string) {
        protectedPending = sheetProtected;
        protectedSnap = undefined;
      },
      protect(_opts?: object, password?: string) {
        protectCalls += 1;
        lastPassword = password;
        if (options.failVerifyProtect) {
          sheetProtected = false;
        } else {
          sheetProtected = true;
        }
      },
      unprotect(password?: string) {
        unprotectCalls += 1;
        lastPassword = password;
        sheetProtected = false;
      },
      _flushLoad() {
        if (protectedPending !== undefined) {
          protectedSnap = protectedPending;
          protectedPending = undefined;
        }
      },
    },
    getRange(_addr: string) {
      return range;
    },
    getUsedRangeOrNullObject() {
      return {
        isNullObject: false,
        address: `${hostSheetName}!A1:B3`,
        load() {},
      };
    },
    _flushLoad() {
      if (namePending !== undefined) {
        nameSnap = namePending;
        namePending = undefined;
      }
      sheet.protection._flushLoad();
      range._flushLoad();
      for (const cell of liveCells) cell.format.protection._flushLoad();
    },
  };

  async function sync() {
    // Flush pending loads first (cells registered during this batch), then drop list.
    sheet._flushLoad();
    liveCells.length = 0;
  }

  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as unknown as { Excel: { run: Function } }).Excel = {
    run: async <T>(fn: (ctx: unknown) => Promise<T>) =>
      fn({
        workbook: {
          worksheets: {
            getItem(_name: string) {
              return sheet;
            },
            items: [sheet],
            load() {},
          },
        },
        sync,
      }),
  };

  if (options.missingIsSetSupported) {
    (globalThis as { Office?: unknown }).Office = { context: { requirements: {} } };
  } else if (options.isSetSupportedThrows) {
    (globalThis as { Office?: unknown }).Office = {
      context: {
        requirements: {
          isSetSupported() {
            throw new Error("boom");
          },
        },
      },
    };
  } else {
    (globalThis as { Office?: unknown }).Office = {
      context: {
        requirements: {
          isSetSupported(_name: string, version?: string) {
            if (version === "1.2") return excelApi12;
            return false;
          },
        },
      },
    };
  }

  return {
    protectCalls: () => protectCalls,
    unprotectCalls: () => unprotectCalls,
    lastPassword: () => lastPassword,
    locked: () => locked.map((row) => [...row]),
    sheetProtected: () => sheetProtected,
    lockWrites: () => [...lockWrites],
    setSheetProtected(v: boolean) {
      sheetProtected = v;
    },
    /** test helper: must not appear in tool results */
    secretPasswordSeen: () => lastPassword,
  };
}
