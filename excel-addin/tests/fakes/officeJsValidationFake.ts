/**
 * Sync-gated CF/DV fake (ExcelApi 1.6 / 1.8 shapes).
 * - Mutations queue until context.sync(); loads see committed state only.
 * - cellValue.rule: whole-object assign; custom.rule: ClientObject formula property.
 * - list.source accepts string or Range-like { address }.
 * - Office.context.requirements.isSetSupported injectable.
 */
import { makeConditionalFormatsApi } from "./officeJsValidationFakeCf";
import { makeDvProxy } from "./officeJsValidationFakeDv";
import {
  keyOf,
  seedContainsText,
  seedInconsistentDv,
  seedManyCf,
  type CfState,
  type CtxPending,
  type DvState,
} from "./officeJsValidationFakeState";

export function installValidationExcel(options?: {
  excelApi16?: boolean;
  excelApi18?: boolean;
  missingIsSetSupported?: boolean;
  isSetSupportedThrows?: boolean;
  failAddSync?: boolean;
  failDeleteReadback?: boolean;
  failWriteSync?: boolean;
  failClearReadback?: boolean;
  seedContainsText?: boolean;
  seedInconsistentDv?: boolean;
  /** Seed many CF rules for O(1) list sync tests. */
  seedManyCf?: number;
}) {
  const excelApi16 = options?.excelApi16 !== false;
  const excelApi18 = options?.excelApi18 !== false;

  const cfs = new Map<string, CfState[]>();
  const dvs = new Map<string, DvState>();
  let seq = 0;
  let syncCount = 0;

  if (options?.seedContainsText) seedContainsText(cfs);
  if (options?.seedManyCf && options.seedManyCf > 0) seedManyCf(cfs, options.seedManyCf);
  if (options?.seedInconsistentDv) seedInconsistentDv(dvs);

  function makeContext() {
    const pending: CtxPending = {
      cfAdds: [],
      cfDeletes: [],
      cfPatches: [],
      dvWrites: new Map(),
      loads: [],
    };

    function makeRange(sheetName: string, address: string) {
      const key = keyOf(sheetName, address);
      let addressValue = `${sheetName}!${address}`;
      let itemProxies: ReturnType<
        ReturnType<typeof makeConditionalFormatsApi>["getItem"]
      >[] = [];

      const dvProxy = makeDvProxy(key, dvs, pending);
      const conditionalFormats = makeConditionalFormatsApi({
        key,
        cfs,
        pending,
        nextSeq: () => {
          seq += 1;
          return seq;
        },
        getItemProxies: () => itemProxies,
        setItemProxies: (items) => {
          itemProxies = items;
        },
      });

      return {
        get address() {
          return addressValue;
        },
        load(props: string) {
          if (props.includes("address")) {
            pending.loads.push(() => {
              addressValue = `${sheetName}!${address}`;
            });
          }
        },
        conditionalFormats,
        dataValidation: dvProxy,
        getRange(inner: string) {
          return { address: `${sheetName}!${inner}` };
        },
      };
    }

    return {
      workbook: {
        worksheets: {
          getItem(name: string) {
            return {
              getRange(address: string) {
                return makeRange(name, address);
              },
            };
          },
        },
      },
      async sync() {
        syncCount += 1;
        if (options?.failAddSync && pending.cfAdds.length > 0) {
          throw new Error("sync failed on add");
        }
        if (options?.failWriteSync && pending.dvWrites.size > 0) {
          const onlyClear = [...pending.dvWrites.values()].every((v) => v === "clear");
          if (!onlyClear) throw new Error("sync failed on write");
        }
        for (const load of pending.loads) load();
        pending.loads.length = 0;

        for (const add of pending.cfAdds) {
          const list = cfs.get(add.key) ?? [];
          list.push({ ...add.state });
          cfs.set(add.key, list);
        }
        for (const patch of pending.cfPatches) {
          const list = cfs.get(patch.key) ?? [];
          const hit = list.find((s) => s.id === patch.id);
          if (hit) Object.assign(hit, patch.patch);
        }
        for (const del of pending.cfDeletes) {
          if (options?.failDeleteReadback) continue;
          cfs.set(
            del.key,
            (cfs.get(del.key) ?? []).filter((s) => s.id !== del.id),
          );
        }
        for (const [k, v] of pending.dvWrites) {
          if (v === "clear") {
            if (!options?.failClearReadback) dvs.delete(k);
          } else {
            dvs.set(k, v);
          }
        }
        pending.cfAdds.length = 0;
        pending.cfDeletes.length = 0;
        pending.cfPatches.length = 0;
        pending.dvWrites.clear();
      },
    };
  }

  const g = globalThis as unknown as {
    window: unknown;
    Excel: { run: Function };
    Office?: {
      context?: {
        requirements?: { isSetSupported?: (name: string, minVersion?: string) => boolean };
      };
    };
  };
  g.window = globalThis;
  if (options?.missingIsSetSupported) {
    g.Office = { context: { requirements: {} } };
  } else if (options?.isSetSupportedThrows) {
    g.Office = {
      context: {
        requirements: {
          isSetSupported: () => {
            throw new Error("isSetSupported boom");
          },
        },
      },
    };
  } else {
    g.Office = {
      context: {
        requirements: {
          isSetSupported: (_name: string, minVersion?: string) => {
            if (minVersion === "1.6") return excelApi16;
            if (minVersion === "1.8") return excelApi18;
            return false;
          },
        },
      },
    };
  }
  g.Excel = {
    run: async <T>(fn: (ctx: ReturnType<typeof makeContext>) => Promise<T>) =>
      fn(makeContext()),
  };

  return {
    getSyncCount: () => syncCount,
    resetSyncCount: () => {
      syncCount = 0;
    },
  };
}
