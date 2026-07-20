/**
 * Sync-gated CF/DV fake (ExcelApi 1.6 / 1.8).
 * - add() proxy: id/type PropertyNotLoaded until load+sync
 * - list.source keeps Range-like object by default
 * - ignoreBlanks is sync-gated with rule
 */
import { makeConditionalFormatsApi } from "./officeJsValidationFakeCf";
import { makeDvProxy, type DvWriteCounts } from "./officeJsValidationFakeDv";
import {
  keyOf,
  seedContainsText,
  seedInconsistentDv,
  seedManyCf,
  type CfState,
  type CtxPending,
  type DvState,
  type ValidationFakeOptions,
} from "./officeJsValidationFakeState";

export type { ValidationFakeOptions };

export function installValidationExcel(options?: ValidationFakeOptions) {
  const excelApi16 = options?.excelApi16 !== false;
  const excelApi18 = options?.excelApi18 !== false;
  const keepListSourceAsRangeObject = options?.keepListSourceAsRangeObject !== false;

  const cfs = new Map<string, CfState[]>();
  const dvs = new Map<string, DvState>();
  let seq = 0;
  let syncCount = 0;
  const writeCounts: DvWriteCounts = { rule: 0, ignoreBlanks: 0, errorAlert: 0, prompt: 0 };
  const loadPropsLog: string[] = [];

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

      const dvProxy = makeDvProxy(key, dvs, pending, {
        keepListSourceAsRangeObject,
        getTamper: () => options?.tamperDvReadback,
        clearLeavesHostType: options?.clearLeavesHostType,
        missingDvErrorAlert: options?.missingDvErrorAlert,
        missingDvPrompt: options?.missingDvPrompt,
        missingDvErrorAlertFields: options?.missingDvErrorAlertFields,
        writeCounts,
        recordLoadProps: (props) => {
          loadPropsLog.push(props);
        },
      });
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
        getTamper: () => options?.tamperCfReadback,
      });

      let addressLoaded = false;
      const committedAddress = `${sheetName}!${address}`;
      addressValue = committedAddress;
      return {
        __rangeAddress: committedAddress,
        get address() {
          if (!addressLoaded) throw new Error("PropertyNotLoaded:address");
          return addressValue;
        },
        load(props: string) {
          if (props.includes("address")) {
            pending.loads.push(() => {
              addressValue = committedAddress;
              addressLoaded = true;
            });
          }
        },
        conditionalFormats,
        dataValidation: dvProxy,
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

        // Commit mutations first, then apply loads (Office.js: sync flushes then resolves loads).
        for (const add of pending.cfAdds) {
          const list = cfs.get(add.key) ?? [];
          list.push({ ...add.state });
          cfs.set(add.key, list);
        }
        for (const patch of pending.cfPatches) {
          const list = cfs.get(patch.key) ?? [];
          const hit = list.find((s) => s.id === patch.id);
          if (hit) Object.assign(hit, patch.patch);
          // Also patch in-memory add state object (same reference when add not yet listed).
          for (const add of pending.cfAdds) {
            if (add.key === patch.key && add.state.id === patch.id) {
              Object.assign(add.state, patch.patch);
            }
          }
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
            // failClearReadback: leave previous DV so clear appears incomplete.
            // clearLeavesHostType: delete so load synthesizes non-None hostType.
            if (!options?.failClearReadback) {
              dvs.delete(k);
            }
          } else {
            const prev = dvs.get(k);
            dvs.set(k, {
              type: v.type,
              ignoreBlanks: v.ignoreBlanks,
              rule: v.rule,
              // Omit-preserving: only replace metadata when the write queue marked it.
              errorAlert: "errorAlert" in v ? v.errorAlert : prev?.errorAlert,
              prompt: "prompt" in v ? v.prompt : prev?.prompt,
            });
          }
        }
        pending.cfAdds.length = 0;
        pending.cfDeletes.length = 0;
        pending.cfPatches.length = 0;
        pending.dvWrites.clear();

        for (const load of pending.loads) load();
        pending.loads.length = 0;
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
    getDvWriteCounts: () => ({ ...writeCounts }),
    resetDvWriteCounts: () => {
      writeCounts.rule = 0;
      writeCounts.ignoreBlanks = 0;
      writeCounts.errorAlert = 0;
      writeCounts.prompt = 0;
    },
    getLoadPropsLog: () => [...loadPropsLog],
    resetLoadPropsLog: () => {
      loadPropsLog.length = 0;
    },
  };
}
