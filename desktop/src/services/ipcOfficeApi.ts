import type { IIpcApi } from "./ipcApiTypes";

type RawIpcApiGetter = () => IIpcApi | null;

export function createOfficeIpcApi(getRaw: RawIpcApiGetter): Pick<IIpcApi, "excel" | "office"> {
  return {
    excel: {
      detectStatus: async () => {
        const raw = getRaw();
        if (!raw) return { connected: false, host: "" };
        return raw.excel.detectStatus();
      },
      connect: async () => {
        const raw = getRaw();
        if (!raw) return { connected: false, host: "" };
        return raw.excel.connect();
      },
      selectHost: async (host) => {
        const raw = getRaw();
        if (!raw) return { connected: false, host: "" };
        return raw.excel.selectHost(host);
      },
      getSelection: async () => {
        const raw = getRaw();
        if (!raw) return { address: "", values: [], sheetName: "" };
        return raw.excel.getSelection();
      },
      getSelectionAddress: async () => {
        const raw = getRaw();
        if (!raw) return { address: "", sheetName: "" };
        if (raw.excel.getSelectionAddress) return raw.excel.getSelectionAddress();
        const selection = await raw.excel.getSelection();
        return { address: selection.address, sheetName: selection.sheetName };
      },
      readRange: async (sheetName, range, expand) => {
        const raw = getRaw();
        if (!raw) return { values: [] };
        return raw.excel.readRange(sheetName, range, expand);
      },
      inspectWorkbook: async () => {
        const raw = getRaw();
        if (!raw) return null;
        return raw.excel.inspectWorkbook();
      },
      writeRange: async (sheetName, range, values) => {
        const raw = getRaw();
        if (!raw) return { success: false, error: "IPC not available" };
        return raw.excel.writeRange(sheetName, range, values);
      },
    },

    office: {
      detectWordStatus: async () => {
        const raw = getRaw();
        if (!raw) return { connected: false, host: "unknown" };
        return raw.office.detectWordStatus();
      },
      detectPresentationStatus: async () => {
        const raw = getRaw();
        if (!raw) return { connected: false, host: "unknown" };
        return raw.office.detectPresentationStatus();
      },
    },
  };
}
