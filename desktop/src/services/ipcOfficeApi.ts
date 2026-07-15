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
      automation: {
        documents: {
          list: async (app) => getRaw()?.office.automation?.documents.list(app) ?? ipcUnavailable(),
          activate: async (input) =>
            getRaw()?.office.automation?.documents.activate(input) ?? ipcUnavailable(),
        },
        objects: {
          list: async (input) =>
            getRaw()?.office.automation?.objects.list(input) ?? ipcUnavailable(),
          activate: async (input) =>
            getRaw()?.office.automation?.objects.activate(input) ?? ipcUnavailable(),
        },
        workflows: {
          list: async () => getRaw()?.office.automation?.workflows.list() ?? ipcUnavailable(),
          get: async (id) => getRaw()?.office.automation?.workflows.get(id) ?? ipcUnavailable(),
          cancel: async (id) =>
            getRaw()?.office.automation?.workflows.cancel(id) ?? ipcUnavailable(),
          resume: async (id) =>
            getRaw()?.office.automation?.workflows.resume(id) ?? ipcUnavailable(),
        },
        templates: {
          list: async () => getRaw()?.office.automation?.templates.list() ?? ipcUnavailable(),
          saveFromWorkflow: async (input) =>
            getRaw()?.office.automation?.templates.saveFromWorkflow(input) ?? ipcUnavailable(),
          delete: async (id) =>
            getRaw()?.office.automation?.templates.delete(id) ?? ipcUnavailable(),
          run: async (input) =>
            getRaw()?.office.automation?.templates.run(input) ?? ipcUnavailable(),
        },
        transactions: {
          list: async () => getRaw()?.office.automation?.transactions.list() ?? ipcUnavailable(),
          get: async (id) => getRaw()?.office.automation?.transactions.get(id) ?? ipcUnavailable(),
          undo: async (id, force) =>
            getRaw()?.office.automation?.transactions.undo(id, force) ?? ipcUnavailable(),
          redo: async (id, force) =>
            getRaw()?.office.automation?.transactions.redo(id, force) ?? ipcUnavailable(),
        },
      },
    },
  };
}

function ipcUnavailable() {
  return Promise.resolve({ success: false as const, error: "Office 自动化 IPC 不可用" });
}
