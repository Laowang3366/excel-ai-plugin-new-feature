import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  disconnectOfficeBridges,
  getExcelBridge,
  getOrCreateExcelBridge,
  getOrCreateOfficeBridges,
  getVbaBridge,
  resetOfficeBridgeRegistry,
  setExcelBridgeInstance,
} from "./bridgeRegistry";

describe("bridgeRegistry lifecycle", () => {
  beforeEach(() => {
    resetOfficeBridgeRegistry();
  });

  it("clears cached bridge instances when reset", () => {
    const first = getOrCreateOfficeBridges();
    expect(getExcelBridge()).toBe(first.excelBridge);
    expect(getVbaBridge()).toBe(first.vbaBridge);

    resetOfficeBridgeRegistry();

    expect(getExcelBridge()).toBeNull();
    expect(getVbaBridge()).toBeNull();

    const second = getOrCreateOfficeBridges();
    expect(second.excelBridge).not.toBe(first.excelBridge);
    expect(second.vbaBridge).not.toBe(first.vbaBridge);
  });

  it("clears dependent bridges when replacing the Excel bridge", () => {
    const first = getOrCreateOfficeBridges();
    setExcelBridgeInstance(first.excelBridge);

    expect(getVbaBridge()).toBeNull();

    resetOfficeBridgeRegistry();
  });

  it("creates the Excel bridge without eagerly creating dependent bridges", () => {
    const bridge = getOrCreateExcelBridge();

    expect(getExcelBridge()).toBe(bridge);
    expect(getVbaBridge()).toBeNull();
  });

  it("disconnects Excel and releases owned Word and PowerPoint apps", async () => {
    const bridges = getOrCreateOfficeBridges();
    const disconnectExcel = vi.spyOn(bridges.excelBridge, "disconnect").mockResolvedValue(undefined);
    const saveWord = vi.spyOn(bridges.wordBridge, "saveDocument").mockResolvedValue({ success: true });
    const savePresentation = vi.spyOn(bridges.presentationBridge, "savePresentation").mockResolvedValue({ success: true });

    await disconnectOfficeBridges();

    expect(disconnectExcel).toHaveBeenCalled();
    expect(saveWord).toHaveBeenCalled();
    expect(savePresentation).toHaveBeenCalled();
  });
});
