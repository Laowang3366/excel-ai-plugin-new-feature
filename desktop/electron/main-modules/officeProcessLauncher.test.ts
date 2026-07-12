import { describe, expect, it } from "vitest";

import { getOfficeExecutableName, parseRegistryDefaultValue } from "./officeProcessLauncher";

describe("parseRegistryDefaultValue", () => {
  it("reads an App Paths default executable with spaces", () => {
    expect(
      parseRegistryDefaultValue(`
HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\excel.exe
    (Default)    REG_SZ    C:\\Program Files\\Microsoft Office\\Root\\Office16\\EXCEL.EXE
`),
    ).toBe("C:\\Program Files\\Microsoft Office\\Root\\Office16\\EXCEL.EXE");
  });

  it("returns undefined when the registry value is absent", () => {
    expect(parseRegistryDefaultValue("ERROR: The system was unable to find the key")).toBeUndefined();
  });

  it("launches the WPS spreadsheet process used by the connection bridge", () => {
    expect(getOfficeExecutableName("wps")).toBe("et.exe");
    expect(getOfficeExecutableName("excel")).toBe("excel.exe");
  });
});
