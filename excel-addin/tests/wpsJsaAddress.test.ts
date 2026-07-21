import { afterEach, describe, expect, it } from "vitest";
import {
  hasWpsAddressSurface,
  normalizeWpsA1Address,
  readWpsAddress,
} from "../shared/host/wpsJsaAddress";
import { WpsJsaAdapter } from "../shared/host/wpsJsaAdapter";

afterEach(() => {
  delete (globalThis as { Application?: unknown }).Application;
  delete (globalThis as { wps?: unknown }).wps;
  delete (globalThis as { Wps?: unknown }).Wps;
});

function installSelectionApp(selection: Record<string, unknown>) {
  (globalThis as unknown as { window: unknown }).window = globalThis;
  (globalThis as { Application?: unknown }).Application = {
    Name: "WPS 表格",
    ActiveWorkbook: {
      Name: "工作簿1",
      ActiveSheet: { Name: "Sheet1" },
    },
    Selection: selection,
  };
}

describe("readWpsAddress", () => {
  it("reads non-empty string property", () => {
    expect(readWpsAddress({ Address: "G17" })).toBe("G17");
    expect(readWpsAddress({ Address: "  $A$1:$B$2  " })).toBe("$A$1:$B$2");
  });

  it("calls zero-arg Address method with correct this", () => {
    const owner = {
      sheet: "Sheet1",
      Address() {
        // real WPS binds this to the range/selection owner
        return `${(this as { sheet: string }).sheet}!G17`;
      },
    };
    expect(readWpsAddress(owner)).toBe("Sheet1!G17");
  });

  it("never returns Function.toString() source", () => {
    const owner = {
      Address() {
        return "G17";
      },
    };
    const addr = readWpsAddress(owner);
    expect(addr).toBe("G17");
    expect(addr).not.toMatch(/^function\b/);
    // raw String(fn) regression guard
    // native String(method) leaks source (classic or method shorthand)
    expect(String(owner.Address)).toMatch(/Address\s*\(|^function\b/);
  });

  it("uses fallback when method returns function/object/empty/non-string", () => {
    expect(readWpsAddress({ Address: () => ownerFn }, "A1")).toBe("A1");
    function ownerFn() {
      return "nope";
    }
    expect(readWpsAddress({ Address: () => ({ a: 1 }) }, "B2")).toBe("B2");
    expect(readWpsAddress({ Address: () => "" }, "C3")).toBe("C3");
    expect(readWpsAddress({ Address: () => 12 as unknown as string }, "D4")).toBe("D4");
    expect(readWpsAddress({ Address: () => null as unknown as string }, "E5")).toBe("E5");
  });

  it("uses fallback when getter/method throws or member missing", () => {
    expect(
      readWpsAddress(
        {
          get Address(): string {
            throw new Error("boom");
          },
        },
        "F6",
      ),
    ).toBe("F6");
    expect(
      readWpsAddress(
        {
          Address() {
            throw new Error("call boom");
          },
        },
        "G7",
      ),
    ).toBe("G7");
    expect(readWpsAddress({}, "H8")).toBe("H8");
    expect(readWpsAddress(null, "I9")).toBe("I9");
    expect(readWpsAddress(undefined, "J10")).toBe("J10");
  });

  it("without fallback returns undefined on failure (no fake address)", () => {
    expect(readWpsAddress({ Address: () => "" })).toBeUndefined();
    expect(readWpsAddress({ Address: () => ({}) })).toBeUndefined();
    expect(
      readWpsAddress({
        Address() {
          throw new Error("x");
        },
      }),
    ).toBeUndefined();
    expect(readWpsAddress({ Address: "function Address() { return 'G17'; }" })).toBeUndefined();
  });

  it("hasWpsAddressSurface detects string and method forms", () => {
    expect(hasWpsAddressSurface({ Address: "A1" })).toBe(true);
    expect(hasWpsAddressSurface({ Address: () => "A1" })).toBe(true);
    expect(hasWpsAddressSurface({ Address: "" })).toBe(false);
    expect(hasWpsAddressSurface({})).toBe(false);
  });
});

describe("normalizeWpsA1Address", () => {
  it("strips absolute $ only from A1 refs; keeps sheet qualifiers quote-aware", () => {
    expect(normalizeWpsA1Address("$G$17")).toBe("G17");
    expect(normalizeWpsA1Address("$A$1:$B$2")).toBe("A1:B2");
    expect(normalizeWpsA1Address("Sheet1!$G$17")).toBe("Sheet1!G17");
    expect(normalizeWpsA1Address("Sheet$1!$A$1:$B$2")).toBe("Sheet$1!A1:B2");
    expect(normalizeWpsA1Address("'Budget$2026'!$A$1")).toBe("'Budget$2026'!A1");
    expect(normalizeWpsA1Address("'A!B'!$A$1:$C$2")).toBe("'A!B'!A1:C2");
    expect(normalizeWpsA1Address("'A,!$''B'!$A$1:$C$2")).toBe("'A,!$''B'!A1:C2");
    expect(normalizeWpsA1Address("Sheet$1!$A$1,Sheet$1!$C$3")).toBe("Sheet$1!A1,Sheet$1!C3");
    expect(normalizeWpsA1Address("G17")).toBe("G17");
  });
});

describe("selection.get WPS Address method (host evidence shape)", () => {
  it("returns G17 when Address is a method, not function source", async () => {
    const selection = {
      Address() {
        return "G17";
      },
      Value2: [[null]],
      Formula: [[""]],
      Worksheet: { Name: "Sheet1" },
    };
    installSelectionApp(selection);
    const adapter = new WpsJsaAdapter();
    const result = await adapter.getSelection();
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.data.sheetName).toBe("Sheet1");
    expect(result.data.address).toBe("G17");
    expect(result.data.address).not.toMatch(/function/i);
    expect(result.data.values).toEqual([[null]]);
    expect(result.data.formulas).toEqual([[""]]);
  });

  it("normalizes real-host absolute Address method $G$17 to G17", async () => {
    installSelectionApp({
      Address() {
        return "$G$17";
      },
      Value2: [[null]],
      Formula: [[""]],
      Worksheet: { Name: "Sheet1" },
    });
    const result = await new WpsJsaAdapter().getSelection();
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.data.address).toBe("G17");
    expect(result.data.address).not.toContain("$");
    expect(result.data.sheetName).toBe("Sheet1");
    expect(result.data.values).toEqual([[null]]);
  });

  it("normalizes absolute Address string property $G$17 to G17", async () => {
    installSelectionApp({
      Address: "$G$17",
      Value2: [[null]],
      Formula: [[""]],
      Worksheet: { Name: "Sheet1" },
    });
    const result = await new WpsJsaAdapter().getSelection();
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.data.address).toBe("G17");
  });

  it("returns unsupported when Address method yields unusable value and no fallback", async () => {
    installSelectionApp({
      Address() {
        return "";
      },
      Value2: [[1]],
      Formula: [["1"]],
      Worksheet: { Name: "Sheet1" },
    });
    const result = await new WpsJsaAdapter().getSelection();
    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.unsupported).toBe(true);
    expect(result.capability).toBe("selection.get");
  });

  it("string Address property still works", async () => {
    installSelectionApp({
      Address: "B5",
      Value2: [[2]],
      Formula: [["2"]],
      Worksheet: { Name: "Sheet1" },
    });
    const result = await new WpsJsaAdapter().getSelection();
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.data.address).toBe("B5");
  });
});
