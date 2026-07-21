import {
  getSheet,
  requireWorkbook,
  type WpsRange,
} from "./wpsJsaRuntime";
import { readWpsAddress } from "./wpsJsaAddress";
import type { HostResult, RangeFormat, RangeFormatData } from "./types";
import { fail, ok, unsupported } from "./types";

/**
 * Assumed ET Range format members (COM parity with desktop ExcelActionService),
 * not present in the in-repo JSA bridge contract and not device-verified.
 */
const EVIDENCE =
  "Assumed Range.Font/Interior/NumberFormat/WrapText/HorizontalAlignment/VerticalAlignment (desktop COM parity; not in bridge contract; not device-verified)";

/** Excel/WPS COM HorizontalAlignment constants. */
const H_ALIGN_TO_COM: Record<string, number> = {
  general: 1,
  left: -4131,
  center: -4108,
  right: -4152,
  fill: 5,
  justify: -4130,
};
const H_ALIGN_FROM_COM: Record<number, string> = {
  1: "general",
  [-4131]: "left",
  [-4108]: "center",
  [-4152]: "right",
  5: "fill",
  [-4130]: "justify",
};

/** Excel/WPS COM VerticalAlignment constants. */
const V_ALIGN_TO_COM: Record<string, number> = {
  top: -4160,
  center: -4108,
  bottom: -4107,
  justify: -4130,
};
const V_ALIGN_FROM_COM: Record<number, string> = {
  [-4160]: "top",
  [-4108]: "center",
  [-4107]: "bottom",
  [-4130]: "justify",
};

function messageOf(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function asNullableBoolean(value: unknown): boolean | null {
  if (typeof value === "boolean") return value;
  if (value === -1 || value === 1) return true;
  if (value === 0) return false;
  if (value == null) return null;
  return null;
}


/** #RRGGBB → OLE BGR int (desktop ExcelActionService.OleColor parity). */
export function oleColorFromHex(hex: string): number | null {
  const value = hex.trim().replace(/^#/, "");
  if (!/^[0-9A-Fa-f]{6}$/.test(value)) return null;
  const rgb = Number.parseInt(value, 16);
  const red = (rgb >> 16) & 255;
  const green = (rgb >> 8) & 255;
  const blue = rgb & 255;
  return red | (green << 8) | (blue << 16);
}

/** OLE BGR int → #RRGGBB. */
export function hexFromOleColor(value: unknown): string | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  const n = value >>> 0;
  const red = n & 255;
  const green = (n >> 8) & 255;
  const blue = (n >> 16) & 255;
  const hex = ((red << 16) | (green << 8) | blue).toString(16).padStart(6, "0");
  return `#${hex.toUpperCase()}`;
}

function readNumberFormat(value: unknown): string | null {
  if (value == null) return null;
  if (typeof value === "string") return value;
  if (Array.isArray(value)) {
    const first = value[0];
    if (Array.isArray(first)) return first[0] == null ? null : String(first[0]);
    return first == null ? null : String(first);
  }
  return String(value);
}

function mapHAlign(value: unknown): string | null {
  if (typeof value === "string") return value.toLowerCase();
  if (typeof value === "number" && value in H_ALIGN_FROM_COM) return H_ALIGN_FROM_COM[value]!;
  return value == null ? null : String(value);
}

function mapVAlign(value: unknown): string | null {
  if (typeof value === "string") return value.toLowerCase();
  if (typeof value === "number" && value in V_ALIGN_FROM_COM) return V_ALIGN_FROM_COM[value]!;
  return value == null ? null : String(value);
}

function resolveRange(
  capability: string,
  sheetName: string,
  address: string,
): HostResult<{ range: WpsRange; sheetName: string; address: string }> {
  const workbookResult = requireWorkbook(capability);
  if (!workbookResult.ok) return workbookResult;
  const sheet = getSheet(workbookResult.data, sheetName);
  if (!sheet?.Range) {
    return unsupported(
      capability,
      "wps-jsa",
      `Sheet "${sheetName}" or Range API missing`,
      EVIDENCE,
    );
  }
  try {
    const range = sheet.Range(address);
    return ok({
      range,
      sheetName,
      address: readWpsAddress(range, address) ?? address,
    });
  } catch (error) {
    return fail(capability, "wps-jsa", messageOf(error), EVIDENCE);
  }
}

function readFormatSnapshot(range: WpsRange): RangeFormat {
  const font = range.Font;
  const interior = range.Interior;
  return {
    fontName: font && "Name" in font ? (font.Name ?? null) : null,
    fontSize:
      font && typeof font.Size === "number" && Number.isFinite(font.Size) ? font.Size : null,
    fontBold: font ? asNullableBoolean(font.Bold) : null,
    fontColor: font ? hexFromOleColor(font.Color) : null,
    fillColor: interior ? hexFromOleColor(interior.Color) : null,
    numberFormat: readNumberFormat(range.NumberFormat),
    horizontalAlignment: mapHAlign(range.HorizontalAlignment),
    verticalAlignment: mapVAlign(range.VerticalAlignment),
    wrapText: asNullableBoolean(range.WrapText),
  };
}

function hasAnyFormatSurface(range: WpsRange): boolean {
  return Boolean(
    range.Font ||
      range.Interior ||
      range.NumberFormat !== undefined ||
      range.HorizontalAlignment !== undefined ||
      range.VerticalAlignment !== undefined ||
      range.WrapText !== undefined,
  );
}

export async function wpsReadFormat(
  sheetName: string,
  address: string,
): Promise<HostResult<RangeFormatData>> {
  const resolved = resolveRange("range.format.read", sheetName, address);
  if (!resolved.ok) return resolved;
  const { range } = resolved.data;
  if (!hasAnyFormatSurface(range)) {
    return unsupported(
      "range.format.read",
      "wps-jsa",
      "Range format members are unavailable",
      EVIDENCE,
    );
  }
  try {
    return ok({
      sheetName,
      address: resolved.data.address,
      format: readFormatSnapshot(range),
    });
  } catch (error) {
    return fail("range.format.read", "wps-jsa", messageOf(error), EVIDENCE);
  }
}

export async function wpsWriteFormat(
  sheetName: string,
  address: string,
  format: RangeFormat,
): Promise<HostResult<RangeFormatData>> {
  const resolved = resolveRange("range.format.write", sheetName, address);
  if (!resolved.ok) return resolved;
  const { range } = resolved.data;
  if (!hasAnyFormatSurface(range)) {
    return unsupported(
      "range.format.write",
      "wps-jsa",
      "Range format members are unavailable",
      EVIDENCE,
    );
  }

  try {
    if (format.fontName != null || format.fontSize != null || format.fontBold != null || format.fontColor != null) {
      if (!range.Font) {
        return unsupported(
          "range.format.write",
          "wps-jsa",
          "Range.Font is unavailable",
          EVIDENCE,
        );
      }
      if (format.fontName != null) range.Font.Name = format.fontName;
      if (format.fontSize != null) range.Font.Size = format.fontSize;
      if (format.fontBold != null) range.Font.Bold = format.fontBold;
      if (format.fontColor != null) {
        const ole = oleColorFromHex(format.fontColor);
        if (ole == null) {
          return fail(
            "range.format.write",
            "wps-jsa",
            `fontColor must be #RRGGBB, got "${format.fontColor}"`,
            EVIDENCE,
          );
        }
        range.Font.Color = ole;
      }
    }

    if (format.fillColor != null) {
      if (!range.Interior) {
        return unsupported(
          "range.format.write",
          "wps-jsa",
          "Range.Interior is unavailable",
          EVIDENCE,
        );
      }
      const ole = oleColorFromHex(format.fillColor);
      if (ole == null) {
        return fail(
          "range.format.write",
          "wps-jsa",
          `fillColor must be #RRGGBB, got "${format.fillColor}"`,
          EVIDENCE,
        );
      }
      range.Interior.Color = ole;
    }

    if (format.numberFormat != null) {
      if (range.NumberFormat === undefined && !("NumberFormat" in range)) {
        return unsupported(
          "range.format.write",
          "wps-jsa",
          "Range.NumberFormat is unavailable",
          EVIDENCE,
        );
      }
      range.NumberFormat = format.numberFormat;
    }

    if (format.horizontalAlignment != null) {
      const key = String(format.horizontalAlignment).toLowerCase();
      const com = H_ALIGN_TO_COM[key];
      if (com === undefined) {
        return fail(
          "range.format.write",
          "wps-jsa",
          `unsupported horizontalAlignment "${format.horizontalAlignment}"`,
          EVIDENCE,
        );
      }
      if (range.HorizontalAlignment === undefined && !("HorizontalAlignment" in range)) {
        return unsupported(
          "range.format.write",
          "wps-jsa",
          "Range.HorizontalAlignment is unavailable",
          EVIDENCE,
        );
      }
      range.HorizontalAlignment = com;
    }

    if (format.verticalAlignment != null) {
      const key = String(format.verticalAlignment).toLowerCase();
      const com = V_ALIGN_TO_COM[key];
      if (com === undefined) {
        return fail(
          "range.format.write",
          "wps-jsa",
          `unsupported verticalAlignment "${format.verticalAlignment}"`,
          EVIDENCE,
        );
      }
      if (range.VerticalAlignment === undefined && !("VerticalAlignment" in range)) {
        return unsupported(
          "range.format.write",
          "wps-jsa",
          "Range.VerticalAlignment is unavailable",
          EVIDENCE,
        );
      }
      range.VerticalAlignment = com;
    }

    if (format.wrapText != null) {
      if (range.WrapText === undefined && !("WrapText" in range)) {
        return unsupported(
          "range.format.write",
          "wps-jsa",
          "Range.WrapText is unavailable",
          EVIDENCE,
        );
      }
      range.WrapText = format.wrapText;
    }

    return ok({
      sheetName,
      address: readWpsAddress(range, address) ?? address,
      format: readFormatSnapshot(range),
    });
  } catch (error) {
    return fail("range.format.write", "wps-jsa", messageOf(error), EVIDENCE);
  }
}
