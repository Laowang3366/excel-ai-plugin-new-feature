/**
 * Office.js formula protection manage (lock/unlock formula cells).
 */
import type {
  FormulaProtectionInspectInput,
  FormulaProtectionManageInfo,
  FormulaProtectionManageInput,
} from "./formulaProtectionTypes";
import type { ExcelWorksheet } from "./officeJsRuntime";
import { withExcel } from "./officeJsRuntime";
import type { HostResult } from "./types";
import { unsupported } from "./types";
import {
  inspectAll,
  isExcelApi12Supported,
  LOCKED_EVIDENCE,
  requireScope,
  setFormulaLocks,
  verifyManage,
} from "./officeJsFormulaProtectionHelpers";

export async function officeJsManageFormulaProtection(
  input: FormulaProtectionManageInput,
): Promise<HostResult<FormulaProtectionManageInfo>> {
  if (!isExcelApi12Supported()) {
    return unsupported(
      "formula.protection.manage",
      "office-js",
      "ExcelApi 1.2 is not supported in this host (Office.context.requirements.isSetSupported)",
      LOCKED_EVIDENCE,
    );
  }
  requireScope(input.scope, input.sheetName, input.range);
  if (input.command !== "lock" && input.command !== "unlock") {
    throw new Error("command must be lock|unlock");
  }

  // Password must never appear in returned data / error detail payloads we construct.
  const password =
    input.password != null && input.password !== "" ? input.password : undefined;
  const unlockInputs = input.unlockInputs !== false;
  const protectSheet = input.command === "lock" ? input.protectSheet !== false : false;

  return withExcel("formula.protection.manage", async (context) => {
    const limitations: string[] = [];
    const inspectInput: FormulaProtectionInspectInput = {
      scope: input.scope,
      sheetName: input.sheetName,
      range: input.range,
    };
    const before = await inspectAll(context, inspectInput);

    const sheetsToTouch: ExcelWorksheet[] = [];
    if (input.scope === "workbook") {
      context.workbook.worksheets.load("items/name");
      await context.sync();
      sheetsToTouch.push(...context.workbook.worksheets.items);
    } else {
      sheetsToTouch.push(context.workbook.worksheets.getItem(input.sheetName!));
    }

    for (const sheet of sheetsToTouch) {
      sheet.load("name");
      sheet.protection.load("protected");
      await context.sync();
      if (sheet.protection.protected === true) {
        try {
          if (password != null) sheet.protection.unprotect(password);
          else sheet.protection.unprotect();
          sheet.protection.load("protected");
          await context.sync();
        } catch (error) {
          const message = error instanceof Error ? error.message : String(error);
          throw new Error(
            `cannot unprotect sheet before formula lock changes (password not echoed): ${message}`,
          );
        }
      }

      await setFormulaLocks(
        context,
        sheet,
        input.scope === "workbook" ? undefined : input.range,
        input.command === "lock",
        input.command === "lock" ? unlockInputs : false,
        limitations,
      );

      if (protectSheet) {
        if (password != null) sheet.protection.protect({}, password);
        else sheet.protection.protect();
        sheet.protection.load("protected");
        await context.sync();
        if (sheet.protection.protected !== true) {
          throw new Error("protectSheet failed: sheet.protection.protected is not true after protect");
        }
      }
    }

    const after = await inspectAll(context, inspectInput);
    const verification = verifyManage(
      input.command,
      unlockInputs,
      protectSheet,
      before,
      after,
    );
    limitations.push(...verification.limitations);

    if (!verification.verified) {
      throw new Error(
        `formula protection write-back verification failed: ${verification.limitations.join("; ")}`,
      );
    }

    return {
      command: input.command,
      scope: input.scope,
      unlockInputs: input.command === "lock" ? unlockInputs : false,
      protectSheet,
      protection: after,
      verified: true,
      limitations,
    };
  });
}
