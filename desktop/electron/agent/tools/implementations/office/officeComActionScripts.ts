import type { OfficeActionInput } from "../../officeCore/types";
import { buildCrossOfficeScript, isCrossOfficeOperation } from "./officeComCrossActionScripts";
import { buildExcelScript } from "./officeComExcelActionScripts";
import { buildPresentationScript } from "./officeComPresentationActionScripts";
import { buildWordScript } from "./officeComWordActionScripts";

export function buildComScript(input: OfficeActionInput): string {
  if (isCrossOfficeOperation(input.operation)) return buildCrossOfficeScript(input);
  if (input.app === "excel") return buildExcelScript(input);
  if (input.app === "word") return buildWordScript(input);
  return buildPresentationScript(input);
}
